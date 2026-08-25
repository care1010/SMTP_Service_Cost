const db = require('../config/db');
const xlsx = require('xlsx');
const pgFormat = require('pg-format');
const { triggerAutoSync } = require('./cronController');
const mailService = require('../services/mailService');

const formatExcelDate = (excelDate) => {
    if (!excelDate) return null;
    if (typeof excelDate === 'number') {
        const date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
        return date.toISOString().split('T')[0];
    }
    return excelDate;
};

exports.uploadPtdData = async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ message: "No file uploaded" });
        
        console.log("🚀 PTD UPLOAD: Started processing...");
        const workbook = xlsx.readFile(req.file.path, { cellDates: true });
        const sheetNames = workbook.SheetNames;
        
        let affectedLoas = new Set();
        let detectedPeriods = new Set(); 
        const batchSize = 500;

        // 1. CJI5 Processing
        if (sheetNames.includes('CJI5')) {
            const cji5Data = xlsx.utils.sheet_to_json(workbook.Sheets['CJI5']);
            await db.query("TRUNCATE TABLE cji5_new");
            const cji5Rows = cji5Data.filter(row => row['WBS Element']).map(row => {
                if (row['LOA_ID']) affectedLoas.add(row['LOA_ID'].toString().trim());
                if (row['Per']) detectedPeriods.add(`P${row['Per'].toString().trim().padStart(3, '0')}`);
                return [
                    row['Project Def.'], row['WBS Element'], row['RefDocNo'], row['Item'],
                    row['CO object name'], row['Supplier'], row['Name'], row['Year'],
                    row['Per'], row['Cost elem.'], row['Cost element descr.'], row['Matl Group'],
                    row['Material'], row['Description'], row['User Name'], row['DocC'],
                    row['CoCode'], row['Exch. Rate'], row['Quantity'], row['Qty/plan'],
                    formatExcelDate(row['Debit date']), formatExcelDate(row['Doc. Date']),
                    row['Report currency'], row['Val.in rep.cur.'], row['TCurr'], row['Value TCur'], 
                    row['Obj Curr.'], row['Value in Obj. Crcy']
                ];
            });
            for (let i = 0; i < cji5Rows.length; i += batchSize) {
                const sql = pgFormat(`INSERT INTO cji5_new (project_def, wbs_element, refdocno, item, co_object_name, supplier, name, year, per, cost_element, cost_element_descr, matl_group, material, description, user_name, docc, cocode, exch_rate, quantity, qty_plan, debit_date, doc_date, report_currency, val_in_rep_cur, tcurr, value_tcur, obj_curr, value_in_obj_crcy) VALUES %L`, cji5Rows.slice(i, i + batchSize));
                await db.query(sql); 
            }
            console.log(`✅ CJI5 uploaded.`);
        }

        // 2. CJ74 Processing
        if (sheetNames.includes('CJ74')) {
            const cj74Data = xlsx.utils.sheet_to_json(workbook.Sheets['CJ74']);
            const cj74Rows = cj74Data.filter(row => row['Object']).map(row => {
                if (row['LOA_ID']) affectedLoas.add(row['LOA_ID'].toString().trim());
                if (row['Per']) detectedPeriods.add(`P${row['Per'].toString().trim().padStart(3, '0')}`);
                return [
                    row['CoCd'], row['Year'], row['Per'], row['Project def.'], 
                    row['Object'], row['Object'], row['Object'], row['Profit Ctr'], 
                    row['Cost Element'], row['Cost element name'], row['Cost element descr.'], 
                    row['Pur. Doc.'], row['Purchase order text'], row['DocumentNo'], 
                    row['Material'], row['Material Description'], row['Name'], row['RefDocNo'], 
                    row['frm'], row['User Name'], row['Offst.acct'], row['Name of offsetting account'], 
                    row['Quantity'], formatExcelDate(row['Created on']), formatExcelDate(row['Postg Date']), 
                    formatExcelDate(row['Doc. Date']), row['TCurr'], row['Value TranCurr'], 
                    row['ObCur'], row['Value in Obj. Crcy'], row['RCurr'], row['Val.in RC']
                ];
            });
            for (let i = 0; i < cj74Rows.length; i += batchSize) {
                const sql = pgFormat(`INSERT INTO cj74_new (cocd, year, per, proj_def, object_1, object_2, object_3, profit_ctr, cost_element, cost_element_name, cost_element_descr, pur_doc, purchase_order_text, document_no, material, material_description, name1, refdocno, frm, user_name, offst_acct, name_of_offsetting_account, quantity, created_on, postg_date, doc_date, tcurr, value_trancurr, obcur, val_in_obj_crcy, rcurr, val_in_rc) VALUES %L`, cj74Rows.slice(i, i + batchSize));
                await db.query(sql);
                console.log(`📦 CJ74: Batch ${Math.floor(i/batchSize) + 1} Done`);
            }
        }

        // 3. Dashboard Sync
        const loaList = Array.from(affectedLoas).filter(id => id);
        if (loaList.length > 0) {
            const syncSql = pgFormat(`
                UPDATE final_dashboard_table f
                SET ptd = COALESCE(src.ptd_sum, 0),
                    "open_commitment_KEUR" = COALESCE(src.oc_sum, 0),
                    eac = (COALESCE(src.ptd_sum, 0) + COALESCE(src.oc_sum, 0) + f.non_committed_editable),
                    eac_vs_asbl = (f.asbl - (COALESCE(src.ptd_sum, 0) + COALESCE(src.oc_sum, 0) + f.non_committed_editable))
                FROM (
                    SELECT f_inner.loa_id, f_inner.categories, SUM(cj.ptd_val) as ptd_sum, SUM(ci.open_commitment_keur) as oc_sum
                    FROM final_dashboard_table f_inner
                    LEFT JOIN v_cj74_transformed cj ON f_inner.loa_id = cj.loa_id AND f_inner.categories = cj.categories
                    LEFT JOIN v_cji5_transformed ci ON f_inner.loa_id = ci.loa_id AND f_inner.categories = ci.categories
                    WHERE f_inner.loa_id IN (%L)
                    GROUP BY f_inner.loa_id, f_inner.categories
                ) src
                WHERE f.loa_id = src.loa_id AND f.categories = src.categories
            `, loaList);
            await db.query(syncSql);
            console.log("✅ Dashboard Sync Done");
        }

        // 4. 🔥 FORCE TRIGGER MAIL (At the very end)
        let periodArray = Array.from(detectedPeriods);
        
        // --- SAFE FALLBACK: Agar Excel mein 'Per' nahi mila, toh Current Month use karo ---
        if (periodArray.length === 0) {
            const curMonthNum = new Date().getMonth() + 1; // August = 8
            periodArray.push(`P${curMonthNum.toString().padStart(3, '0')}`); // P008
        }

        try {
            const testingRecipient = ["neha.sain.ext@nokia.com"];
            for (const p of periodArray) {
                await mailService.sendPTDUpdateAlert(testingRecipient, p);
                console.log(`📧 PTD Mail Sent for period: ${p}`);
            }
        } catch (mailErr) {
            console.error("❌ Mail failed:", mailErr.message);
        }

        triggerAutoSync('ptd_uploaded');
        
        // Final Response jiske baad UI pe Successful wala pop-up aata h
        res.status(200).json({ message: "Everything Uploaded and Synced Successfully!" });

    } catch (error) { 
        console.error("❌ PTD ERROR:", error); 
        res.status(500).json({ error: error.message }); 
    }
};