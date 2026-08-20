const db = require('../config/db');
const { triggerAutoSync } = require('./cronController');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const fs = require('fs');
const NodeCache = require('node-cache');

const filterCache = new NodeCache({ stdTTL: 300 });
const inFlightRequests = new Map();

// ==============================
// COMMON RLS FUNCTION
// ==============================
const applyRLS = (type, allowedCustomers, conditions, params) => {
    if (type === 'super_admin') return;

    if (allowedCustomers && typeof allowedCustomers === 'string') {
        // 🔥 FIX: Split by ||| to handle commas inside customer names!
        const customersArray = allowedCustomers
            .split(',')
            .map(c => c.trim().toLowerCase())
            .filter(Boolean);

        if (customersArray.length > 0) {
            conditions.push(`TRIM(LOWER(customer)) IN (?)`);
            params.push(customersArray);
            return;
        }
    }
    conditions.push(`1=0`);
};

// ══════════════════════════════════════════════════
// SHARED FILTER BUILDER — SummaryView + Dashboard
// dono pages yahi use karenge — guaranteed sync
// ══════════════════════════════════════════════════
const buildCommonFilters = (reqQuery, conditions, params) => {
    const filterColumnMap = {
        'bu':              'bu',
        'customer':        'customer',
        'loa_id':          'loa_id',
        'loa_name':        'loa_name',
        'wbs_type':        'wbs_type',
        'wbs':             'wbs_element_single',
        'wbs_description': 'wbs_description',
        'period':          'period',
        'active_inactive': 'active_inactive',
    };
 
    let hasActiveFilters = false;
 
    Object.keys(filterColumnMap).forEach(key => {
        const vals = getValArray(reqQuery[key], reqQuery, key);
        if (!vals || vals.length === 0) return;
 
        hasActiveFilters = true;
        const dbCol = filterColumnMap[key];
 
        if (key === 'active_inactive') {
            const hasActive   = vals.some(v => v.toLowerCase() === 'active');
            const hasInactive = vals.some(v => v.toLowerCase() === 'inactive');
            if (hasActive && hasInactive) return; // dono hain = no filter needed
            if (hasActive)   conditions.push(`(TRIM(LOWER("${dbCol}")) = 'active' OR "${dbCol}" IS NULL OR TRIM("${dbCol}") = '')`);
            if (hasInactive) conditions.push(`TRIM(LOWER("${dbCol}")) = 'inactive'`);
        } else {
            const lowerVals = vals.map(v => String(v).trim().toLowerCase());
            conditions.push(`TRIM(LOWER("${dbCol}")) IN (?)`);
            params.push(lowerVals);
        }
    });
 
    // Category Type filter
    applyCategoryTypeFilter(reqQuery.category_type || reqQuery['category_type[]'], conditions);
 
    return hasActiveFilters;
};

const getValArray = (val, reqQuery = null, keyName = null) => {
    let targetVal = val;
    if (reqQuery && keyName && targetVal === undefined) {
        targetVal = reqQuery[`${keyName}[]`] || reqQuery[`${keyName}%5B%5D`];
    }

    if (targetVal === undefined || targetVal === null || targetVal === '' || targetVal === 'null' || (Array.isArray(targetVal) && targetVal.length === 0)) {
        return null;
    }
    let arr = Array.isArray(targetVal) ? targetVal : targetVal.toString().split(',').map(v => v.trim()).filter(Boolean);
    arr = arr.filter(v => v.toLowerCase() !== 'all'); 
    return arr.length > 0 ? arr : null;
};

const applyCategoryTypeFilter = (catType, conditions) => {
    let catArr = Array.isArray(catType) ? catType : (catType ? catType.split(",") : []);
    catArr = catArr.map(v => v.trim()).filter(Boolean);

    const hasAll = catArr.includes("All");
    const hasLM = catArr.includes("Local Materials");

    if (hasAll && hasLM) return;
    if (hasLM && !hasAll) { conditions.push("TRIM(categories) = 'Local Materials'"); return; }
    if (hasAll || catArr.length === 0) { conditions.push("TRIM(categories) <> 'Local Materials'"); return; }
};

// ==============================
// 2. Summary View Cascading Filters
// ==============================
exports.getFilterOptions = async (req, res) => {
    try {
        // 1. Smarter cache key — includes ALL query params for correct cascading
        const cacheKey = `FILTERS_${JSON.stringify({
            type: req.query.type,
            allowedCustomers: req.query.allowedCustomers,
            bu: req.query.bu,
            customer: req.query.customer,
            loa_id: req.query.loa_id,
            loa_name: req.query.loa_name,
            wbs_type: req.query.wbs_type,
            wbs: req.query.wbs,
            wbs_description: req.query.wbs_description,
            period: req.query.period,
            active_inactive: req.query.active_inactive,
            category_type: req.query.category_type,
        })}`;

        // 2. Cache hit — instant return
        const cached = filterCache.get(cacheKey);
        if (cached) return res.status(200).json(cached);

        // 3. In-flight deduplication — same request already running? wait for it
        if (inFlightRequests.has(cacheKey)) {
            try {
                const result = await inFlightRequests.get(cacheKey);
                return res.status(200).json(result);
            } catch (e) {
                // if in-flight failed, proceed to fetch fresh
            }
        }

        const { type, allowedCustomers } = req.query;

        // Base conditions (RLS + exclusions)
        let baseConditions = [
            "(categories IS NULL OR categories NOT IN ('Not to considered'))",
            "(cost_revenue IS NULL OR cost_revenue <> 'NTC')"
        ];
        let baseParams = [];
        applyRLS(type, allowedCustomers, baseConditions, baseParams);

        const columnMapping = {
            'bu': 'bu',
            'customer': 'customer',
            'loa_id': 'loa_id',
            'loa_name': 'loa_name',
            'wbs_type': 'wbs_type',
            'wbs': 'wbs_element_single',
            'wbs_description': 'wbs_description',
            'period': 'period',
            'active_inactive': 'active_inactive',
            'category_type': 'categories'
        };

        // Build one filter query for a specific target field
        const getFilteredOptions = async (targetField) => {
            let conditions = [...baseConditions];
            let filterValues = [...baseParams];

            if (targetField !== 'category_type') {
                applyCategoryTypeFilter(req.query.category_type, conditions);
            }

            Object.keys(columnMapping).forEach(key => {
                if (key !== targetField && key !== 'category_type') {
                    const vals = getValArray(req.query[key], req.query, key);
                    if (vals && vals.length > 0 && !vals.includes('All')) {
                        const dbCol = columnMapping[key];
                        if (key === 'active_inactive') {
                            const hasActive = vals.some(v => v.toLowerCase() === 'active');
                            const hasInactive = vals.some(v => v.toLowerCase() === 'inactive');
                            if (hasActive && hasInactive) return;
                            if (hasActive) conditions.push(`(TRIM(LOWER("${dbCol}")) = 'active' OR "${dbCol}" IS NULL OR TRIM("${dbCol}") = '')`);
                            else if (hasInactive) conditions.push(`TRIM(LOWER("${dbCol}")) = 'inactive'`);
                        } else {
                            const lowerVals = vals.map(v => v.trim().toLowerCase());
                            conditions.push(`TRIM(LOWER("${dbCol}")) IN (?)`);
                            filterValues.push(lowerVals);
                        }
                    }
                }
            });

            const dbColName = columnMapping[targetField];
            const sortOrder = targetField === 'period' ? 'DESC' : 'ASC';
            const sql = `
                SELECT DISTINCT "${dbColName}" as value 
                FROM final_dashboard_table 
                WHERE ${conditions.join(' AND ')} 
                AND "${dbColName}" IS NOT NULL AND "${dbColName}" <> ''
                ORDER BY 1 ${sortOrder}
                LIMIT 500
            `;
            const [rows] = await db.query(sql, filterValues);
            return rows.map(r => r.value);
        };

        // 4. Create the promise and register it for dedup
        const fetchPromise = (async () => {
            // 🔥 SEQUENTIAL instead of Promise.all — less DB pressure
            // Tradeoff: slightly slower but NEVER overwhelms pool
            // On cached data (second request): still fast from cache
            const keys = ['bu', 'customer', 'loa_id', 'loa_name', 'wbs_type', 'wbs', 'wbs_description', 'period'];
            const results = {};

            for (const k of keys) {
                results[k] = await getFilteredOptions(k);
            }

            const response = {
                category_type: ['All', 'Local Materials'],
                active_inactive: ['Active', 'Inactive'],
                ...results,
            };

            // Cache for 2 minutes
            filterCache.set(cacheKey, response, 120);
            return response;
        })();

        inFlightRequests.set(cacheKey, fetchPromise);

        const response = await fetchPromise;
        inFlightRequests.delete(cacheKey);

        return res.status(200).json(response);

    } catch (error) {
        console.error("Filter Options Error:", error.message);
        // Don't expose internal error details to client
        res.status(500).json({ error: "Filter options temporarily unavailable. Please retry." });
    }
};

const calculateSM = (rev, cost) => {
    const r = Math.abs(parseFloat(rev) || 0); 
    const c = parseFloat(cost) || 0;
    if (r === 0) return "0.00"; 
    const margin = ((r - c) / r) * 100;
    return margin.toFixed(2);
};

// 🔥 DYNAMIC ASBL COLUMN CALCULATOR BY WBS TYPE
const getDynamicSumColumns = (wbsTypes, prefix) => {
    if (!wbsTypes || wbsTypes.length === 0 || wbsTypes.includes('All')) return "0";
    let cols = [];
    if (wbsTypes.some(v => v.toLowerCase().includes('project'))) cols.push(`COALESCE(${prefix}_project, 0)`);
    if (wbsTypes.some(v => v.toLowerCase().includes('amc'))) cols.push(`COALESCE(${prefix}_amc, 0)`);
    if (wbsTypes.some(v => v.toLowerCase().includes('warranty'))) cols.push(`COALESCE(${prefix}_warranty, 0)`);
    return cols.length > 0 ? `(${cols.join(' + ')})` : "0";
};

// 🔥 SMART DYNAMIC NON-COMMITTED COLUMN CALCULATOR (Fallback Enabled)
const getDynamicNCColumns = (wbsTypes) => {
    if (!wbsTypes || wbsTypes.length === 0 || wbsTypes.includes('All')) return "0";
    let cols = [];
    if (wbsTypes.some(v => v.toLowerCase().includes('project'))) {
        cols.push(`COALESCE(NULLIF(non_committed_editable_project, 0), COALESCE(NULLIF(non_committed_project, 0), 0))`);
    }
    if (wbsTypes.some(v => v.toLowerCase().includes('amc'))) {
        cols.push(`COALESCE(NULLIF(non_committed_editable_amc, 0), COALESCE(NULLIF(non_committed_amc, 0), 0))`);
    }
    if (wbsTypes.some(v => v.toLowerCase().includes('warranty'))) {
        cols.push(`COALESCE(NULLIF(non_committed_editable_warranty, 0), COALESCE(NULLIF(non_committed_warranty, 0), 0))`);
    }
    return cols.length > 0 ? `(${cols.join(' + ')})` : "0";
};

// ==============================
// 3. WBS Summary View (Matrix - DYNAMIC ASBL/NC + ZERO VALUE ROWS FILTERED)
// ==============================
exports.getWbsSummary = async (req, res) => {
    try {
        const { draw, start, length, showAll, type, allowedCustomers } = req.query;
        const startIdx = parseInt(start) || 0;
        const limitIdx = parseInt(length) || 100;
 
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
 
        // DYNAMIC ASBL COLUMNS — unchanged
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const asblValExpression = asblCols === "0" ? "0" : `COALESCE(NULLIF(${asblCols}, 0), asbl, 0)`;
 
        // DYNAMIC NON-COMMITTED COLUMNS — unchanged
        const ncCols = getDynamicNCColumns(wTArr);
        const ncValExpression = ncCols === "0" ? "0" : `COALESCE(NULLIF(${ncCols}, 0), non_committed_editable, non_committed, 0)`;
 
        let filterParams = [];
        let conditions = [
            "(categories IS NULL OR categories NOT IN ('Not to considered'))",
            "(cost_revenue IS NULL OR cost_revenue <> 'NTC')"
        ];
        let baseParams = [];
 
        applyRLS(type, allowedCustomers, conditions, baseParams);
 
        // 🔥 REPLACED: filterColumnMap loop → buildCommonFilters (shared with Dashboard)
        const hasActiveFilters = buildCommonFilters(req.query, conditions, filterParams);
 
        const isDefaultView = !hasActiveFilters && (startIdx === 0);
        const cacheKey = `DEFAULT_SUMMARY_VIEW_${type}_${allowedCustomers || 'ALL'}`;
 
        if (isDefaultView) {
            const cachedPage = filterCache.get(cacheKey);
            if (cachedPage) {
                return res.status(200).json({ ...cachedPage, draw: parseInt(draw) || 0 });
            }
        }
 
        const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const combinedParams = [...baseParams, ...filterParams];
 
        // ── MATRIX QUERY — 100% UNCHANGED ──
        const matrixQuery = `
            SELECT 
                t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories, 
                t."Merged_wbs_categories",
                
                ROUND(MAX(COALESCE(static.asbl_val, 0)), 2) as asbl,
                ROUND(MAX(COALESCE(static.asbl_loa_val, 0)), 2) as asbl_loa,
                
                ROUND(SUM(t.ptd_val), 2) as ptd, 
                ROUND(SUM(t.oc_val), 2) as open_commitment_KEUR, 
                ROUND(SUM(t.oc_val), 2) as open_commitment,
                
                ROUND(MAX(COALESCE(static.nc_val, 0)), 2) as non_committed_editable, 
                ROUND(MAX(COALESCE(static.nc_val, 0)), 2) as non_committed, 
 
                ROUND(SUM(t.ptd_val) + SUM(t.oc_val) + MAX(COALESCE(static.nc_val, 0)), 2) as eac,
                ROUND(MAX(COALESCE(static.asbl_val, 0)) - (SUM(t.ptd_val) + SUM(t.oc_val) + MAX(COALESCE(static.nc_val, 0))), 2) as eac_vs_asbl
 
            FROM (
                SELECT 
                    bu, customer, loa_id, loa_name, cost_revenue, categories, "Merged_wbs_categories",
                    ptd as ptd_val, open_commitment_KEUR as oc_val
                FROM final_dashboard_table
                ${whereClause}
            ) as t
            LEFT JOIN (
                SELECT 
                    "Merged_wbs_categories", 
                    MAX(${asblValExpression}) as asbl_val, 
                    MAX(asbl_loa) as asbl_loa_val,
                    MAX(${ncValExpression}) as nc_val
                FROM final_dashboard_table
                GROUP BY "Merged_wbs_categories"
            ) as static ON t."Merged_wbs_categories" = static."Merged_wbs_categories"
            
            GROUP BY t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories, t."Merged_wbs_categories"
            HAVING 1=1 
            ${String(showAll) === 'false' ? 'AND (ABS(SUM(t.ptd_val)) > 0.01 OR ABS(SUM(t.oc_val)) > 0.01 OR ABS(MAX(COALESCE(static.asbl_val, 0))) > 0.01 OR ABS(MAX(COALESCE(static.nc_val, 0))) > 0.01)' : ''}
            ORDER BY loa_name ASC, cost_revenue ASC
        `;
 
        const [dataRows] = await db.query(`${matrixQuery} LIMIT ?, ?`, [...combinedParams, startIdx, limitIdx]);
        const [countRes] = await db.query(`SELECT COUNT(*) as total FROM (${matrixQuery}) as temp`, combinedParams);
 
        const totalCount = parseInt(countRes[0]?.total || countRes[0]?.count || 0);
 
        // ── KPI QUERY — 100% UNCHANGED ──
        const kpiQuery = `
            SELECT 
                SUM(CASE WHEN cost_revenue = 'Revenue' THEN cat_asbl ELSE 0 END) as asbl_rev,
                SUM(CASE WHEN cost_revenue = 'Revenue' THEN cat_ptd ELSE 0 END) as ptd_rev,
                SUM(CASE WHEN cost_revenue = 'Cost' THEN cat_ptd ELSE 0 END) as ptd_cost,
                SUM(CASE WHEN cost_revenue = 'Revenue' THEN (cat_ptd + cat_oc + cat_nc) ELSE 0 END) as eac_rev,
                SUM(CASE WHEN cost_revenue = 'Cost' THEN (cat_ptd + cat_oc + cat_nc) ELSE 0 END) as eac_cost
            FROM (
                SELECT t.cost_revenue, t."Merged_wbs_categories",
                       MAX(static.asbl_val) as cat_asbl,
                       SUM(t.ptd_val) as cat_ptd,
                       SUM(t.oc_val) as cat_oc,
                       MAX(static.nc_val) as cat_nc
                FROM (
                    SELECT cost_revenue, "Merged_wbs_categories", ptd as ptd_val, open_commitment_KEUR as oc_val 
                    FROM final_dashboard_table ${whereClause}
                ) as t
                LEFT JOIN (
                    SELECT "Merged_wbs_categories", MAX(${asblValExpression}) as asbl_val, MAX(${ncValExpression}) as nc_val 
                    FROM final_dashboard_table GROUP BY 1
                ) as static ON t."Merged_wbs_categories" = static."Merged_wbs_categories"
                GROUP BY t.cost_revenue, t."Merged_wbs_categories"
            ) as aggregated_t
        `;
        
        const [kpiRes] = await db.query(kpiQuery, combinedParams);
        const k = kpiRes[0] || {};
 
        const responsePayload = {
            draw: parseInt(draw) || 0, 
            recordsTotal: totalCount,
            recordsFiltered: totalCount, 
            data: dataRows,
            kpis: {
                asbl_rev: Number(k.asbl_rev || 0).toFixed(2),
                asbl_cost: Number(k.asbl_cost || 0).toFixed(2),
                asbl_sm: calculateSM(k.asbl_rev, k.asbl_cost),
                ptd_rev: Number(k.ptd_rev || 0).toFixed(2),
                ptd_cost: Number(k.ptd_cost || 0).toFixed(2),
                ptd_sm: calculateSM(k.ptd_rev, k.ptd_cost),
                eac_sm: calculateSM(k.eac_rev, k.eac_cost)
            }
        };
 
        if (isDefaultView) {
            filterCache.set(cacheKey, responsePayload);
        }
 
        res.status(200).json(responsePayload);
 
    } catch (error) {
        console.error("WbsSummary Error:", error);
        res.status(500).json({ error: error.message });
    }
};

// ==============================
// 4. Summary View (Collapse)
// ==============================
exports.getWbsSummaryCollapse = async (req, res) => {
    try {
        const { draw, start, length, type, allowedCustomers } = req.query;
        const startIdx = parseInt(start) || 0;
        const limitIdx = parseInt(length) || 100;
 
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
 
        let conditions = [
            "(categories IS NULL OR categories NOT IN ('Not to considered'))",
            "(cost_revenue IS NULL OR cost_revenue <> 'NTC')"
        ];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
 
        let filterParams = [];
 
        // 🔥 REPLACED: manual dbFilters loop → buildCommonFilters (shared with Dashboard)
        buildCommonFilters(req.query, conditions, filterParams);
 
        const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
 
        // ASBL logic — unchanged
        let asblSumLogic = "MAX(asbl)";
        if (wTArr && wTArr.length > 0) {
            let parts = [];
            if (wTArr.some(v => v.toLowerCase().includes('project'))) parts.push("MAX(asbl_project)");
            if (wTArr.some(v => v.toLowerCase().includes('amc')))     parts.push("MAX(asbl_amc)");
            if (wTArr.some(v => v.toLowerCase().includes('warranty'))) parts.push("MAX(asbl_warranty)");
            if (parts.length > 0) asblSumLogic = `(${parts.join(' + ')})`;
        }
 
        // ── SQL QUERY — 100% UNCHANGED ──
        const sql = `
            SELECT 
                bu, customer, loa_name, loa_id, cost_revenue,
                ROUND(${asblSumLogic}, 2) AS asbl, 
                ROUND(MAX(asbl_loa), 2) AS asbl_loa,
                ROUND(SUM(ptd), 2) AS ptd,
                ROUND(SUM(open_commitment_KEUR), 2) AS open_commitment_KEUR, 
                ROUND(SUM(open_commitment_KEUR), 2) AS open_commitment, 
                ROUND(SUM(non_committed_editable), 2) AS non_committed_editable,
                ROUND(SUM(non_committed_editable), 2) AS non_committed,
                ROUND(SUM(ptd) + SUM(open_commitment_KEUR) + SUM(non_committed_editable), 2) as eac
            FROM final_dashboard_table
            ${whereClause}
            GROUP BY bu, customer, loa_name, loa_id, cost_revenue
            ORDER BY loa_name ASC
        `;
 
        const [dataRows] = await db.query(`${sql} LIMIT ?, ?`, [...baseParams, ...filterParams, startIdx, limitIdx]);
        const [countRes] = await db.query(`SELECT COUNT(*) as total FROM (${sql}) temp`, [...baseParams, ...filterParams]);
 
        const totalCount = parseInt(countRes[0]?.total || 0);
 
        res.status(200).json({
            draw: parseInt(draw) || 0,
            recordsTotal: totalCount,
            recordsFiltered: totalCount,
            data: dataRows
        });
    } catch (error) {
        console.error("WbsSummaryCollapse Error:", error);
        res.status(500).json({ error: error.message });
    }
};

// ==============================
// 5. FULL REFRESH (PostgreSQL Native - Blazing Fast)
// ==============================
// exports.fullRefresh = async (req, res) => {
//     try {
//         console.log("🚀 Starting Sync (PTD & OC Corrected - PostgreSQL Version)...");

//         // Phase 1: PTD Staging
//         await db.query(`DROP TABLE IF EXISTS stg_cj74_agg`);
//         await db.query(`
//             CREATE TABLE stg_cj74_agg AS
//             SELECT
//                 TRIM(REPLACE(REPLACE(REPLACE(object_1, ' ', ''), CHR(10), ''), CHR(13), '')) AS clean_wbs,
//                 cost_element,
//                 TRIM(CONCAT(year, '-P', LPAD(TRIM(per), 3, '0'))) AS period,
//                 SUM(
//                     CASE 
//                         WHEN TRIM(val_in_rc::text) ~ '^[+-]?[0-9]*\.?[0-9]+$' THEN CAST(TRIM(val_in_rc::text) AS NUMERIC(18,2))
//                         ELSE 0 
//                     END / 1000
//                 ) AS ptd_val
//             FROM cj74_new 
//             WHERE year IS NOT NULL AND per IS NOT NULL 
//               AND TRIM(year::text) != 'NULL' AND TRIM(per::text) != 'NULL'
//               AND TRIM(year::text) != '' AND TRIM(per::text) != ''
//             GROUP BY 1, 2, 3
//         `);
//         await db.query("CREATE INDEX idx_stg_cj74_wbs ON stg_cj74_agg (clean_wbs)");
//         await db.query("CREATE INDEX idx_stg_cj74_ce ON stg_cj74_agg (cost_element)");

//         // Phase 2: OC Staging
//         await db.query(`DROP TABLE IF EXISTS stg_cji5_agg`);
//         await db.query(`
//             CREATE TABLE stg_cji5_agg AS
//             SELECT
//                 TRIM(REPLACE(REPLACE(REPLACE(wbs_element, ' ', ''), CHR(10), ''), CHR(13), '')) AS clean_wbs,
//                 TRIM(cost_element) AS cost_element,
//                 SUM(
//                     CASE 
//                         WHEN TRIM(val_in_rep_cur::text) ~ '^[+-]?[0-9]*\.?[0-9]+$' THEN CAST(TRIM(val_in_rep_cur::text) AS NUMERIC(18,2))
//                         ELSE 0 
//                     END / 1000
//                 ) AS oc_val
//             FROM cji5_new 
//             GROUP BY 1, 2
//         `);
//         await db.query("CREATE INDEX idx_stg_cji5_wbs ON stg_cji5_agg (clean_wbs)");
//         await db.query("CREATE INDEX idx_stg_cji5_ce ON stg_cji5_agg (cost_element)");

//         // Phase 3: Master Mapping
//         await db.query(`DROP TABLE IF EXISTS stg_master_mapping`);
//         await db.query(`
//             CREATE TABLE stg_master_mapping AS
//             SELECT
//                 TRIM(m.single_wbs) AS single_wbs,
//                 m.bu, m.customer, m.loa_id, m.loa_name, m.merged_wbs, m.wbs_type, m.wbs_description,
//                 cm.categories, cm.cost_element, cm.cost_revenue AS mapped_cost_revenue,
//                 TRIM(CONCAT(COALESCE(m.merged_wbs, ''), '-', COALESCE(cm.categories, ''))) AS "Merged_wbs_categories"
//             FROM wbs_loa_id_mapping1 m
//             CROSS JOIN (SELECT DISTINCT cost_element, categories, cost_revenue FROM cost_mapping) cm
//         `);
//         await db.query("CREATE INDEX idx_stg_mm_wbs ON stg_master_mapping (single_wbs)");
//         await db.query("CREATE INDEX idx_stg_mm_ce ON stg_master_mapping (cost_element)");
//         await db.query('CREATE INDEX idx_stg_mm_cat ON stg_master_mapping ("Merged_wbs_categories")');

//         // Phase 4: Final Table Fill
//         await db.query("TRUNCATE TABLE final_dashboard_table");

//         const finalInsertSql = `
//             INSERT INTO final_dashboard_table 
//             (id, bu, customer, loa_id, loa_name, cost_revenue, categories, merged_wbs, active_inactive, 
//              asbl, asbl_amc, asbl_project, asbl_warranty, asbl_loa, 
//              non_committed, non_committed_amc, non_committed_project, non_committed_warranty,
//              non_committed_editable, non_committed_editable_amc, non_committed_editable_project, non_committed_editable_warranty,
//              period, ptd, wbs_element_single, wbs_type, wbs_description, 
//              open_commitment_KEUR, eac, eac_vs_asbl, "Merged_wbs_categories", updated_by, updated_at)
            
//             SELECT 
//                 id, bu, customer, loa_id, loa_name, cost_revenue, categories, merged_wbs, active_inactive,
                
//                 CASE WHEN rank_project = 1 THEN asbl ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN asbl_amc ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN asbl_project ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN asbl_warranty ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN asbl_loa ELSE 0 END,
                
//                 CASE WHEN rank_project = 1 THEN non_committed ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN non_committed_amc ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN non_committed_project ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN non_committed_warranty ELSE 0 END,
                
//                 CASE WHEN rank_project = 1 THEN non_committed_editable ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN non_committed_editable_amc ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN non_committed_editable_project ELSE 0 END,
//                 CASE WHEN rank_project = 1 THEN non_committed_editable_warranty ELSE 0 END,
                
//                 period, ptd, wbs_element_single, wbs_type, wbs_description,
                
//                 CASE WHEN rank_oc = 1 THEN oc_val_raw ELSE 0 END,
                
//                 (ptd + CASE WHEN rank_oc = 1 THEN oc_val_raw ELSE 0 END + CASE WHEN rank_project = 1 THEN non_committed_editable ELSE 0 END),
//                 (CASE WHEN rank_project = 1 THEN asbl ELSE 0 END - (ptd + CASE WHEN rank_oc = 1 THEN oc_val_raw ELSE 0 END + CASE WHEN rank_project = 1 THEN non_committed_editable ELSE 0 END)),
                
//                 "Merged_wbs_categories", updated_by, updated_at
//             FROM (
//                 SELECT 
//                     COALESCE(s.id::text, CONCAT('NEW-', m."Merged_wbs_categories")) AS id,
//                     COALESCE(s.bu, m.bu) AS bu, COALESCE(s.customer, m.customer) AS customer, 
//                     COALESCE(s.loa_id, m.loa_id) AS loa_id, COALESCE(s.loa_name, m.loa_name) AS loa_name,
//                     COALESCE(s.cost_revenue, m.mapped_cost_revenue) AS cost_revenue, m.categories, 
//                     COALESCE(s.merged_wbs, m.merged_wbs) AS merged_wbs, COALESCE(s.active_inactive, 'Active') AS active_inactive,
//                     COALESCE(s.asbl, 0) AS asbl, COALESCE(s.asbl_amc, 0) AS asbl_amc, COALESCE(s.asbl_project, 0) AS asbl_project, COALESCE(s.asbl_warranty, 0) AS asbl_warranty, COALESCE(s.asbl_loa, 0) AS asbl_loa,
//                     COALESCE(s.non_committed, 0) AS non_committed, COALESCE(s.non_committed_amc, 0) AS non_committed_amc, COALESCE(s.non_committed_project, 0) AS non_committed_project, COALESCE(s.non_committed_warranty, 0) AS non_committed_warranty,
//                     COALESCE(s.non_committed_editable, 0) AS non_committed_editable, COALESCE(s.non_committed_editable_amc, 0) AS non_committed_editable_amc, COALESCE(s.non_committed_editable_project, 0) AS non_committed_editable_project, COALESCE(s.non_committed_editable_warranty, 0) AS non_committed_editable_warranty,
//                     cj.period, COALESCE(cj.ptd_val, 0) AS ptd, m.single_wbs AS wbs_element_single, m.wbs_type, m.wbs_description,
//                     COALESCE(ci.oc_val, 0) AS oc_val_raw, m."Merged_wbs_categories", s.updated_by, s.updated_at,
                    
//                     ROW_NUMBER() OVER (PARTITION BY m.single_wbs, m.cost_element ORDER BY cj.period DESC) AS rank_oc,
//                     ROW_NUMBER() OVER (PARTITION BY m."Merged_wbs_categories" ORDER BY cj.period DESC) AS rank_project
//                 FROM stg_master_mapping m
//                 LEFT JOIN stg_cj74_agg cj ON (m.single_wbs = cj.clean_wbs AND m.cost_element = cj.cost_element)
//                 LEFT JOIN stg_cji5_agg ci ON (m.single_wbs = ci.clean_wbs AND m.cost_element = ci.cost_element)
//                 LEFT JOIN summary s ON (m."Merged_wbs_categories" = s."Merged_wbs_category")
//                 WHERE cj.ptd_val IS NOT NULL OR ci.oc_val IS NOT NULL OR s.asbl > 0
//             ) AS final_src
//         `;
        
//         await db.query(finalInsertSql);

//         // Phase 5: Flush Cache & Sync Drilldowns
//         filterCache.flushAll(); 
//         if (typeof exports.syncDrilldownTables === 'function') {
//             await exports.syncDrilldownTables();
//         }

//         // Phase 6: Pre-Aggregated Table for Blazing Fast Summary
//         await db.query(`DROP TABLE IF EXISTS summary_matrix_aggregated`);
//         await db.query(`
//             CREATE TABLE summary_matrix_aggregated AS
//             SELECT 
//                 t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories, 
//                 t."Merged_wbs_categories",
//                 ROUND(MAX(COALESCE(static.asbl_val, 0)), 2) as asbl,
//                 ROUND(MAX(COALESCE(static.asbl_loa_val, 0)), 2) as asbl_loa,
//                 ROUND(SUM(t.ptd_val), 2) as ptd, 
//                 ROUND(SUM(t.oc_val), 2) as open_commitment_KEUR, 
//                 ROUND(SUM(t.oc_val), 2) as open_commitment,
//                 ROUND(MAX(COALESCE(static.nc_val, 0)), 2) as non_committed_editable, 
//                 ROUND(MAX(COALESCE(static.nc_val, 0)), 2) as non_committed, 
//                 ROUND(SUM(t.ptd_val) + SUM(t.oc_val) + MAX(COALESCE(static.nc_val, 0)), 2) as eac,
//                 ROUND(MAX(COALESCE(static.asbl_val, 0)) - (SUM(t.ptd_val) + SUM(t.oc_val) + MAX(COALESCE(static.nc_val, 0))), 2) as eac_vs_asbl
//             FROM (
//                 SELECT 
//                     bu, customer, loa_id, loa_name, cost_revenue, categories, "Merged_wbs_categories",
//                     ptd as ptd_val, open_commitment_KEUR as oc_val
//                 FROM final_dashboard_table
//             ) as t
//             LEFT JOIN (
//                 SELECT 
//                     "Merged_wbs_categories", 
//                     MAX(asbl) as asbl_val, 
//                     MAX(asbl_loa) as asbl_loa_val,
//                     MAX(non_committed_editable) as nc_val
//                 FROM final_dashboard_table
//                 GROUP BY "Merged_wbs_categories"
//             ) as static ON t."Merged_wbs_categories" = static."Merged_wbs_categories"
//             GROUP BY t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories, t."Merged_wbs_categories"
//         `);
//         await db.query(`CREATE INDEX idx_sma_loa ON summary_matrix_aggregated (loa_id)`);
//         await db.query(`CREATE INDEX idx_sma_bu ON summary_matrix_aggregated (bu)`);
//         filterCache.flushAll(); // Clear old cache
//         console.log("⚡ Default Page 1 Pre-calculated & Saved in RAM for Instant User Load!");

//         res.status(200).json({ message: "Sync Success! Everything is now accurate and Postgres-compatible." });

//     } catch (error) {
//         console.error("Full Refresh Error:", error);
//         res.status(500).json({ error: error.message });
//     }
// };


// 🔥 NEW: STANDALONE CORE ENGINE (Can be called by Cron or API)
exports.runFullSyncCore = async () => {
    console.log("🚀 Starting Sync Engine (PostgreSQL)...");

    // Phase 1: PTD Staging
    await db.query(`DROP TABLE IF EXISTS stg_cj74_agg`);
    await db.query(`
        CREATE TABLE stg_cj74_agg AS
        SELECT TRIM(REPLACE(REPLACE(REPLACE(object_1, ' ', ''), CHR(10), ''), CHR(13), '')) AS clean_wbs, cost_element,
            TRIM(CONCAT(year, '-P', LPAD(TRIM(per), 3, '0'))) AS period,
            SUM(CASE WHEN TRIM(val_in_rc::text) ~ '^[+-]?[0-9]*\.?[0-9]+$' THEN CAST(TRIM(val_in_rc::text) AS NUMERIC(18,2)) ELSE 0 END / 1000) AS ptd_val
        FROM cj74_new 
        WHERE year IS NOT NULL AND per IS NOT NULL AND TRIM(year::text) != 'NULL' AND TRIM(per::text) != 'NULL' AND TRIM(year::text) != '' AND TRIM(per::text) != ''
        GROUP BY 1, 2, 3
    `);
    await db.query("CREATE INDEX idx_stg_cj74_wbs ON stg_cj74_agg (clean_wbs)");
    await db.query("CREATE INDEX idx_stg_cj74_ce ON stg_cj74_agg (cost_element)");

    // Phase 2: OC Staging
    await db.query(`DROP TABLE IF EXISTS stg_cji5_agg`);
    await db.query(`
        CREATE TABLE stg_cji5_agg AS
        SELECT TRIM(REPLACE(REPLACE(REPLACE(wbs_element, ' ', ''), CHR(10), ''), CHR(13), '')) AS clean_wbs, TRIM(cost_element) AS cost_element,
            SUM(CASE WHEN TRIM(val_in_rep_cur::text) ~ '^[+-]?[0-9]*\.?[0-9]+$' THEN CAST(TRIM(val_in_rep_cur::text) AS NUMERIC(18,2)) ELSE 0 END / 1000) AS oc_val
        FROM cji5_new 
        GROUP BY 1, 2
    `);
    await db.query("CREATE INDEX idx_stg_cji5_wbs ON stg_cji5_agg (clean_wbs)");
    await db.query("CREATE INDEX idx_stg_cji5_ce ON stg_cji5_agg (cost_element)");

    // Phase 3: Master Mapping
    await db.query(`DROP TABLE IF EXISTS stg_master_mapping`);
    await db.query(`
        CREATE TABLE stg_master_mapping AS
        SELECT TRIM(m.single_wbs) AS single_wbs, m.bu, m.customer, m.loa_id, m.loa_name, m.merged_wbs, m.wbs_type, m.wbs_description, cm.categories, cm.cost_element, cm.cost_revenue AS mapped_cost_revenue,
            TRIM(CONCAT(COALESCE(m.merged_wbs, ''), '-', COALESCE(cm.categories, ''))) AS "Merged_wbs_categories"
        FROM wbs_loa_id_mapping1 m CROSS JOIN (SELECT DISTINCT cost_element, categories, cost_revenue FROM cost_mapping) cm
    `);
    await db.query("CREATE INDEX idx_stg_mm_wbs ON stg_master_mapping (single_wbs)");
    await db.query("CREATE INDEX idx_stg_mm_ce ON stg_master_mapping (cost_element)");
    await db.query('CREATE INDEX idx_stg_mm_cat ON stg_master_mapping ("Merged_wbs_categories")');

    // Phase 4: Final Table Fill
    await db.query("TRUNCATE TABLE final_dashboard_table");
    const finalInsertSql = `
        INSERT INTO final_dashboard_table 
        (id, bu, customer, loa_id, loa_name, cost_revenue, categories, merged_wbs, active_inactive, 
         asbl, asbl_amc, asbl_project, asbl_warranty, asbl_loa, non_committed, non_committed_amc, non_committed_project, non_committed_warranty,
         non_committed_editable, non_committed_editable_amc, non_committed_editable_project, non_committed_editable_warranty,
         period, ptd, wbs_element_single, wbs_type, wbs_description, open_commitment_KEUR, eac, eac_vs_asbl, "Merged_wbs_categories", updated_by, updated_at)
        SELECT id, bu, customer, loa_id, loa_name, cost_revenue, categories, merged_wbs, active_inactive,
            CASE WHEN rank_project = 1 THEN asbl ELSE 0 END, CASE WHEN rank_project = 1 THEN asbl_amc ELSE 0 END, CASE WHEN rank_project = 1 THEN asbl_project ELSE 0 END, CASE WHEN rank_project = 1 THEN asbl_warranty ELSE 0 END, CASE WHEN rank_project = 1 THEN asbl_loa ELSE 0 END,
            CASE WHEN rank_project = 1 THEN non_committed ELSE 0 END, CASE WHEN rank_project = 1 THEN non_committed_amc ELSE 0 END, CASE WHEN rank_project = 1 THEN non_committed_project ELSE 0 END, CASE WHEN rank_project = 1 THEN non_committed_warranty ELSE 0 END,
            CASE WHEN rank_project = 1 THEN non_committed_editable ELSE 0 END, CASE WHEN rank_project = 1 THEN non_committed_editable_amc ELSE 0 END, CASE WHEN rank_project = 1 THEN non_committed_editable_project ELSE 0 END, CASE WHEN rank_project = 1 THEN non_committed_editable_warranty ELSE 0 END,
            period, ptd, wbs_element_single, wbs_type, wbs_description, CASE WHEN rank_oc = 1 THEN oc_val_raw ELSE 0 END,
            (ptd + CASE WHEN rank_oc = 1 THEN oc_val_raw ELSE 0 END + CASE WHEN rank_project = 1 THEN non_committed_editable ELSE 0 END),
            (CASE WHEN rank_project = 1 THEN asbl ELSE 0 END - (ptd + CASE WHEN rank_oc = 1 THEN oc_val_raw ELSE 0 END + CASE WHEN rank_project = 1 THEN non_committed_editable ELSE 0 END)),
            "Merged_wbs_categories", updated_by, updated_at
        FROM (
            SELECT COALESCE(s.id::text, CONCAT('NEW-', m."Merged_wbs_categories")) AS id, COALESCE(s.bu, m.bu) AS bu, COALESCE(s.customer, m.customer) AS customer, COALESCE(s.loa_id, m.loa_id) AS loa_id, COALESCE(s.loa_name, m.loa_name) AS loa_name, COALESCE(s.cost_revenue, m.mapped_cost_revenue) AS cost_revenue, m.categories, COALESCE(s.merged_wbs, m.merged_wbs) AS merged_wbs, COALESCE(s.active_inactive, 'Active') AS active_inactive,
                COALESCE(s.asbl, 0) AS asbl, COALESCE(s.asbl_amc, 0) AS asbl_amc, COALESCE(s.asbl_project, 0) AS asbl_project, COALESCE(s.asbl_warranty, 0) AS asbl_warranty, COALESCE(s.asbl_loa, 0) AS asbl_loa, COALESCE(s.non_committed, 0) AS non_committed, COALESCE(s.non_committed_amc, 0) AS non_committed_amc, COALESCE(s.non_committed_project, 0) AS non_committed_project, COALESCE(s.non_committed_warranty, 0) AS non_committed_warranty, COALESCE(s.non_committed_editable, 0) AS non_committed_editable, COALESCE(s.non_committed_editable_amc, 0) AS non_committed_editable_amc, COALESCE(s.non_committed_editable_project, 0) AS non_committed_editable_project, COALESCE(s.non_committed_editable_warranty, 0) AS non_committed_editable_warranty,
                cj.period, COALESCE(cj.ptd_val, 0) AS ptd, m.single_wbs AS wbs_element_single, m.wbs_type, m.wbs_description, COALESCE(ci.oc_val, 0) AS oc_val_raw, m."Merged_wbs_categories", s.updated_by, s.updated_at,
                ROW_NUMBER() OVER (PARTITION BY m.single_wbs, m.cost_element ORDER BY cj.period DESC) AS rank_oc, ROW_NUMBER() OVER (PARTITION BY m."Merged_wbs_categories" ORDER BY cj.period DESC) AS rank_project
            FROM stg_master_mapping m LEFT JOIN stg_cj74_agg cj ON (m.single_wbs = cj.clean_wbs AND m.cost_element = cj.cost_element) LEFT JOIN stg_cji5_agg ci ON (m.single_wbs = ci.clean_wbs AND m.cost_element = ci.cost_element) LEFT JOIN summary s ON (m."Merged_wbs_categories" = s."Merged_wbs_category")
            WHERE cj.ptd_val IS NOT NULL OR ci.oc_val IS NOT NULL OR s.asbl > 0
        ) AS final_src
    `;
    await db.query(finalInsertSql);

    // Phase 5: Flush Cache & Sync Drilldowns
    filterCache.flushAll(); 
    if (typeof exports.syncDrilldownTables === 'function') {
        await exports.syncDrilldownTables();
    }
};

// 5. FULL REFRESH (API Wrapper)
exports.fullRefresh = async (req, res) => {
    try {
        await exports.runFullSyncCore();
        res.status(200).json({ message: "Sync Success! Everything is now accurate and Postgres-compatible." });
    } catch (error) {
        console.error("Full Refresh Error:", error);
        res.status(500).json({ error: error.message });
    }
};


// ==========================================
// DRILLDOWN TABLES SYNC HELPER
// ==========================================
exports.syncDrilldownTables = async () => {
    try {
        console.log("🔄 Syncing Drilldown Tables (PostgreSQL Direct)...");

        await db.query("TRUNCATE TABLE t_cj74_transformed");
        await db.query(`
            INSERT INTO t_cj74_transformed (
                id, sap_wbs, year, per, cost_element, cost_element_name, ptd_val, period, cocd, proj_def, 
                profit_ctr, name2, tcurr, value_trancurr, obcur, val_in_obj_crcy, val_in_rc, rcurr, 
                cost_element_descr, refdocno, document_no, doc_date, postg_date, offst_acct, 
                name_of_offsetting_account, material, material_description, name1, name22, created_on, 
                origin_form, user_name, pur_doc, quantity, purchase_order_text, loa_id, wbs_string, 
                wbs_type, wbs_description, categories, cost_revenue
            )
            SELECT 
                c.id, 
                TRIM(REPLACE(REPLACE(REPLACE(c.object_1, ' ', ''), CHR(10), ''), CHR(13), '')) AS sap_wbs, 
                c.year, 
                CASE WHEN TRIM(c.per::text) ~ '^[0-9]+$' THEN CAST(TRIM(c.per::text) AS INTEGER) ELSE NULL END AS per, 
                c.cost_element, c.cost_element_name, 
                CAST(COALESCE(c.val_in_rc, 0) AS NUMERIC(15,2)) / 1000 AS ptd_val, 
                TRIM(CONCAT(c.year, '-P', LPAD(CASE WHEN TRIM(c.per::text) ~ '^[0-9]+$' THEN TRIM(c.per::text) ELSE '0' END, 3, '0'))) AS period, 
                c.cocd, c.proj_def, c.profit_ctr, c.name2, c.tcurr, c.value_trancurr, c.obcur, 
                c.val_in_obj_crcy, c.val_in_rc, c.rcurr, c.cost_element_descr, c.refdocno, 
                c.document_no, c.doc_date, c.postg_date, c.offst_acct, c.name_of_offsetting_account, 
                c.material, c.material_description, c.name1, c.name22, c.created_on, c.frm, 
                c.user_name, c.pur_doc, c.quantity, c.purchase_order_text, 
                m.loa_id, m.merged_wbs, m.wbs_type, m.wbs_description, 
                cm.categories, cm.cost_revenue
            FROM cj74_new c
            LEFT JOIN (
                SELECT DISTINCT single_wbs, loa_id, merged_wbs, wbs_type, wbs_description
                FROM wbs_loa_id_mapping1
            ) m ON TRIM(REPLACE(REPLACE(REPLACE(c.object_1, ' ', ''), CHR(10), ''), CHR(13), '')) = m.single_wbs
            LEFT JOIN (
                SELECT DISTINCT cost_element, categories, cost_revenue
                FROM cost_mapping
            ) cm ON TRIM(c.cost_element) = TRIM(cm.cost_element)
        `);

        await db.query("TRUNCATE TABLE t_cji5_transformed");
        await db.query(`
            INSERT INTO t_cji5_transformed (
                id, project_def, sap_wbs, refdocno, item, co_object_name, supplier, name, exch_rate, 
                year, per, cost_element, cost_element_descr, matl_group, material, description, 
                user_name, docc, quantity, qty_plan, debit_date, doc_date, cocode, report_currency, 
                val_in_rep_cur, tcurr, value_tcur, obj_curr, value_in_obj_crcy, oc_val, loa_id, wbs_type, categories
            )
            SELECT 
                c.id, c.project_def, TRIM(c.wbs_element) AS sap_wbs, c.refdocno, c.item, 
                c.co_object_name, c.supplier, c.name, c.exch_rate, c.year, c.per, c.cost_element, 
                c.cost_element_descr, c.matl_group, c.material, c.description, c.user_name, c.docc, 
                c.quantity, c.qty_plan, c.debit_date, c.doc_date, c.cocode, c.report_currency, 
                c.val_in_rep_cur, c.tcurr, c.value_tcur, c.obj_curr, c.value_in_obj_crcy, 
                CAST(COALESCE(c.val_in_rep_cur, 0) AS NUMERIC(15,2)) / 1000 AS oc_val, 
                m.loa_id, m.wbs_type, cm.categories
            FROM cji5_new c
            LEFT JOIN (
                SELECT DISTINCT single_wbs, loa_id, wbs_type
                FROM wbs_loa_id_mapping1
            ) m ON TRIM(c.wbs_element) = m.single_wbs
            LEFT JOIN (
                SELECT DISTINCT cost_element, categories
                FROM cost_mapping
            ) cm ON TRIM(c.cost_element) = TRIM(cm.cost_element)
        `);

        console.log("✅ Drilldown Tables Synced on PostgreSQL!");
    } catch (err) {
        console.error("❌ Error in syncing Drilldown tables:", err);
    }
};

// ===========================================
// Helper: Dynamic Drilldown Filters Builder (PostgreSQL Case & Space Safe)
// ===========================================
const buildDrilldownConditions = (filters, tableName) => {
    let conds = [];
    let params = [];

    if (!filters) return { sql: '', params: [] };

    const getArray = (val) => {
        if (!val) return [];
        if (Array.isArray(val)) return val;
        return val.split(',').map(v => v.trim()).filter(v => v && v.toLowerCase() !== 'all');
    };

    // 1. WBS Type Filter
    const wbsTypes = getArray(filters.wbs_type).map(v => v.toLowerCase());
    if (wbsTypes.length > 0) {
        conds.push(`TRIM(LOWER(wbs_type)) IN (?)`);
        params.push(wbsTypes);
    }

    // 2. WBS Element Filter (sap_wbs)
    const wbsElements = getArray(filters.wbs).map(v => v.toLowerCase());
    if (wbsElements.length > 0) {
        conds.push(`TRIM(LOWER(sap_wbs)) IN (?)`);
        params.push(wbsElements);
    }

    // 3. Period Filter
    const periods = getArray(filters.period).map(v => v.toLowerCase());
    if (periods.length > 0) {
        if (tableName === 't_cj74_transformed') {
            conds.push(`TRIM(LOWER(period)) IN (?)`);
            params.push(periods);
        } else {
            conds.push(`TRIM(LOWER(CONCAT(year, '-P', LPAD(per::text, 3, '0')))) IN (?)`);
            params.push(periods);
        }
    }

    return { 
        sql: conds.length > 0 ? ' AND ' + conds.join(' AND ') : '', 
        params 
    };
};

// 🔥 NAYI COLUMN MAPPINGS (UI Titles ke saath sync)
const DRILL_MAPPING = {
    ptd: [
        { key: 'sap_wbs', header: 'WBS' }, { key: 'year', header: 'Year' }, { key: 'per', header: 'Per' },
        { key: 'cost_element', header: 'Cost Element' }, { key: 'cost_element_name', header: 'Cost Element Name' },
        { key: 'ptd_val', header: 'PTD VAL (K€)' }, { key: 'period', header: 'Period' }, { key: 'cocd', header: 'CoCd' },
        { key: 'proj_def', header: 'Project Def' }, { key: 'profit_ctr', header: 'Profit Ctr' },
        { key: 'tcurr', header: 'T Curr' }, { key: 'cost_element_descr', header: 'COST ELEMENT DESCR' },
        { key: 'refdocno', header: 'Ref Doc No' }, { key: 'document_no', header: 'Document No' },
        { key: 'doc_date', header: 'Doc Date' }, { key: 'postg_date', header: 'Postg Date' },
        { key: 'offst_acct', header: 'Offset Acct' }, { key: 'material', header: 'Material' },
        { key: 'material_description', header: 'Material Description' }, { key: 'created_on', header: 'Created On' },
        { key: 'user_name', header: 'User Name' }, { key: 'pur_doc', header: 'Pur Doc' },
        { key: 'purchase_order_text', header: 'Purchase Order Text' }, { key: 'loa_id', header: 'LOA ID' }
    ],
    oc: [
        { key: 'project_def', header: 'PROJ DEF' }, { key: 'sap_wbs', header: 'WBS' },
        { key: 'oc_val', header: 'OC VAL (K€)' }, { key: 'refdocno', header: 'REFDOCNO' },
        { key: 'item', header: 'ITEM' }, { key: 'co_object_name', header: 'CO_OBJECT_NAME' },
        { key: 'supplier', header: 'SUPPLIER' }, { key: 'name', header: 'NAME' },
        { key: 'exch_rate', header: 'EXCH_RATE' }, { key: 'year', header: 'YEAR' },
        { key: 'per', header: 'PER' }, { key: 'cost_element', header: 'COST_ELEMENT' },
        { key: 'cost_element_descr', header: 'COST_ELEMENT_DESCR' }, { key: 'matl_group', header: 'MATL GROUP' },
        { key: 'material', header: 'MATERIAL' }, { key: 'description', header: 'DESCRIPTION' },
        { key: 'user_name', header: 'USER_NAME' }, { key: 'docc', header: 'DOCC' },
        { key: 'quantity', header: 'QUANTITY' }, { key: 'qty_plan', header: 'QTY_PLAN' },
        { key: 'debit_date', header: 'DEBIT_DATE' }, { key: 'doc_date', header: 'DOC_DATE' },
        { key: 'cocode', header: 'COCODE' }, { key: 'report_currency', header: 'REPORT_CURRENCY' },
        { key: 'tcurr', header: 'TCURR' }, { key: 'value_tcur', header: 'VALUE TCUR' },
        { key: 'obj_curr', header: 'OBJ CURR' }, { key: 'value_in_obj_crcy', header: 'VALUE IN OBJ CRCY' },
        { key: 'loa_id', header: 'LOA ID' }
    ]
};

exports.getDrillDownData = async (req, res) => {
    try {
        const { field, row, filters } = req.body;
        const loaId = row?.loa_id;
        const category = row?.categories;

        if (!loaId) return res.status(400).json({ error: "Missing LOA ID" });

        const type = field === 'ptd' ? 'ptd' : 'oc';
        const tableName = field === 'ptd' ? 't_cj74_transformed' : 't_cji5_transformed';
        
        // 🔥 SELECT specifically only requested columns
        const selectColumns = DRILL_MAPPING[type].map(c => `"${c.key}"`).join(', ');
        
        let sql = `SELECT ${selectColumns} FROM ${tableName} WHERE TRIM(LOWER(loa_id)) = TRIM(LOWER(?))`;
        let params = [loaId];

        if (category) {
            sql += ` AND TRIM(LOWER(categories)) = TRIM(LOWER(?))`;
            params.push(category);
        }

        const dynamicFilters = buildDrilldownConditions(filters, tableName);
        sql += dynamicFilters.sql;
        params.push(...dynamicFilters.params);

        const [rows] = await db.query(sql + ` LIMIT 10000`, params);
        res.status(200).json(rows); 
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.exportDrillDown = async (req, res) => {
    try {
        const { field, loa_id, categories, filters } = req.query;
        let parsedFilters = {}; try { parsedFilters = JSON.parse(filters || '{}'); } catch(e) {}

        const type = field === 'ptd' ? 'ptd' : 'oc';
        const tableName = field === 'ptd' ? 't_cj74_transformed' : 't_cji5_transformed';
        const columnsToExport = DRILL_MAPPING[type];
        const selectFields = columnsToExport.map(c => `"${c.key}"`).join(', ');

        let sql = `SELECT ${selectFields} FROM ${tableName} WHERE TRIM(LOWER(loa_id)) = TRIM(LOWER(?))`;
        let params = [loa_id];

        if (categories && categories !== 'null') {
            sql += ` AND TRIM(LOWER(categories)) = TRIM(LOWER(?))`;
            params.push(categories);
        }

        const dynamicFilters = buildDrilldownConditions(parsedFilters, tableName);
        sql += dynamicFilters.sql;
        params.push(...dynamicFilters.params);

        const [rows] = await db.query(sql + ` ORDER BY year DESC`, params);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=DrillDown_${type}_${loa_id}.xlsx`);

        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
        const worksheet = workbook.addWorksheet('Details');

        // 🔥 Set headers exactly as UI titles
        worksheet.columns = columnsToExport.map(col => ({ header: col.header, key: col.key, width: 20 }));

        rows.forEach(row => {
            const cleanRow = { ...row };
            Object.keys(cleanRow).forEach(key => {
                if (key.includes('date') || key === 'created_on') {
                    cleanRow[key] = cleanRow[key] ? new Date(cleanRow[key]).toLocaleDateString('en-GB') : '-';
                }
            });
            worksheet.addRow(cleanRow).commit();
        });
        await workbook.commit();
    } catch (error) { res.status(500).send("Export failed"); }
};

// 🔥 DYNAMIC NON-COMMITTED UPDATER (LOA ID + LOA Name Dual Matching)
exports.updateNonCommitted = async (req, res) => {
    const { updates, createdBy } = req.body;
    try {
        const monthYear = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');

        for (let item of updates) {
            const { loa_name, categories, value, wbs_type } = item;
            const numVal = parseFloat(value) || 0;

            // Mapping Column identification
            let ncCol = 'non_committed_editable_project';
            const wTypeStr = String(wbs_type || '').toLowerCase();
            if (wTypeStr.includes('amc')) ncCol = 'non_committed_editable_amc';
            if (wTypeStr.includes('warranty')) ncCol = 'non_committed_editable_warranty';

            const [existing] = await db.query(
                `SELECT non_committed_editable, bu, customer, loa_id, active_inactive 
                 FROM summary WHERE TRIM(LOWER(loa_name)) = TRIM(LOWER(?)) AND TRIM(LOWER(categories)) = TRIM(LOWER(?))`,
                [loa_name, categories]
            );

            if (!existing || existing.length === 0) continue;
            const oldValue = existing[0].non_committed_editable || 0;

            // 🔥 Update both: Master editable column and Specific bucket column
            const updateSql = `
                UPDATE summary SET 
                non_committed_editable = ?, 
                ${ncCol} = ?, 
                updated_by = ? 
                WHERE TRIM(LOWER(loa_name)) = TRIM(LOWER(?)) AND TRIM(LOWER(categories)) = TRIM(LOWER(?))`;

            await db.query(updateSql, [numVal, numVal, createdBy, loa_name, categories]);
            
            // Mirror same to dashboard table
            await db.query(updateSql.replace('summary', 'final_dashboard_table'), [numVal, numVal, createdBy, loa_name, categories]);

            // Log activity
            await db.query(
                `INSERT INTO user_activity_logs (user_email, bu, customer, loa_name, loa_id, categories, old_value, new_value, active_inactive, month_year, wbs_type, is_finalized) 
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, false)`,
                [createdBy, existing[0].bu, existing[0].customer, loa_name, existing[0].loa_id, categories, oldValue, numVal, existing[0].active_inactive, monthYear, wbs_type]
            );
        }

        // Global Recalculation for UI
        await db.query(`
            UPDATE final_dashboard_table 
            SET eac = (ptd + open_commitment_KEUR + non_committed_editable),
                eac_vs_asbl = (asbl - (ptd + open_commitment_KEUR + non_committed_editable))
            WHERE non_committed_editable <> 0 OR ABS(non_committed - non_committed_editable) > 0.01
        `);

        filterCache.flushAll(); 
        res.status(200).json({ message: "Changes saved to draft!" });
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getUserActivityLogs = async (req, res) => {
    try {
        const [rows] = await db.query(`SELECT id, user_email, bu, customer, loa_name, loa_id, categories, old_value, new_value, month_year, created_at FROM user_activity_logs ORDER BY created_at DESC`);
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getAsblActivityLogs = async (req, res) => {
    try {
        const [rows] = await db.query(`SELECT * FROM asbl_activity_logs ORDER BY created_at DESC`);
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getProjectActivityLogs = async (req, res) => {
    try {
        const [rows] = await db.query(`SELECT * FROM project_activity_logs ORDER BY created_at DESC`);
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getPendingUsers = async (req, res) => {
    try {
        const monthYear = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');
        const [rows] = await db.query(`
            SELECT u.email, u.type FROM users u
            LEFT JOIN (SELECT DISTINCT user_email FROM user_activity_logs WHERE month_year = ?) l ON u.email = l.user_email
            WHERE l.user_email IS NULL ORDER BY u.email
        `, [monthYear]);
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

// ===========================================
// Export Summary View & Dahsboard Page Data to Excel (100% Sync with UI Table)
// ===========================================
exports.exportToExcel = async (req, res) => {
    try {
        const { showAll, collapseView, type, allowedCustomers } = req.query;
 
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        const wEArr = getValArray(req.query.wbs, req.query, 'wbs');
 
        // 🔥 STRICT DYNAMIC ASBL & NON-COMMITTED (Exact match with UI logic)
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const asblValExpression = asblCols !== "0" ? asblCols : "0";
 
        const ncCols = getDynamicNCColumns(wTArr);
        const ncValExpression = ncCols !== "0" ? ncCols : "0";
 
        let conditions = ["(categories IS NULL OR categories NOT IN ('Not to considered'))", "(cost_revenue IS NULL OR cost_revenue <> 'NTC')"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
 
        const filterColumnMap = {
            'bu': 'bu', 'customer': 'customer', 'loa_id': 'loa_id', 'loa_name': 'loa_name',
            'wbs_type': 'wbs_type', 'wbs': 'wbs_element_single', 'wbs_description': 'wbs_description',
            'period': 'period', 'active_inactive': 'active_inactive'
        };
 
        let filterParams = [];
        Object.keys(filterColumnMap).forEach(key => {
            const vals = getValArray(req.query[key], req.query, key);
            if (vals && vals.length > 0 && !vals.includes('All')) {
                const dbCol = filterColumnMap[key];
                if (key === 'active_inactive') {
                    const hasActive = vals.some(v => String(v).toLowerCase() === 'active');
                    const hasInactive = vals.some(v => String(v).toLowerCase() === 'inactive');
                    if (hasActive && hasInactive) return;
                    if (hasActive) conditions.push(`(TRIM(LOWER("${dbCol}")) = 'active' OR "${dbCol}" IS NULL OR TRIM("${dbCol}") = '')`);
                    else if (hasInactive) conditions.push(`TRIM(LOWER("${dbCol}")) = 'inactive'`);
                } else {
                    const lowerVals = vals.map(v => String(v).trim().toLowerCase());
                    conditions.push(`TRIM(LOWER("${dbCol}")) IN (?)`);
                    filterParams.push(lowerVals);
                }
            }
        });
 
        const catTypeVal = req.query.category_type || req.query['category_type[]'];
        applyCategoryTypeFilter(catTypeVal, conditions);
 
        const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const combinedParams = [...baseParams, ...filterParams];
 
        let exportQuery = '';
 
        // 🟢 If user is exporting COLLAPSED View (Cost Level)
        if (String(collapseView) === 'true') {
            exportQuery = `
                SELECT 
                    bu, customer, loa_name, loa_id, cost_revenue,
                    ROUND(SUM(asbl), 2) AS asbl, 
                    ROUND(MAX(asbl_loa), 2) AS asbl_loa,
                    ROUND(SUM(ptd), 2) AS ptd,
                    ROUND(SUM(open_commitment), 2) AS open_commitment,
                    ROUND(SUM(non_committed), 2) AS non_committed,
                    ROUND(SUM(ptd) + SUM(open_commitment) + SUM(non_committed), 2) AS eac,
                    ROUND(SUM(asbl) - (SUM(ptd) + SUM(open_commitment) + SUM(non_committed)), 2) AS eac_vs_asbl
                FROM (
                    SELECT 
                        t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories,
                        MAX(COALESCE(static.asbl_val, 0)) as asbl,
                        MAX(COALESCE(static.asbl_loa_val, 0)) as asbl_loa,
                        SUM(t.ptd_val) as ptd, 
                        SUM(t.oc_val) as open_commitment, 
                        MAX(COALESCE(static.nc_val, 0)) as non_committed
                    FROM (
                        SELECT 
                            bu, customer, loa_id, loa_name, cost_revenue, categories, "Merged_wbs_categories",
                            ptd as ptd_val, open_commitment_KEUR as oc_val
                        FROM final_dashboard_table
                        ${whereClause}
                    ) as t
                    LEFT JOIN (
                        SELECT 
                            "Merged_wbs_categories", 
                            MAX(${asblValExpression}) as asbl_val, 
                            MAX(asbl_loa) as asbl_loa_val,
                            MAX(${ncValExpression}) as nc_val      
                        FROM final_dashboard_table
                        GROUP BY "Merged_wbs_categories"
                    ) as static ON t."Merged_wbs_categories" = static."Merged_wbs_categories"
                    GROUP BY t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories, t."Merged_wbs_categories"
                ) as category_level_data
                GROUP BY bu, customer, loa_name, loa_id, cost_revenue
                ORDER BY loa_name ASC, cost_revenue ASC
            `;
        } 
        // 🟢 If user is exporting NORMAL View (Element Level - Identical to UI)
        else {
            exportQuery = `
                SELECT 
                    t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories, 
                    ROUND(MAX(COALESCE(static.asbl_val, 0)), 2) as asbl,
                    ROUND(MAX(COALESCE(static.asbl_loa_val, 0)), 2) as asbl_loa,
                    ROUND(SUM(t.ptd_val), 2) as ptd, 
                    ROUND(SUM(t.oc_val), 2) as open_commitment, 
                    ROUND(MAX(COALESCE(static.nc_val, 0)), 2) as non_committed, 
                    ROUND(SUM(t.ptd_val) + SUM(t.oc_val) + MAX(COALESCE(static.nc_val, 0)), 2) as eac,
                    ROUND(MAX(COALESCE(static.asbl_val, 0)) - (SUM(t.ptd_val) + SUM(t.oc_val) + MAX(COALESCE(static.nc_val, 0))), 2) as eac_vs_asbl
                FROM (
                    SELECT 
                        bu, customer, loa_id, loa_name, cost_revenue, categories, "Merged_wbs_categories",
                        ptd as ptd_val, open_commitment_KEUR as oc_val
                    FROM final_dashboard_table
                    ${whereClause}
                ) as t
                LEFT JOIN (
                    SELECT 
                        "Merged_wbs_categories", 
                        MAX(${asblValExpression}) as asbl_val, 
                        MAX(asbl_loa) as asbl_loa_val,
                        MAX(${ncValExpression}) as nc_val      
                    FROM final_dashboard_table
                    GROUP BY "Merged_wbs_categories"
                ) as static ON t."Merged_wbs_categories" = static."Merged_wbs_categories"
                GROUP BY t.bu, t.customer, t.loa_id, t.loa_name, t.cost_revenue, t.categories, t."Merged_wbs_categories"
                HAVING 1=1 
                ${String(showAll) === 'false' ? 'AND (ABS(SUM(t.ptd_val)) > 0.01 OR ABS(SUM(t.oc_val)) > 0.01 OR ABS(MAX(COALESCE(static.asbl_val, 0))) > 0.01 OR ABS(MAX(COALESCE(static.nc_val, 0))) > 0.01)' : ''}
                ORDER BY loa_name ASC, categories ASC
            `;
        }
 
        const [rows] = await db.query(exportQuery, combinedParams);
 
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Summary_Export_${new Date().getTime()}.xlsx`);
        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
        const worksheet = workbook.addWorksheet('Matrix Data');
        
        const cols = [
            { header: 'BU', key: 'bu', width: 12 },
            { header: 'Customer', key: 'customer', width: 25 },
            { header: 'LOA Name', key: 'loa_name', width: 40 },
            { header: 'LOA ID', key: 'loa_id', width: 15 },
            { header: 'Cost/Revenue', key: 'cost_revenue', width: 15 }
        ];
        if (String(collapseView) !== 'true') cols.push({ header: 'Category', key: 'categories', width: 25 });
        cols.push(
            { header: 'ASBL', key: 'asbl', width: 15 },
            // { header: 'ASBL LOA', key: 'asbl_loa', width: 15 },
            { header: 'PTD', key: 'ptd', width: 15 },
            { header: 'Open Commitment', key: 'open_commitment', width: 15 },
            { header: 'Non Committed', key: 'non_committed', width: 15 },
            { header: 'EAC', key: 'eac', width: 15 },
            { header: 'EAC vs ASBL', key: 'eac_vs_asbl', width: 15 }
        );
        worksheet.columns = cols;
 
        // Ensure proper numeric formatting in Excel
        rows.forEach(row => {
            const cleanRow = { ...row };
            ['asbl', 'ptd', 'open_commitment', 'non_committed', 'eac', 'eac_vs_asbl'].forEach(k => {
                if (cleanRow[k] === null || cleanRow[k] === undefined) {
                    cleanRow[k] = (k === 'asbl' || k === 'eac_vs_asbl') ? '-' : 0;
                } else {
                    cleanRow[k] = Number(cleanRow[k]);
                }
            });
            worksheet.addRow(cleanRow).commit();
        });
        await workbook.commit();
 
    } catch (error) {
        console.error("Export Error:", error);
        res.status(500).send("Export failed: " + error.message);
    }
};

exports.clearDraftChanges = async (req, res) => {
    try {
        await db.query("UPDATE final_dashboard_table SET non_committed_editable = non_committed");
        await db.query("UPDATE summary SET non_committed_editable = non_committed");
        filterCache.flushAll();
        res.status(200).json({ message: "Draft cleared! All values reset to original." });
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.exportReviewExcel = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        let conditions = ["categories != 'Revenue'", "ABS(COALESCE(non_committed, 0) - COALESCE(non_committed_editable, 0)) > 0.01"];
        let params = [];
        applyRLS(type, allowedCustomers, conditions, params);
        buildCommonFilters(req.query, conditions, params);
        const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Review_Changes_Export.xlsx`);
        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
        const worksheet = workbook.addWorksheet('Review Data');
        
        worksheet.columns = [
            { header: 'BU', key: 'bu', width: 10 },
            { header: 'Customer', key: 'customer', width: 25 },
            { header: 'LOA Name', key: 'loa_name', width: 35 },
            { header: 'LOA ID', key: 'loa_id', width: 20 },
            { header: 'Category', key: 'categories', width: 25 },
            { header: 'ASBL', key: 'asbl', width: 15 },
            { header: 'PTD', key: 'ptd', width: 15 },
            { header: 'Open Commitment', key: 'open_commitment', width: 15 },
            { header: 'Old Non Committed', key: 'non_committed_original', width: 20 },
            { header: 'New Non Committed', key: 'non_committed', width: 20 },
            { header: 'EAC', key: 'eac', width: 15 },
            { header: 'Modified By', key: 'updated_by', width: 30 }, // 🔥 Added to Excel
            { header: 'Last Modified', key: 'updated_at', width: 25 }  // 🔥 Added to Excel
        ];

        const query = `
            SELECT bu, customer, loa_id, loa_name, cost_revenue, categories, MAX(asbl) as asbl, SUM(ptd) as ptd, 
            MAX(open_commitment_KEUR) as open_commitment, MAX(non_committed_editable) as non_committed, 
            MAX(non_committed) as non_committed_original, MAX(updated_by) as updated_by,
            TO_CHAR(MAX(updated_at), 'DD-Mon-YYYY HH24:MI') as updated_at,
            (SUM(ptd) + MAX(open_commitment_KEUR) + MAX(non_committed_editable)) as eac
            FROM final_dashboard_table ${whereClause}
            GROUP BY bu, customer, loa_id, loa_name, cost_revenue, categories
        `;
        const [rows] = await db.query(query, params);
        rows.forEach(row => worksheet.addRow(row).commit());
        await workbook.commit();
    } catch (error) { res.status(500).send("Export failed"); }
};

exports.getCategories = async (req, res) => {
    try {
        const [rows] = await db.query("SELECT DISTINCT categories FROM summary WHERE categories IS NOT NULL AND categories <> '' ORDER BY categories ASC");
        res.status(200).json(rows.map(r => r.categories));
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
};

exports.saveProjectData = async (req, res) => {
    const { bu, customer, loa_id, loa_name, wbs, asblData } = req.body;
    try {
        const [existing] = await db.query("SELECT id FROM summary WHERE loa_name = ?", [loa_name]);
        if (existing.length > 0) {
            await db.query("DELETE FROM summary WHERE loa_name = ?", [loa_name]);
        }

        const insertPromises = Object.keys(asblData).map(cat => {
            const val = asblData[cat] || 0;
            if (val === 0 || val === '') return null;
            return db.query(
                `INSERT INTO summary (bu, customer, loa_id, loa_name, merged_wbs, categories, asbl, active_inactive) 
                 VALUES (?, ?, ?, ?, ?, ?, ?, 'Active')`,
                [bu, customer, loa_id, loa_name, wbs, cat, val]
            );
        }).filter(p => p !== null);

        await Promise.all(insertPromises);
        filterCache.flushAll();
        res.status(200).json({ message: "Data saved successfully!" });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
};

exports.getDashboardFilters = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;

        const buildConditions = (excludeKey) => {
            let conditions = ["customer IS NOT NULL", "loa_name IS NOT NULL"];
            let params = [];

            applyRLS(type, allowedCustomers, conditions, params);

            // Category Type
            const { category_type } = req.query;
            if (!category_type) {
                conditions.push(`categories <> 'Local Materials'`);
            } else {
                let catArr = Array.isArray(category_type)
                    ? category_type
                    : category_type.split(',').map(v => v.trim());
                const hasAll = catArr.includes('All');
                const hasLM  = catArr.includes('Local Materials');
                if (hasAll && !hasLM)       conditions.push(`categories <> 'Local Materials'`);
                else if (!hasAll && hasLM)  conditions.push(`categories = 'Local Materials'`);
                else if (!hasAll && !hasLM) conditions.push(`categories <> 'Local Materials'`);
            }

            // 🔥 KEY FIX: filterColumnMap — same as buildCommonFilters
            const filterColumnMap = {
                'bu':              'bu',
                'customer':        'customer',
                'loa_id':          'loa_id',
                'loa_name':        'loa_name',
                'wbs_type':        'wbs_type',
                'wbs':             'wbs_element_single',
                'wbs_description': 'wbs_description',
                'period':          'period',
                'active_inactive': 'active_inactive',
            };

            Object.keys(filterColumnMap).forEach(key => {
                // Exclude current field (cascading logic)
                if (key === excludeKey) return;

                const vals = getValArray(req.query[key], req.query, key);
                if (!vals || vals.length === 0) return;

                const dbCol = filterColumnMap[key];

                if (key === 'active_inactive') {
                    const hasActive   = vals.some(v => v.toLowerCase() === 'active');
                    const hasInactive = vals.some(v => v.toLowerCase() === 'inactive');
                    if (hasActive && hasInactive) return;
                    if (hasActive)   conditions.push(`(TRIM(LOWER("${dbCol}")) = 'active' OR "${dbCol}" IS NULL OR TRIM("${dbCol}") = '')`);
                    if (hasInactive) conditions.push(`TRIM(LOWER("${dbCol}")) = 'inactive'`);
                } else {
                    const lowerVals = vals.map(v => String(v).trim().toLowerCase());
                    conditions.push(`TRIM(LOWER("${dbCol}")) IN (?)`);
                    params.push(lowerVals);
                }
            });

            // years — alag handle (period LIKE syntax)
            const { years } = req.query;
            if (years && excludeKey !== 'period') {
                const yearArray = years.split(',').map(y => y.trim()).filter(Boolean);
                if (yearArray.length > 0) {
                    conditions.push(`(${yearArray.map(() => 'period LIKE ?').join(' OR ')})`);
                    params.push(...yearArray.map(y => `${y}-%`));
                }
            }

            return {
                whereSql: conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '',
                params
            };
        };

        // Queries — excludeKey match karo filterColumnMap keys se
        const buQ      = buildConditions('bu');
        const [buRows] = await db.query(`SELECT DISTINCT bu FROM final_dashboard_table ${buQ.whereSql} AND bu IS NOT NULL ORDER BY bu ASC`, buQ.params);

        const wbsTypeQ      = buildConditions('wbs_type');
        const [wbsTypeRows] = await db.query(`SELECT DISTINCT wbs_type FROM final_dashboard_table ${wbsTypeQ.whereSql} AND wbs_type IS NOT NULL ORDER BY wbs_type ASC`, wbsTypeQ.params);

        const custQ         = buildConditions('customer');
        const [customerRows]= await db.query(`SELECT DISTINCT customer FROM final_dashboard_table ${custQ.whereSql} AND customer IS NOT NULL ORDER BY customer ASC`, custQ.params);

        const perQ          = buildConditions('period');
        const [periodRows]  = await db.query(`SELECT DISTINCT period FROM final_dashboard_table ${perQ.whereSql} AND period IS NOT NULL ORDER BY period DESC`, perQ.params);

        const loaIdQ        = buildConditions('loa_id');
        const [loaIdRows]   = await db.query(`SELECT DISTINCT loa_id FROM final_dashboard_table ${loaIdQ.whereSql} AND loa_id IS NOT NULL ORDER BY loa_id ASC`, loaIdQ.params);

        const loaQ          = buildConditions('loa_name');
        const [loaRows]     = await db.query(`SELECT DISTINCT loa_name FROM final_dashboard_table ${loaQ.whereSql} AND loa_name IS NOT NULL ORDER BY loa_name ASC`, loaQ.params);

        const wbsValQ       = buildConditions('wbs');
        const [wbsRows]     = await db.query(`SELECT DISTINCT wbs_element_single as wbs FROM final_dashboard_table ${wbsValQ.whereSql} AND wbs_element_single IS NOT NULL ORDER BY 1 ASC`, wbsValQ.params);

        const wbsDescQ      = buildConditions('wbs_description');
        const [wbsDescRows] = await db.query(`SELECT DISTINCT wbs_description FROM final_dashboard_table ${wbsDescQ.whereSql} AND wbs_description IS NOT NULL ORDER BY 1 ASC`, wbsDescQ.params);

        const yearsList = [...new Set(periodRows.map(r => r.period?.split('-')[0]))].filter(Boolean).sort((a, b) => b - a);

        res.status(200).json({
            category_types:   ['All', 'Local Materials'],
            bus:              buRows.map(r => r.bu),
            wbs_types:        wbsTypeRows.map(r => r.wbs_type),
            years:            yearsList,
            periods:          periodRows.map(r => r.period),
            customers:        customerRows.map(r => r.customer),
            loa_ids:          loaIdRows.map(r => r.loa_id),
            loa_names:        loaRows.map(r => r.loa_name),
            wbs:              wbsRows.map(r => r.wbs),
            wbs_descriptions: wbsDescRows.map(r => r.wbs_description)
        });

    } catch (error) {
        console.error("Dashboard Filters Error:", error);
        res.status(500).json({ error: error.message });
    }
};

// ==========================================
// 1. DASHBOARD ANALYTICS SQL GENERATOR (Fixed Postgres MAX(0) Error)
// ==========================================
const getDashboardAnalyticsSQL = (groupByCol, asblCols, ncCols) => {
    const hasAsbl = asblCols !== "0";
    const hasNc = ncCols !== "0"; 
    
    // 1. Safe Subquery Selection
    const asblSubquery = hasAsbl ? `MAX(${asblCols})` : `0`;
    const ncSubquery = hasNc ? `MAX(${ncCols})` : `0`;

    // 2. Safe Middle Aggregation (Prevents Postgres MAX(0) Error!)
    const catAsbl = hasAsbl ? `MAX(COALESCE(static.asbl_val, 0))` : `0`;
    const catNc = hasNc ? `MAX(COALESCE(static.nc_val, 0))` : `0`;

    // 3. Final Outer Selection
    const asblSelect = hasAsbl ? 'ROUND(SUM(cat_asbl), 2)' : '0.00';
    const ncSelect = hasNc ? 'ROUND(SUM(cat_nc), 2)' : '0.00';
    const varSelect = hasAsbl ? 'ROUND(SUM(cat_asbl) - SUM(cat_ptd + cat_oc + cat_nc), 2)' : '0.00';

    const prefixCol = `t.${groupByCol.trim()} AS ${groupByCol.trim()}`;
    const groupByInner = `t.${groupByCol.trim()}`;

    return `
        SELECT 
            ${groupByCol},
            CAST(${asblSelect} AS NUMERIC(15,2)) as asbl,
            ROUND(SUM(cat_ptd), 2) as ptd,
            ROUND(SUM(cat_ptd + cat_oc + cat_nc), 2) as eac,
            ROUND(SUM(cat_oc), 2) as open_commitment,
            CAST(${ncSelect} AS NUMERIC(15,2)) as non_committed,
            CAST(${varSelect} AS NUMERIC(15,2)) as eac_vs_asbl
        FROM (
            SELECT 
                ${prefixCol},
                t."Merged_wbs_categories",
                ${catAsbl} as cat_asbl,
                SUM(t.ptd_val) as cat_ptd,
                SUM(t.oc_val) as cat_oc,
                ${catNc} as cat_nc
            FROM (
                SELECT 
                    ${groupByCol}, "Merged_wbs_categories", 
                    ptd as ptd_val, open_commitment_KEUR as oc_val
                FROM final_dashboard_table
                {{WHERE_CLAUSE}}
            ) as t
            LEFT JOIN (
                SELECT 
                    "Merged_wbs_categories", 
                    ${asblSubquery} as asbl_val, 
                    ${ncSubquery} as nc_val
                FROM final_dashboard_table
                GROUP BY "Merged_wbs_categories"
            ) as static ON t."Merged_wbs_categories" = static."Merged_wbs_categories"
            
            GROUP BY ${groupByInner}, t."Merged_wbs_categories"
        ) as category_rollup
        GROUP BY ${groupByCol}
        ORDER BY ${groupByCol} ASC
    `;
};

// ==============================
// COMMON Dashboard Filters FUNCTION (Fixed for WBS)
// ==============================
const applyDashboardFilters = (query, conditions, params) => {
    buildCommonFilters(query, conditions, params);
};

// ==========================================
// 3. BU ANALYTICS API (Removed Extra wT params)
// ==========================================
exports.getBuAnalytics = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr); // Added ncCols here

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const sql = getDashboardAnalyticsSQL('bu', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql);

        // 🔥 FIX: Passed ONLY baseParams
        const [rows] = await db.query(sql, baseParams);
        res.status(200).json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

// ==========================================
// 4. LOA ANALYTICS API (Removed Extra wT params)
// ==========================================
exports.getLoaAnalytics = async (req, res) => {
    try {
        const { type, allowedCustomers, showAll } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr); // Added ncCols here
        const limitSql = showAll === 'true' ? '' : 'LIMIT 10';

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const rawSql = getDashboardAnalyticsSQL('loa_name', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql);
        
        const sql = `SELECT * FROM (${rawSql}) final_t ORDER BY asbl DESC ${limitSql}`;

        // 🔥 FIX: Passed ONLY baseParams
        const [rows] = await db.query(sql, baseParams);
        res.status(200).json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

// exports.getNonCommittedTrend = async (req, res) => {
//     try {
//         let { loa_name = '', active_inactive = '', wbs_type = '', category_type = '' } = req.query;

//         // 🔥 BACKEND FAILSAFE: Agar frontend galti se Array bhej de, toh uski pehli value uthao (Prevents Postgres Crash)
//         if (Array.isArray(loa_name)) loa_name = loa_name[0];
//         if (Array.isArray(active_inactive)) active_inactive = active_inactive[0];

//         const currentMonthYear = new Date().toLocaleString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');

//         // 1. Array parsing for WBS and Category
//         const wTArr = wbs_type ? String(wbs_type).split(',').map(v => v.trim().toLowerCase()).filter(Boolean) : [];
//         const catArr = category_type ? String(category_type).split(',').map(v => v.trim().toLowerCase()).filter(Boolean) : [];

//         const [rows] = await db.query(
//             `
//             SELECT
//                 latest.month_year,
//                 SUM(latest.new_value) AS total_non_committed
//             FROM
//             (
//                 SELECT l1.*
//                 FROM user_activity_logs l1
//                 INNER JOIN
//                 (
//                     SELECT loa_name, categories, month_year, MAX(id) AS latest_id
//                     FROM user_activity_logs
//                     GROUP BY loa_name, categories, month_year
//                 ) l2 ON l1.id = l2.latest_id
//             ) latest
//             WHERE (? = '' OR latest.loa_name = ?)
//               AND (? = '' OR latest.active_inactive = ?)
//               AND latest.month_year <> ?
//             GROUP BY latest.month_year
//             ORDER BY TO_DATE(CONCAT('01-', latest.month_year), 'DD-Mon-YYYY') DESC
//             LIMIT 6
//             `,
//             [loa_name, loa_name, active_inactive, active_inactive, currentMonthYear]
//         );
//         res.json(rows);
//     } catch (error) {
//         res.status(500).json({ error: error.message });
//     }
// };

// ==========================================
// NON-COMMITTED TREND (Rolling 6 Months - Separated by WBS Type)
// ==========================================
exports.getNonCommittedTrend = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const now = new Date();
        const curMonthName = now.toLocaleString('en-US', { month: 'short' });
        const curYear = now.getFullYear();
        const curMonthYearLog = now.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');

        // 1. Base conditions
        let conditions = ["categories != 'Revenue'"];
        let params = [];
        applyRLS(type, allowedCustomers, conditions, params);

        // 2. Applying Global Filters (Syncing with Sidebar)
        const buArr = getValArray(req.query.bu, req.query, 'bu');
        if (buArr) { conditions.push(`TRIM(LOWER(bu)) IN (?)`); params.push(buArr.map(v => v.trim().toLowerCase())); }

        const custArr = getValArray(req.query.customer, req.query, 'customer');
        if (custArr) { conditions.push(`TRIM(LOWER(customer)) IN (?)`); params.push(custArr.map(v => v.trim().toLowerCase())); }

        const loaIdArr = getValArray(req.query.loa_ids, req.query, 'loa_ids');
        if (loaIdArr) { conditions.push(`TRIM(LOWER(loa_id)) IN (?)`); params.push(loaIdArr.map(v => v.trim().toLowerCase())); }

        const loaNameArr = getValArray(req.query.loa_names, req.query, 'loa_names');
        if (loaNameArr) { conditions.push(`TRIM(LOWER(loa_name)) IN (?)`); params.push(loaNameArr.map(v => v.trim().toLowerCase())); }

        const catArr = getValArray(req.query.category_type, req.query, 'category_type');
        if (catArr) {
            const hasAll = catArr.includes('all');
            const hasLM = catArr.includes('local materials');
            if (hasLM && !hasAll) conditions.push(`TRIM(LOWER(categories)) = 'local materials'`);
            else if (!hasAll && !hasLM) conditions.push(`TRIM(LOWER(categories)) <> 'local materials'`);
        } else {
            conditions.push(`TRIM(LOWER(categories)) <> 'local materials'`);
        }

        const activeArr = getValArray(req.query.active_inactive, req.query, 'active_inactive');
        if (activeArr) {
            if (activeArr.includes('active') && !activeArr.includes('inactive')) conditions.push(`active_inactive = 'Active'`);
            if (activeArr.includes('inactive') && !activeArr.includes('active')) conditions.push(`active_inactive = 'Inactive'`);
        }

        const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const doubleParams = [...params, ...params];

        const sql = `
            WITH rolling_data AS (
                -- 🟢 PART 1: Current Month Live Data (ONLY FINALIZED)
                SELECT 
                    '${curMonthName}' as r_month, 
                    ${curYear} as r_year, 
                    SUM(non_committed_project) as nc_project,
                    SUM(non_committed_amc) as nc_amc,
                    SUM(non_committed_warranty) as nc_warranty
                FROM summary
                ${whereClause} 
                AND active_inactive = 'Active'
                AND EXISTS (
                    SELECT 1 FROM user_activity_logs ual 
                    WHERE ual.loa_id = summary.loa_id 
                    AND ual.categories = summary.categories
                    AND ual.month_year = '${curMonthYearLog}'
                    -- 🔥 CRITICAL FIX: Sirf wahi data dikhao jo Admin finalize kar chuka hai
                    AND ual.is_finalized = true 
                )

                UNION ALL

                -- 🟡 PART 2: Past Data from History Table
                SELECT 
                    report_month as r_month, 
                    report_year as r_year, 
                    SUM(non_committed_project) as nc_project,
                    SUM(non_committed_amc) as nc_amc,
                    SUM(non_committed_warranty) as nc_warranty
                FROM historic_summary
                ${whereClause}
                AND NOT (report_month = '${curMonthName}' AND report_year = ${curYear})
                GROUP BY report_year, report_month
            )
            SELECT 
                CONCAT(r_month, '-', r_year) as month_year,
                ROUND(CAST(COALESCE(nc_project, 0) AS NUMERIC), 2) as nc_project,
                ROUND(CAST(COALESCE(nc_amc, 0) AS NUMERIC), 2) as nc_amc,
                ROUND(CAST(COALESCE(nc_warranty, 0) AS NUMERIC), 2) as nc_warranty
            FROM rolling_data
            
            ORDER BY r_year DESC, TO_DATE(r_month, 'Mon') DESC
            LIMIT 6
        `;

        const [rows] = await db.query(sql, doubleParams);
        res.json(rows.reverse());

    } catch (error) {
        console.error("Trend Data Error:", error.message);
        res.status(500).json({ error: error.message });
    }
};

// We don't need getTrendLoas anymore, but keep it empty to prevent route crashes if mapped
exports.getTrendLoas = async (req, res) => { res.json([]); };

// exports.getTrendLoas = async (req, res) => {
//     try {
//         const [rows] = await db.query(`SELECT DISTINCT loa_name FROM user_activity_logs ORDER BY loa_name`);
//         res.json(rows);
//     } catch (error) { res.status(500).json({ error: error.message }); }
// };

// server/controllers/dataController.js mein is function ko update karein
exports.getTrendLoas = async (req, res) => {
    try {
        const { wbs_type, bu, customer, type, allowedCustomers } = req.query;

        // Multi-select helper
        const getFilterArray = (val) => (!val || val === 'All' || val === 'null') ? [] : val.split(',').map(v => v.trim().toLowerCase());

        const selWT = getFilterArray(wbs_type);
        const selBU = getFilterArray(bu);
        const selCust = getFilterArray(customer);
        const allowed = getFilterArray(allowedCustomers);

        const buildSubQuery = (tableName) => {
            let conds = ["loa_name IS NOT NULL"];
            let subParams = [];
            
            // Value check taaki sirf wahi dikhein jinka trend hai
            const valCol = tableName === 'historic_summary' ? 'non_committed' : 'new_value';
            conds.push(`${valCol} <> 0`);

            if (selWT.length > 0) { conds.push(`LOWER(wbs_type) IN (?)`); subParams.push(selWT); }
            if (selBU.length > 0) { conds.push(`LOWER(bu) IN (?)`); subParams.push(selBU); }
            if (selCust.length > 0) { conds.push(`LOWER(customer) IN (?)`); subParams.push(selCust); }
            
            if (type !== 'super_admin' && allowed.length > 0) {
                conds.push(`LOWER(customer) IN (?)`);
                subParams.push(allowed);
            }
            return { sql: `SELECT DISTINCT loa_name FROM ${tableName} WHERE ${conds.join(' AND ')}`, params: subParams };
        };

        const logs = buildSubQuery('user_activity_logs');
        const hist = buildSubQuery('historic_summary');

        const [rows] = await db.query(
            `SELECT DISTINCT loa_name FROM (${logs.sql} UNION ${hist.sql}) as combined ORDER BY loa_name ASC`,
            [...logs.params, ...hist.params]
        );
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

// ==========================================
// 2. DASHBOARD TABLE SQL GENERATOR (Fixed Postgres MAX(0) Error)
// ==========================================
const getDashboardTableSQL = (groupByCols, asblCols, ncCols) => {
    const hasAsbl = asblCols !== "0";
    const hasNc = ncCols !== "0"; 
    
    // 1. Safe Subquery Selection
    const asblSubquery = hasAsbl ? `MAX(${asblCols})` : `0`;
    // const asblLoaSubquery = hasAsbl ? `MAX(asbl_loa)` : `0`;
    const ncSubquery = hasNc ? `MAX(${ncCols})` : `0`;

    // 2. Safe Middle Aggregation (Prevents Postgres MAX(0) Error!)
    const catAsbl = hasAsbl ? `MAX(COALESCE(static.asbl_val, 0))` : `0`;
    // const catAsblLoa = hasAsbl ? `MAX(COALESCE(static.asbl_loa_val, 0))` : `0`;
    const catNc = hasNc ? `MAX(COALESCE(static.nc_val, 0))` : `0`;

    // 3. Final Outer Selection
    const asblSelect = hasAsbl ? 'ROUND(SUM(cat_asbl), 2)' : '0.00';
    // const asblLoaSelect = hasAsbl ? 'ROUND(SUM(cat_asbl_loa), 2)' : '0.00';
    const ncSelect = hasNc ? 'ROUND(SUM(cat_nc), 2)' : '0.00';
    const varSelect = hasAsbl ? 'ROUND(SUM(cat_asbl) - SUM(cat_ptd + cat_oc + cat_nc), 2)' : '0.00';

    const prefixCols = groupByCols.split(',').map(c => `t.${c.trim()} AS ${c.trim()}`).join(', ');
    const groupByInner = groupByCols.split(',').map(c => `t.${c.trim()}`).join(', ');

    return `
        SELECT 
            ${groupByCols},
            CAST(${asblSelect} AS NUMERIC(15,2)) as asbl,
            
            ROUND(SUM(cat_ptd), 2) as ptd,
            ROUND(SUM(cat_oc), 2) as open_commitment,
            CAST(${ncSelect} AS NUMERIC(15,2)) as non_committed,
            ROUND(SUM(cat_ptd + cat_oc + cat_nc), 2) as eac,
            CAST(${varSelect} AS NUMERIC(15,2)) as eac_vs_asbl
        FROM (
            SELECT 
                ${prefixCols},
                t."Merged_wbs_categories",
                ${catAsbl} as cat_asbl,
                
                SUM(t.ptd_val) as cat_ptd,
                SUM(t.oc_val) as cat_oc,
                ${catNc} as cat_nc
            FROM (
                SELECT 
                    ${groupByCols}, "Merged_wbs_categories", 
                    ptd as ptd_val, open_commitment_KEUR as oc_val
                FROM final_dashboard_table
                {{WHERE_CLAUSE}}
            ) as t
            LEFT JOIN (
                SELECT 
                    "Merged_wbs_categories", 
                    ${asblSubquery} as asbl_val, 
                   
                    ${ncSubquery} as nc_val
                FROM final_dashboard_table
                GROUP BY "Merged_wbs_categories"
            ) as static ON t."Merged_wbs_categories" = static."Merged_wbs_categories"
            
            GROUP BY ${groupByInner}, t."Merged_wbs_categories"
        ) as category_rollup
        GROUP BY ${groupByCols}
    `;
};

// 3. BU Only Table Views (Updated with Dynamic ASBL)
exports.getFinalDashboardTable = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr);

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const sql = getDashboardTableSQL('bu', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql) + " ORDER BY bu ASC";

        const [rows] = await db.query(sql, baseParams);
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getNegativeLOATable = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr);
        const hasAsbl = asblCols !== "0";

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);

        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const varCondition = hasAsbl ? '(SUM(cat_asbl) - SUM(cat_ptd + cat_oc + cat_nc))' : '(0 - SUM(cat_ptd + cat_oc + cat_nc))';

        const sql = getDashboardTableSQL('bu, customer, loa_id, loa_name', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql) + 
                    ` HAVING ${varCondition} < 0 ORDER BY eac_vs_asbl ASC`;

        const [rows] = await db.query(sql, baseParams);
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getCostViewTable = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr);

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const sql = getDashboardTableSQL('bu, customer, loa_id, loa_name', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql) + " ORDER BY asbl DESC";

        const [rows] = await db.query(sql, baseParams); 
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getCustomerViewTable = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr);

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const sql = getDashboardTableSQL('customer', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql) + " ORDER BY asbl DESC";

        const [rows] = await db.query(sql, baseParams); 
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getBuCustomerViewTable = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr);

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const sql = getDashboardTableSQL('bu, customer', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql) + " ORDER BY bu ASC, asbl DESC";

        const [rows] = await db.query(sql, baseParams); 
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getCustomerBuViewTable = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr);

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const sql = getDashboardTableSQL('customer, bu', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql) + " ORDER BY customer ASC";

        const [rows] = await db.query(sql, baseParams); 
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

exports.getCustomerBuLoaViewTable = async (req, res) => {
    try {
        const { type, allowedCustomers } = req.query;
        const wTArr = getValArray(req.query.wbs_type, req.query, 'wbs_type');
        
        const asblCols = getDynamicSumColumns(wTArr, 'asbl');
        const ncCols = getDynamicNCColumns(wTArr);

        let conditions = ["categories NOT IN ('Not to considered')", "cost_revenue = 'Cost'"];
        let baseParams = [];
        applyRLS(type, allowedCustomers, conditions, baseParams);
        applyDashboardFilters(req.query, conditions, baseParams);
        
        const whereSql = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const sql = getDashboardTableSQL('customer, bu, loa_name, loa_id', asblCols, ncCols).replace('{{WHERE_CLAUSE}}', whereSql) + " ORDER BY customer ASC";

        const [rows] = await db.query(sql, baseParams); 
        res.json(rows);
    } catch (error) { res.status(500).json({ error: error.message }); }
};

// exports.getReviewChanges = async (req, res) => {

//     // 🔥 Define monthYear at the top
//     const monthYear = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');
//     try {
//         const { draw, start, length, type, allowedCustomers } = req.query;
//         const startIdx = parseInt(start) || 0;
//         const limitIdx = parseInt(length) || 100;

//         let conditions = [
//             "categories != 'Revenue'",
//             "ABS(COALESCE(non_committed, 0) - COALESCE(non_committed_editable, 0)) > 0.01"
//         ];
//         let params = [];
        
//         // Apply RLS & Shared Filters
//         applyRLS(type, allowedCustomers, conditions, params);
//         buildCommonFilters(req.query, conditions, params);

//         const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';

//         const matrixQuery = `
//             SELECT 
//                 bu, customer, loa_id, loa_name, cost_revenue, categories,
//                 MAX(wbs_type) as wbs_type, -- 🔥 NAYA: wbs_type add kiya taaki mapping sahi ho
//                 MAX(asbl) as asbl, MAX(asbl_loa) as asbl_loa, SUM(ptd) as ptd, 
//                 MAX(open_commitment_KEUR) as open_commitment, 
//                 MAX(non_committed_editable) as non_committed, 
//                 MAX(non_committed) as non_committed_original,
//                 MAX(updated_by) as updated_by,
//                 TO_CHAR(MAX(updated_at), 'DD-Mon-YYYY HH24:MI') as updated_at,
//                 (SUM(ptd) + MAX(open_commitment_KEUR) + MAX(non_committed_editable)) as eac,
//                 (MAX(asbl) - (SUM(ptd) + MAX(open_commitment_KEUR) + MAX(non_committed_editable))) as eac_vs_asbl
//             FROM final_dashboard_table
//             ${whereClause}
//             GROUP BY bu, customer, loa_id, loa_name, cost_revenue, categories
//             -- 🔥 Filter condition updated:
//             HAVING (
//                 ABS(COALESCE(MAX(non_committed), 0) - COALESCE(MAX(non_committed_editable), 0)) > 0.01
//                 OR 
//                 EXISTS (
//                     SELECT 1 FROM user_activity_logs 
//                     WHERE loa_id = final_dashboard_table.loa_id 
//                     AND categories = final_dashboard_table.categories 
//                     AND month_year = ? -- Current Month
//                 )
//             )
//             ORDER BY loa_name ASC, cost_revenue ASC
//         `;

//         const [countRes] = await db.query(`SELECT COUNT(*) as total FROM (${matrixQuery}) as temp`, params);
//         const [dataRows] = await db.query(`${matrixQuery} LIMIT ?, ?`, [...params, startIdx, limitIdx]);

//         res.status(200).json({
//             draw: parseInt(draw) || 0,
//             recordsTotal: parseInt(countRes[0]?.total || 0),
//             recordsFiltered: parseInt(countRes[0]?.total || 0),
//             data: dataRows
//         });
//     } catch (error) { res.status(500).json({ error: error.message }); }
// };

exports.getReviewChanges = async (req, res) => {
    const monthYear = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');
    
    try {
        const { draw, start, length, type, allowedCustomers } = req.query;
        let conditions = ["categories != 'Revenue'"];
        let params = [];
        
        applyRLS(type, allowedCustomers, conditions, params);
        buildCommonFilters(req.query, conditions, params);

        const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';

        const matrixQuery = `
            SELECT 
                bu, customer, loa_id, loa_name, cost_revenue, categories,
                MAX(wbs_type) as wbs_type,
                MAX(asbl) as asbl, MAX(asbl_loa) as asbl_loa, SUM(ptd) as ptd, 
                MAX(open_commitment_KEUR) as open_commitment, 
                MAX(non_committed_editable) as non_committed, 
                MAX(non_committed) as non_committed_original,
                MAX(updated_by) as updated_by,
                TO_CHAR(MAX(updated_at), 'DD-Mon-YYYY HH24:MI') as updated_at,
                (SUM(ptd) + MAX(open_commitment_KEUR) + MAX(non_committed_editable)) as eac,
                (MAX(asbl) - (SUM(ptd) + MAX(open_commitment_KEUR) + MAX(non_committed_editable))) as eac_vs_asbl
            FROM final_dashboard_table fdt
            ${whereClause}
            GROUP BY bu, customer, loa_id, loa_name, cost_revenue, categories
            
            -- 🔥 ULTRA STRICT FILTER: 
            HAVING (
                -- 1. Check Significant Difference (Ignore tiny float/null mismatches)
                (
                    MAX(non_committed_editable) IS NOT NULL 
                    AND ABS(COALESCE(MAX(non_committed), 0) - COALESCE(MAX(non_committed_editable), 0)) > 0.01
                )
                OR 
                -- 2. Strictly check if an Activity Log exists for this EXACT Category + LOA
                EXISTS (
                    SELECT 1 FROM user_activity_logs ual 
                    WHERE ual.loa_id = fdt.loa_id 
                    AND ual.categories = fdt.categories 
                    AND ual.month_year = ? AND ual.is_finalized = false -- 🔥 Sirf pending logs
                )
            )
            ORDER BY loa_name ASC, categories ASC
        `;

        const countParams = [...params, monthYear];
        const [countRes] = await db.query(`SELECT COUNT(*) as total FROM (${matrixQuery}) as temp`, countParams);
        const [dataRows] = await db.query(`${matrixQuery} LIMIT ?, ?`, [...countParams, parseInt(start) || 0, parseInt(length) || 100]);

        res.status(200).json({
            draw: parseInt(draw) || 0,
            recordsTotal: parseInt(countRes[0]?.total || 0),
            recordsFiltered: parseInt(countRes[0]?.total || 0),
            data: dataRows
        });

    } catch (error) {
        console.error("getReviewChanges Error:", error.message);
        res.status(500).json({ error: error.message });
    }
};

// exports.finalizeChanges = async (req, res) => {
//     try {

//         const monthYear = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');

//         // 🔥 Logic Update: Production table ko un sabhi rows ke liye update karein
//         // jo current mahine ke activity logs mein maujood hain
//         const updateCondition = `
//             WHERE EXISTS (
//                 SELECT 1 FROM user_activity_logs ual 
//                 WHERE ual.loa_id = summary.loa_id 
//                 AND ual.categories = summary.categories 
//                 AND ual.month_year = '${monthYear}'
//             ) OR ABS(COALESCE(non_committed, 0) - COALESCE(non_committed_editable, 0)) > 0.01
//         `;
//         // 1. Move editable -> original in Summary Table
//         await db.query(`
//             UPDATE summary 
//             SET non_committed = non_committed_editable,
//                 non_committed_project = non_committed_editable_project,
//                 non_committed_amc = non_committed_editable_amc,
//                 non_committed_warranty = non_committed_editable_warranty
//             WHERE ABS(COALESCE(non_committed, 0) - COALESCE(non_committed_editable, 0)) > 0.01
//         `);

//         // 2. Move editable -> original in Dashboard Table
//         await db.query(`
//             UPDATE final_dashboard_table 
//             SET non_committed = non_committed_editable,
//                 non_committed_project = non_committed_editable_project,
//                 non_committed_amc = non_committed_editable_amc,
//                 non_committed_warranty = non_committed_editable_warranty
//             WHERE ABS(COALESCE(non_committed, 0) - COALESCE(non_committed_editable, 0)) > 0.01
//         `);

//         // 3. 🔥 CRITICAL: Mark logs as Finalized taaki Review page clear ho jaye
//         await db.query(`
//             UPDATE user_activity_logs 
//             SET is_finalized = true 
//             WHERE month_year = ? AND is_finalized = false
//         `, [monthYear]);

//         // 3. Recalculate EAC and Variance globally
//         await db.query(`
//             UPDATE final_dashboard_table 
//             SET eac = (ptd + open_commitment_KEUR + non_committed),
//                 eac_vs_asbl = (asbl - (ptd + open_commitment_KEUR + non_committed))
//         `);

//         // ✅ Trigger Auto Sync engine specifically after Admin Finalizes
//         triggerAutoSync('non_committed_finalized');

//         filterCache.flushAll(); // Flush RAM Cache
//         res.status(200).json({ message: "All changes finalized successfully!" });
//     } catch (error) {
//         console.error("finalizeChanges Error:", error);
//         res.status(500).json({ error: error.message });
//     }
// };

exports.finalizeChanges = async (req, res) => {
    try {
        const now = new Date();
        const month = now.toLocaleString('en-US', { month: 'short' }); // 'Aug'
        const year = now.getFullYear();
        const monthYear = `${month}-${year}`;

        // 1. Move editable -> original in Summary Table (Existing Logic)
        await db.query(`
            UPDATE summary 
            SET non_committed = non_committed_editable,
                non_committed_project = non_committed_editable_project,
                non_committed_amc = non_committed_editable_amc,
                non_committed_warranty = non_committed_editable_warranty
            WHERE ABS(COALESCE(non_committed, 0) - COALESCE(non_committed_editable, 0)) > 0.01
            OR EXISTS (SELECT 1 FROM user_activity_logs ual WHERE ual.loa_id = summary.loa_id AND ual.categories = summary.categories AND ual.is_finalized = false)
        `);

        // 2. Move editable -> original in Dashboard Table (Existing Logic)
        await db.query(`
            UPDATE final_dashboard_table 
            SET non_committed = non_committed_editable,
                non_committed_project = non_committed_editable_project,
                non_committed_amc = non_committed_editable_amc,
                non_committed_warranty = non_committed_editable_warranty
            WHERE ABS(COALESCE(non_committed, 0) - COALESCE(non_committed_editable, 0)) > 0.01
        `);

        // 🔥 2.1 NEW: Save Snapshot to Historic Table
        // Sirf un LOAs ko copy karega jo abhi finalize huye hain
        await db.query(`
            INSERT INTO historic_summary 
            (bu, customer, loa_id, loa_name, cost_revenue, categories, merged_wbs, "Merged_wbs_category", 
             active_inactive, asbl, asbl_amc, asbl_project, asbl_warranty, asbl_loa, ptd, 
             "open_commitment_KEUR", non_committed, non_committed_amc, non_committed_project, 
             non_committed_warranty, eac, eac_vs_asbl, updated_by, updated_at, report_month, report_year)
            SELECT 
                bu, customer, loa_id, loa_name, cost_revenue, categories, merged_wbs, "Merged_wbs_category", 
                active_inactive, asbl, asbl_amc, asbl_project, asbl_warranty, asbl_loa, ptd, 
                "open_commitment_KEUR", non_committed, non_committed_amc, non_committed_project, 
                non_committed_warranty, eac, eac_vs_asbl, updated_by, updated_at, ?, ?
            FROM summary
            WHERE loa_id IN (SELECT DISTINCT loa_id FROM user_activity_logs WHERE is_finalized = false)
            ON CONFLICT (loa_id, categories, report_month, report_year) 
            DO UPDATE SET 
                non_committed = EXCLUDED.non_committed,
                non_committed_amc = EXCLUDED.non_committed_amc,
                non_committed_project = EXCLUDED.non_committed_project,
                non_committed_warranty = EXCLUDED.non_committed_warranty,
                eac = EXCLUDED.eac,
                eac_vs_asbl = EXCLUDED.eac_vs_asbl,
                updated_at = CURRENT_TIMESTAMP
        `, [month, year]);

        // 3. Mark logs as Finalized (Existing Logic)
        await db.query(`UPDATE user_activity_logs SET is_finalized = true WHERE is_finalized = false`);

        // 4. Recalculate EAC and Variance globally (Existing Logic)
        await db.query(`
            UPDATE final_dashboard_table 
            SET eac = (ptd + open_commitment_KEUR + non_committed),
                eac_vs_asbl = (asbl - (ptd + open_commitment_KEUR + non_committed))
        `);

        triggerAutoSync('non_committed_finalized');
        filterCache.flushAll(); 
        res.status(200).json({ message: "All changes finalized and history snapshot saved!" });
    } catch (error) {
        console.error("finalizeChanges Error:", error);
        res.status(500).json({ error: error.message });
    }
};

exports.checkPendingChanges = async (req, res) => {
    const monthYear = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');
    
    try {
        // 🔥 STRICT CHECK: Count only the specific categories that were touched
        const [rows] = await db.query(`
            SELECT COUNT(*) as count 
            FROM (
                SELECT loa_id, categories
                FROM final_dashboard_table fdt
                WHERE categories != 'Revenue'
                GROUP BY loa_id, categories
                HAVING (
                    ABS(COALESCE(MAX(non_committed), 0) - COALESCE(MAX(non_committed_editable), 0)) > 0.01
                    OR 
                    EXISTS (
                        SELECT 1 FROM user_activity_logs ual 
                        WHERE ual.loa_id = fdt.loa_id 
                        AND ual.categories = fdt.categories 
                        AND ual.month_year = ? AND ual.is_finalized = false -- 🔥 Sirf non-finalized dekho
                    )
                )
            ) as touched_rows
        `, [monthYear]);
        
        const count = parseInt(rows[0]?.count || 0);
        res.status(200).json({ count });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
};

exports.getERPResource = async (req, res) => {
  try {
    const page = parseInt(req.query.page) || 1;
    const pageSize = parseInt(req.query.pageSize) || 50;
    const search = req.query.search || '';

    const offset = (page - 1) * pageSize;

    let whereClause = '';
    let params = [];

    if (search) {
      whereClause = `
        WHERE
          tr_global_period ILIKE ?
          OR lm_nokia_id_name ILIKE ?
          OR resource_nokia_id_name ILIKE ?
          OR home_country ILIKE ?
          OR customer_team ILIKE ?
          OR gic_name ILIKE ?
      `;
      const searchValue = `%${search}%`;
      params = [searchValue, searchValue, searchValue, searchValue, searchValue, searchValue];
    }

    const [countRows] = await db.query(
      `SELECT COUNT(*) as total FROM erp_resource ${whereClause}`,
      params
    );

    const totalRecords = parseInt(countRows[0]?.total || 0);

    const [rows] = await db.query(
      `SELECT * FROM erp_resource ${whereClause} ORDER BY id ASC LIMIT ?, ?`,
      [...params, offset, pageSize]
    );

    res.json({
      data: rows,
      totalRecords,
      page,
      pageSize
    });

  } catch (error) {
    console.error(error);
    res.status(500).json({
      message: 'Failed to fetch ERP Resource data'
    });
  }
};

exports.exportERPResource = async (req, res) => {
  try {
    const search = req.query.search || '';
    let whereClause = '';
    let params = [];

    if (search) {
      whereClause = `
        WHERE
          tr_global_period ILIKE ?
          OR lm_nokia_id_name ILIKE ?
          OR resource_nokia_id_name ILIKE ?
          OR home_country ILIKE ?
          OR customer_team ILIKE ?
      `;
      const searchValue = `%${search}%`;
      params = [searchValue, searchValue, searchValue, searchValue, searchValue];
    }

    const [rows] = await db.query(`SELECT * FROM erp_resource ${whereClause} ORDER BY id ASC`, params);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('ERP Resource');
    const today = new Date();
    const formattedDate = today.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).replace(/ /g, '-');

    worksheet.columns = [
      { header: 'ID', key: 'id', width: 10 },
      { header: 'TR Global Period', key: 'tr_global_period', width: 20 },
      { header: 'LM Nokia ID Name', key: 'lm_nokia_id_name', width: 35 },
      { header: 'Home Country', key: 'home_country', width: 20 },
      { header: 'Resource ERP Type', key: 'resource_erp_type', width: 20 },
      { header: 'Resource Person Number', key: 'resource_person_number', width: 20 },
      { header: 'Resource Nokia ID Name', key: 'resource_nokia_id_name', width: 35 },
      { header: 'Time Entry Date', key: 'time_entry_date', width: 20 },
      { header: 'Recorded Hours', key: 'recorded_hours', width: 15 },
      { header: 'Time Entry Status', key: 'time_entry_status', width: 20 },
      { header: 'Daily Working Hours', key: 'daily_working_hours', width: 20 },
      { header: 'TR WBS/Care Contract/Opp', key: 'tr_wbs_care_contract_opp', width: 35 },
      { header: 'TR WBS Description', key: 'tr_wbs_care_contract_opp_description', width: 50 },
      { header: 'SVO ID', key: 'svo_id', width: 20 },
      { header: 'SVO Description', key: 'svo_description', width: 40 },
      { header: 'GIC', key: 'gic', width: 20 },
      { header: 'GIC Name', key: 'gic_name', width: 30 },
      { header: 'Customer Team', key: 'customer_team', width: 30 },
      { header: 'Time Approval Date', key: 'time_approval_date', width: 20 },
      { header: 'LM Email', key: 'lm_email', width: 35 },
      { header: 'Resource Email', key: 'resource_email', width: 35 }
    ];

    rows.forEach(row => worksheet.addRow(row));
    worksheet.getRow(1).font = { bold: true };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=ERP_Resource_${formattedDate}.xlsx`);
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    res.status(500).json({ message: 'Excel export failed' });
  }
};

const formatSqlDate = (val) => {
  if (!val) return null;
  if (val instanceof Date) return val.toISOString().split('T')[0];
  return typeof val === 'string' && val.trim() !== '' ? val : null;
};

const getValueByAliases = (normalizedRow, aliases) => {
  for (const alias of aliases) {
    if (normalizedRow[alias] !== undefined) return normalizedRow[alias];
    const cleanAlias = alias.replace(/[^a-z0-9]/g, '');
    for (const key of Object.keys(normalizedRow)) {
      const cleanKey = key.replace(/[^a-z0-9]/g, '');
      if (cleanKey === cleanAlias) return normalizedRow[key];
    }
  }
  return null;
};

exports.uploadERPResource = async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ message: 'Please upload a file' });

    const workbook = XLSX.readFile(req.file.path, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const created_by = req.body.created_by || (req.user && req.user.email) || 'System';
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    const currentMonth = new Date().toLocaleDateString('en-US', { month: 'short', year: 'numeric' }).replace(' ', '-');

    const erpBatchRows = [];

    for (const row of rows) {
      const normalizedRow = {};
      for (const key of Object.keys(row)) normalizedRow[key.toLowerCase().trim()] = row[key];

      erpBatchRows.push([
        getValueByAliases(normalizedRow, ['tr global period', 'trglobalperiod']),
        getValueByAliases(normalizedRow, ['lm nokia id, name', 'lm nokia id name', 'lmnokiaidname']),
        getValueByAliases(normalizedRow, ['home country', 'homecountry']),
        getValueByAliases(normalizedRow, ['resource erp type', 'resourceerptype']),
        getValueByAliases(normalizedRow, ['resource persn. number', 'resource persn number', 'resource person number', 'resourcepersonnumber']),
        getValueByAliases(normalizedRow, ['resource nokia id, name', 'resource nokia id name', 'resourcenokiaidname']),
        formatSqlDate(getValueByAliases(normalizedRow, ['time entry date', 'timeentrydate'])),
        getValueByAliases(normalizedRow, ['recorded hours', 'recordedhours']),
        getValueByAliases(normalizedRow, ['time entry status', 'timeentrystatus']),
        getValueByAliases(normalizedRow, ['daily working hours', 'dailyworkinghours']),
        getValueByAliases(normalizedRow, ['tr wbs/care contract/opp', 'tr wbs care contract opp', 'trwbscarecontractopp']),
        getValueByAliases(normalizedRow, ['tr wbs/care contract/opp description', 'tr wbs care contract opp description', 'trwbscarecontractoppdescription']),
        getValueByAliases(normalizedRow, ['svo id', 'svoid']),
        getValueByAliases(normalizedRow, ['svo description', 'svodescription']),
        getValueByAliases(normalizedRow, ['gic']),
        getValueByAliases(normalizedRow, ['gic name', 'gicname']),
        getValueByAliases(normalizedRow, ['ct (customer team)', 'ct customer team', 'customer team', 'ct', 'customerteam']),
        formatSqlDate(getValueByAliases(normalizedRow, ['time approval date', 'timeapprovaldate'])),
        getValueByAliases(normalizedRow, ['lm email', 'lmemail']),
        getValueByAliases(normalizedRow, ['resource email', 'resourceemail']),
        currentMonth,
        created_by
      ]);
    }

    if (erpBatchRows.length > 0) {
        await db.query(
            `INSERT INTO erp_resource 
            (tr_global_period, lm_nokia_id_name, home_country, resource_erp_type, resource_person_number, resource_nokia_id_name, time_entry_date, recorded_hours, time_entry_status, daily_working_hours, tr_wbs_care_contract_opp, tr_wbs_care_contract_opp_description, svo_id, svo_description, gic, gic_name, customer_team, time_approval_date, lm_email, resource_email, month, created_by) 
            VALUES ?`,
            [erpBatchRows]
        );
    }

    fs.unlinkSync(req.file.path);
    res.json({ success: true, uploadedRows: rows.length, month: currentMonth });

  } catch (err) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ message: 'Upload failed: ' + err.message });
  }
};

exports.runBGDMAlertsCore = async () => {
    try {
        const [rows] = await db.query(`
            SELECT bu, customer, 
            SUM(asbl) as total_asbl, 
            SUM(ptd) as total_ptd, 
            SUM(ptd + open_commitment_KEUR + non_committed_editable) as total_eac
            FROM final_dashboard_table
            WHERE categories != 'Revenue' AND active_inactive = 'Active'
            GROUP BY bu, customer
        `);

        let mailsSent = 0;
        for (let row of rows) {
            const ptdPerc = row.total_asbl > 0 ? (row.total_ptd / row.total_asbl) * 100 : 0;
            const eacPerc = row.total_asbl > 0 ? (row.total_eac / row.total_asbl) * 100 : 0;

            if (ptdPerc > 80 || eacPerc > 100) {
                const [bgdmUsers] = await db.query(`
                    SELECT a.email, u.user_role FROM access a
                    JOIN users u ON a.email = u.email
                    WHERE LOWER(TRIM(a.customer)) = LOWER(TRIM(?)) AND u.user_role = 'BGDM'
                `, [row.customer]);

                for (let user of bgdmUsers) {
                    await mailService.sendCustomerUtilizationAlert(
                        { email: user.email, role: 'Business Group Delivery Manager' },
                        { bu: row.bu, customer: row.customer, ptdPerc: ptdPerc.toFixed(1), eacPerc: eacPerc.toFixed(1) }
                    );
                    mailsSent++;
                }
            }
        }
        return mailsSent;
    } catch (error) {
        throw error;
    }
};

// Existing API wrapper (for manual button click)
exports.triggerCustomerAlerts = async (req, res) => {
    try {
        const count = await exports.runBGDMAlertsCore();
        res.status(200).json({ success: true, message: `Alerts sent to ${count} BGDMs.` });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
};