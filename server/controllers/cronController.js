const db = require('../config/db');
const cron = require('node-cron');
const dataController = require('./dataController');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');

let currentCronJob = null;
let monthlyBackupJob = null; // Backup cron job reference
let isSyncing = false;
let autoSyncTimeout = null;
let bgdmAlertJob = null; // 🔥 Naya job reference

// ==========================================
// 📦 DATABASE BACKUP ENGINE (PostgreSQL)
// ==========================================
const runDatabaseBackup = () => {
    console.log("📦 CRON: Starting Monthly Database Backup...");

    // 1. Create Backup Directory (OS Independent: Windows/Linux)
    // Yeh server/controllers folder se 2 level up jayega: 7_Service_Cost_Tracker_Postgres/database/backup
    const backupDir = path.join(__dirname, '../../database/backup');
    if (!fs.existsSync(backupDir)) {
        fs.mkdirSync(backupDir, { recursive: true });
    }

    // 2. Format Date: dd-mmm-yyyy (e.g., 01-Aug-2026)
    const date = new Date();
    const day = String(date.getDate()).padStart(2, '0');
    const month = date.toLocaleString('en-US', { month: 'short' });
    const year = date.getFullYear();
    const fileName = `Backup_ServiceCost_${day}-${month}-${year}.backup`;
    const filePath = path.join(backupDir, fileName);

    // 3. Get DB Credentials
    const host = process.env.DB_HOST || 'localhost';
    const port = process.env.DB_PORT || 5432;
    const user = process.env.DB_USER || 'postgres';
    const password = process.env.DB_PASSWORD || 'postgres';
    const dbName = process.env.DB_NAME || 'service_cost';

    // 4. Build pg_dump Command
    // -F c = Custom format (Compressed, best for 1M+ rows)
    const cmd = `pg_dump -h ${host} -p ${port} -U ${user} -F c -d ${dbName} -f "${filePath}"`;

    // 5. Execute Command (Passing Password securely via Env Variables)
    exec(cmd, { env: { ...process.env, PGPASSWORD: password } }, (error, stdout, stderr) => {
        if (error) {
            console.error(`❌ CRON Backup Failed: ${error.message}`);
            return;
        }
        console.log(`✅ CRON Backup Successful! File saved at: ${filePath}`);
    });
};


// ==========================================
// 🔄 DATA SYNC ENGINE
// ==========================================
const runSync = async (triggeredBy = 'cron') => {
    if (isSyncing) {
        console.log(`⏳ Sync already in progress. Skipping trigger from: ${triggeredBy}`);
        return;
    }

    isSyncing = true;
    console.log(`🟢 CRON/TRIGGER: Starting sync (Triggered by: ${triggeredBy})`);

    try {
        await db.query("UPDATE cron_config SET last_run_at = NOW(), last_run_status = 'running', last_run_message = 'Sync in progress' WHERE job_name = 'full_sync'");
        
        await dataController.runFullSyncCore();

        await db.query("UPDATE cron_config SET last_run_status = 'success', last_run_message = 'Sync completed successfully', run_count = run_count + 1 WHERE job_name = 'full_sync'");
        console.log('✅ CRON/TRIGGER: Sync completed successfully!');
    } catch (error) {
        console.error('❌ CRON/TRIGGER Error:', error.message);
        await db.query("UPDATE cron_config SET last_run_status = 'error', last_run_message = ? WHERE job_name = 'full_sync'", [error.message.substring(0, 250)]);
    } finally {
        isSyncing = false;
    }
};

// 🔥 NAYA: Alerts Runner
const runMonthlyAlerts = async () => {
    console.log("🚀 CRON: Starting Monthly BGDM Alerts...");
    try {
        await db.query("UPDATE cron_config SET last_run_at = NOW(), last_run_status = 'running' WHERE job_name = 'bgdm_alerts'");
        
        const mailsSent = await dataController.runBGDMAlertsCore();

        await db.query("UPDATE cron_config SET last_run_status = 'success', last_run_message = ?, run_count = run_count + 1 WHERE job_name = 'bgdm_alerts'", 
        [`Sent ${mailsSent} alerts successfully`]);
        console.log(`✅ CRON Alerts: Completed. Mails sent: ${mailsSent}`);
    } catch (error) {
        console.error('❌ CRON Alerts Error:', error.message);
        await db.query("UPDATE cron_config SET last_run_status = 'error', last_run_message = ? WHERE job_name = 'bgdm_alerts'", [error.message.substring(0, 250)]);
    }
};


// ==========================================
// ⏰ INITIALIZATION (Sync + Backup)
// ==========================================
exports.initCron = async () => {
    try {
        // --- 1. INITIALIZE DATA SYNC CRON ---
        if (currentCronJob) {
            currentCronJob.stop();
            currentCronJob = null;
        }

        const [rows] = await db.query("SELECT * FROM cron_config WHERE job_name = 'full_sync'");
        if (rows.length > 0) {
            const config = rows[0];
            if (config.is_enabled) {
                currentCronJob = cron.schedule(config.cron_expression, () => {
                    runSync('scheduled_cron');
                });
                console.log(`⏰ Data Sync Cron initialized: ${config.cron_expression}`);
            } else {
                console.log('⏰ Data Sync Cron is currently disabled in DB.');
            }
        }

        // --- 2. INITIALIZE MONTHLY BACKUP CRON ---
        if (monthlyBackupJob) {
            monthlyBackupJob.stop();
        }
        
        // 0 0 1 * * = At 12:00 AM, on day 1 of the month
        monthlyBackupJob = cron.schedule('10 14 20 8 *', () => {
            runDatabaseBackup();
        });
        console.log(`⏰ Monthly DB Backup Cron initialized: 0 0 1 * * (1st of every month at midnight)`);

        // --- 3. 🔥 NAYA: MONTHLY BGDM ALERTS CRON ---
        if (bgdmAlertJob) bgdmAlertJob.stop();

        const [alertRows] = await db.query("SELECT * FROM cron_config WHERE job_name = 'bgdm_alerts'");
        if (alertRows.length > 0 && alertRows[0].is_enabled) {
            bgdmAlertJob = cron.schedule(alertRows[0].cron_expression, () => {
                runMonthlyAlerts();
            });
            console.log(`⏰ BGDM Alerts Cron initialized: ${alertRows[0].cron_expression}`);
        }

    } catch (err) {
        console.error("Cron Init Error:", err);
    }
};

// ==========================================
// 🎛️ API ENDPOINTS
// ==========================================
exports.getCronConfig = async (req, res) => {
    try {
        const [rows] = await db.query("SELECT * FROM cron_config WHERE job_name = 'full_sync'");
        res.json(rows[0] || null);
    } catch (err) { res.status(500).json({ error: err.message }); }
};

exports.getSyncStatus = (req, res) => {
    res.json({ isRunning: isSyncing, cronActive: currentCronJob !== null });
};

exports.updateCronConfig = async (req, res) => {
    try {
        const { cron_expression, is_enabled } = req.body;
        const updates = [];
        const params = [];

        if (cron_expression !== undefined) {
            if (cron_expression !== 'custom' && !cron.validate(cron_expression)) {
                return res.status(400).json({ error: 'Invalid cron expression format' });
            }
            updates.push(`cron_expression = ?`);
            params.push(cron_expression);
        }
        if (is_enabled !== undefined) {
            updates.push(`is_enabled = ?`);
            params.push(is_enabled);
        }

        if (updates.length > 0) {
            params.push('full_sync'); 
            await db.query(`UPDATE cron_config SET ${updates.join(', ')}, updated_at = NOW() WHERE job_name = ?`, params);
            await exports.initCron();
        }

        const [rows] = await db.query("SELECT * FROM cron_config WHERE job_name = 'full_sync'");
        res.json({ message: "Settings updated successfully", config: rows[0] });
    } catch (err) { 
        res.status(500).json({ error: err.message }); 
    }
};

exports.triggerManualSync = async (req, res) => {
    if (isSyncing) return res.status(400).json({ message: "Sync is already running." });
    runSync(req.body.triggeredBy || 'manual_trigger');
    res.json({ message: "Sync started in background." });
};

exports.triggerAutoSync = (source) => {
    if (autoSyncTimeout) clearTimeout(autoSyncTimeout);
    console.log(`⏳ Auto-Sync queued by [${source}]. Waiting 2 seconds...`);
    
    autoSyncTimeout = setTimeout(() => {
        runSync(`auto_trigger_${source}`);
    }, 2000); 
};

// EXPORTING BACKUP FUNCTION SO YOU CAN TEST IT IF NEEDED
exports.testDatabaseBackup = (req, res) => {
    runDatabaseBackup();
    res.json({ message: "Backup process triggered in background. Check console logs." });
};