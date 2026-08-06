const { Pool } = require('pg');
const pgFormat = require('pg-format');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '../.env') });

// ═══════════════════════════════════════════════════════════════
// ENTERPRISE POOL CONFIG
// Paradox: zyada connections ≠ better performance
// Postgres ek time pe ~4-8 queries efficiently run karta hai
// 150 connections = 150 idle threads = memory waste + contention
// ═══════════════════════════════════════════════════════════════
const pool = new Pool({
    host: process.env.DB_HOST || 'localhost',
    user: process.env.DB_USER || 'postgres',
    password: process.env.DB_PASSWORD || 'postgres',
    database: process.env.DB_NAME || 'service_cost',
    port: parseInt(process.env.DB_PORT) || 5432,

    // 🔥 Optimal pool size = (CPU cores * 2) + disk spindles
    // Local dev: 10-20 sufficient; production: 20-30 max
    max: 20,
    min: 2,                        // Always keep 2 warm connections ready
    idleTimeoutMillis: 30000,      // 30s idle = close connection
    connectionTimeoutMillis: 8000, // 8s max wait for a free connection
    maxUses: 7500,                 // Recycle connection after 7500 queries (prevent leaks)

    // Keep connections alive (prevents "connection terminated unexpectedly")
    keepAlive: true,
    keepAliveInitialDelayMillis: 10000,
});

// Pool health monitoring
pool.on('connect', () => {
    // console.log(`🟢 Pool: ${pool.totalCount} total, ${pool.idleCount} idle, ${pool.waitingCount} waiting`);
});

pool.on('error', (err) => {
    console.error('⚡ PostgreSQL Pool Error (Auto-Reconnect):', err.message);
    // Do NOT exit process — pool auto-reconnects
});

pool.on('remove', () => {
    // console.log('🔴 Connection removed from pool');
});

// ═══════════════════════════════════════════════════════════════
// QUERY PARSER (1:1 Value Mapping — prevents Parameter Count Mismatch)
// ═══════════════════════════════════════════════════════════════
const formatQueryToPostgres = (sql, params = []) => {
    let text = sql.replace(/`/g, '');
    text = text.replace(/"?Merged_wbs_categories"?/gi, '"Merged_wbs_categories"');
    text = text.replace(/"?Merged_wbs_category"?/gi, '"Merged_wbs_category"');
    text = text.replace(/"?open_commitment_KEUR"?/gi, '"open_commitment_KEUR"');
    text = text.replace(/\bREGEXP\b/gi, '~');

    let newParams = [...params];

    if (/VALUES\s*\?/i.test(text) && Array.isArray(newParams) && newParams.length > 0 && Array.isArray(newParams[0])) {
        const bulkValues = newParams[0];
        const formattedValues = pgFormat('VALUES %L', bulkValues);
        text = text.replace(/VALUES\s*\?/i, formattedValues);
        newParams = newParams.slice(1);
    }

    if (/LIMIT\s*\?\s*,\s*\?/i.test(text)) {
        const limitVal = parseInt(newParams.pop()) || 100;
        const offsetVal = parseInt(newParams.pop()) || 0;
        text = text.replace(/LIMIT\s*\?\s*,\s*\?/i, `LIMIT ${limitVal} OFFSET ${offsetVal}`);
    }

    const flatParams = [];
    let index = 1;
    let inSingleQuote = false;
    let paramIdx = 0;
    let formattedSql = '';

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        if (char === "'") {
            if (i + 1 < text.length && text[i + 1] === "'") {
                formattedSql += "''";
                i++;
                continue;
            }
            inSingleQuote = !inSingleQuote;
            formattedSql += char;
        } else if (char === '?' && !inSingleQuote) {
            const val = newParams[paramIdx++];
            if (Array.isArray(val)) {
                if (val.length === 0) {
                    formattedSql += 'NULL';
                } else {
                    const placeholders = [];
                    for (const v of val) {
                        placeholders.push(`$${index++}`);
                        flatParams.push(v);
                    }
                    formattedSql += placeholders.join(', ');
                }
            } else {
                formattedSql += `$${index++}`;
                flatParams.push(val);
            }
        } else {
            formattedSql += char;
        }
    }

    return { sql: formattedSql, params: flatParams };
};

// ═══════════════════════════════════════════════════════════════
// QUERY EXECUTOR with per-query timeout
// statement_timeout: 25s max — kills runaway queries automatically
// ═══════════════════════════════════════════════════════════════
const db = {
    query: async (sql, params = []) => {
        const { sql: pgSql, params: pgParams } = formatQueryToPostgres(sql, params);
        const client = await pool.connect();
        try {
            // Set 25s statement timeout — kills slow queries before pool exhaustion
            await client.query('SET statement_timeout = 60000');
            const res = await client.query(pgSql, pgParams);
            const rows = res.rows;
            rows.affectedRows = res.rowCount;
            rows.insertId = res.rows[0]?.id || null;
            return [rows, res.fields || []];
        } finally {
            client.release(); // ALWAYS release — even on error
        }
    },

    getConnection: async () => {
        const client = await pool.connect();
        // Set timeout for transactional connections too
        await client.query('SET statement_timeout = 60000'); // 60s for bulk ops
        return {
            query: async (sql, params = []) => {
                const { sql: pgSql, params: pgParams } = formatQueryToPostgres(sql, params);
                const res = await client.query(pgSql, pgParams);
                const rows = res.rows;
                rows.affectedRows = res.rowCount;
                rows.insertId = res.rows[0]?.id || null;
                return [rows, res.fields || []];
            },
            beginTransaction: () => client.query('BEGIN'),
            commit: () => client.query('COMMIT'),
            rollback: () => client.query('ROLLBACK'),
            release: () => client.release(),
        };
    },

    // Pool health check endpoint ke liye
    getPoolStatus: () => ({
        total: pool.totalCount,
        idle: pool.idleCount,
        waiting: pool.waitingCount,
    }),
};

module.exports = db;