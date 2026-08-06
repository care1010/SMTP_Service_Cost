const db = require('../config/db');
const bcrypt = require('bcrypt');

// We will just print the email to the local terminal.

exports.getDropdownData = async (req, res) => {
    try {
        const [customerRows] = await db.query("SELECT DISTINCT customer_name FROM customer WHERE is_active = 1 ORDER BY customer_name");
        const [buRows] = await db.query("SELECT DISTINCT bu FROM wbs_loa_id_mapping1 WHERE bu IS NOT NULL ORDER BY bu");
        const [loaRows] = await db.query("SELECT DISTINCT loa_name FROM wbs_loa_id_mapping1 WHERE loa_name IS NOT NULL ORDER BY loa_name");

        res.status(200).json({
            customers: customerRows.map(r => r.customer_name),
            bus: buRows.map(r => r.bu),
            loas: loaRows.map(r => r.loa_name)
        });
    } catch (err) { res.status(500).json({ error: err.message }); }
};

exports.submitRequest = async (req, res) => {
    const { email, password, customers, bu, projectName } = req.body;
    try {
        const [existingUser] = await db.query("SELECT id FROM users WHERE email = ?", [email]);
        if (existingUser.length > 0) return res.status(400).json({ error: "User already exists in the system." });

        const [existingReq] = await db.query("SELECT id FROM access_requests WHERE email = ? AND status = 'Pending'", [email]);
        if (existingReq.length > 0) return res.status(400).json({ error: "A request is already pending for this email." });

        const hashedPassword = await bcrypt.hash(password, 10);
        const customerString = customers.join('|||'); 

        await db.query(
            "INSERT INTO access_requests (email, password, requested_customers, bu, project_name) VALUES (?, ?, ?, ?, ?)",
            [email, hashedPassword, customerString, bu, projectName]
        );

        // 🛡️ 100% SECURE LOCAL LOGGING (NO NETWORK CALLS)
        console.log("\n=======================================================");
        console.log("📧 [MOCK EMAIL] - NEW ACCESS REQUEST");
        console.log("=======================================================");
        console.log(`TO      : admin@nokia.com`);
        console.log(`SUBJECT : Action Required: New Access Request Pending`);
        console.log(`BODY    :`);
        console.log(`Hello Admin,\n`);
        console.log(`A new access request has been submitted by: ${email}`);
        console.log(`Requested Customers: ${customers.join(', ')}`);
        console.log(`\nPlease log in to the portal to Approve or Decline this request.`);
        console.log("=======================================================\n");

        res.status(200).json({ message: "Request submitted successfully!" });
    } catch (err) { 
        res.status(500).json({ error: err.message }); 
    }
};

exports.getPendingRequests = async (req, res) => {
    try {
        const [rows] = await db.query("SELECT id, email, requested_customers, bu, project_name, created_at FROM access_requests WHERE status = 'Pending' ORDER BY created_at DESC");
        res.status(200).json(rows);
    } catch (err) { res.status(500).json({ error: err.message }); }
};

exports.approveRequest = async (req, res) => {
    const { id } = req.body;
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();

        const [reqRows] = await connection.query("SELECT * FROM access_requests WHERE id = ?", [id]);
        if (reqRows.length === 0) throw new Error("Request not found");
        
        const request = reqRows[0];
        const customerList = request.requested_customers.split('|||');

        await connection.query("INSERT INTO users (email, password, type) VALUES (?, ?, 'user')", [request.email, request.password]);

        const accessMapping = customerList.map(c => [c, request.email]);
        if (accessMapping.length > 0) {
            await connection.query("INSERT INTO access (customer, email) VALUES ?", [accessMapping]);
        }

        await connection.query("UPDATE access_requests SET status = 'Approved' WHERE id = ?", [id]);

        await connection.commit();
        res.status(200).json({ message: "Access Approved and Account Created!" });
    } catch (err) {
        await connection.rollback();
        res.status(500).json({ error: err.message });
    } finally {
        connection.release();
    }
};

exports.declineRequest = async (req, res) => {
    try {
        const { id } = req.body;
        await db.query("UPDATE access_requests SET status = 'Declined' WHERE id = ?", [id]);
        res.status(200).json({ message: "Request Declined." });
    } catch (err) { res.status(500).json({ error: err.message }); }
};