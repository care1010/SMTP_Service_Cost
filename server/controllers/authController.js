const db = require('../config/db');
const bcrypt = require('bcrypt');

exports.login = async (req, res) => {

    console.log("=================================");
    console.log("🔵 LOGIN REQUEST RECEIVED");
    // console.log("Username:", username);
    console.log("=================================");
    console.log("🚀 1. Login route was hit!");

    const { email, password } = req.body;
    console.log("📧 2. Email received:", email);

    try {
        console.log("⏳ 3. Querying database for user...");

        const [userRows] = await db.query(
            "SELECT * FROM users WHERE email = ?",
            [email]
        );

        console.log("✅ 4. Database responded. Rows found:", userRows.length);

        if (userRows.length === 0) {
            console.log("❌ 5. User not found. Sending 401.");
            return res.status(401).json({ error: "Invalid Email or Password" });
        }

        const user = userRows[0];
        console.log("⏳ 6. Comparing bcrypt passwords...");

        const isMatch = await bcrypt.compare(password, user.password);

        if (!isMatch) {
            console.log("❌ 7. Password mismatch. Sending 401.");
            return res.status(401).json({ error: "Invalid Email or Password" });
        }

        console.log("✅ 8. Password matched! Fetching access rows...");

        const [accessRows] = await db.query(
            "SELECT customer FROM access WHERE email = ?",
            [email]
        );

        const allowedCustomers = accessRows.map(row => row.customer);
        console.log("✅ 9. Access fetched. Sending success response.");

        res.status(200).json({
            message: "Login Successful",
            user: {
                email: user.email,
                type: user.type,
                allowedCustomers
            }
        });

        console.log("🏁 10. Response sent successfully!");

    } catch (error) {
        console.error("🔥 Auth Error Catch Block:", error);
        res.status(500).json({ error: "Database connection error" });
    }
};