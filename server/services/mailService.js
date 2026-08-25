const transporter = require("../config/mailer");

// Tool configuration

const TOOL_NAME = "NI INDIA Financial Cost Tracker";

const TOOL_LINK = "http://10.68.32.163:3000/";



const sendAccessRequestMail = async (request) => {

    const mailOptions = {

        from: '"NI INDIA Financial Cost Tracker" <care.ni_india@nokia.com>',

        to: ["shraddha.dubey@nokia.com", "neha.sain.ext@nokia.com", "mohsin.1.khan.ext@nokia.com"],

        replyTo: request.email,

        subject: `New Access Request: NI INDIA Financial Cost Tracker - ${request.customer}`,

        html: `

        <div style="font-family: Calibri, Arial, sans-serif; max-width:700px; margin:auto; border:1px solid #dcdcdc; border-radius:8px; overflow:hidden;">

            <div style="background:#124191; color:#ffffff; padding:16px 24px;">

                <h2 style="margin:0;">NI INDIA Financial Cost Tracker</h2>

                <p style="margin:6px 0 0;">New Access Request</p>

            </div>

            <div style="padding:24px; color:#333333;">

                <p>Hello Team,</p>

                <p>A new access request is raised for <strong>NI INDIA Financial Cost Tracker</strong>.</p>



                <table style="width:100%; border-collapse:collapse; margin-top:20px;">

                    <tr>

                        <td style="padding:10px; border:1px solid #ddd; background:#f5f5f5; width:35%;"><strong>Customer Account Name</strong></td>

                        <td style="padding:10px; border:1px solid #ddd;">${request.customer}</td>

                    </tr>

                    <tr>

                        <td style="padding:10px; border:1px solid #ddd; background:#f5f5f5;"><strong>Business Unit</strong></td>

                        <td style="padding:10px; border:1px solid #ddd;">${request.bu}</td>

                    </tr>

                    <tr>

                        <td style="padding:10px; border:1px solid #ddd; background:#f5f5f5;"><strong>Project Name</strong></td>

                        <td style="padding:10px; border:1px solid #ddd;">${request.loa}</td>

                    </tr>

                    <tr>

                        <td style="padding:10px; border:1px solid #ddd; background:#f5f5f5;"><strong>Requested By</strong></td>

                        <td style="padding:10px; border:1px solid #ddd;">${request.email}</td>

                    </tr>

                </table>



                <div style="margin-top:30px; text-align:center;">

                    <p style="font-size:15px; color:#666;">Click the Link below to review the request:</p>

                    <a href="${TOOL_LINK}" style="background:#124191; color:#ffffff; padding:12px 25px; text-decoration:none; font-weight:bold; border-radius:5px; display:inline-block;">

                        Go to ${TOOL_NAME}

                    </a>

                </div>



                <p style="margin-top:25px;">Best Regards,<br><strong>NI INDIA PMO Team</strong></p>

            </div>

            <div style="background:#f8f8f8; padding:12px 24px; font-size:12px; color:#666; text-align:center;">

                NOTE: This is an automatically generated email. Please do not reply directly to this message.

            </div>

        </div>`

    };

    return transporter.sendMail(mailOptions);

};



const sendApprovalMail = async (request) => {

    const mailOptions = {

        from: '"NI INDIA Financial Cost Tracker" <care.ni_india@nokia.com>',

        to: request.email,

        cc: ["shraddha.dubey@nokia.com", "neha.sain.ext@nokia.com", "mohsin.1.khan.ext@nokia.com"],

        subject: `Access Approved - NI INDIA Financial Cost Tracker`,

        html: `

        <div style="font-family: Calibri, Arial, sans-serif; max-width:700px; margin:auto; border:1px solid #dcdcdc; border-radius:8px; overflow:hidden;">

            <div style="background:#124191; color:#ffffff; padding:16px 24px;">

                <h2 style="margin:0;">Access Granted!</h2>

            </div>

            <div style="padding:24px; color:#333333;">

                <p>Hello,</p>

                <p>Your access request for the <strong>NI INDIA Financial Cost Tracker</strong> has been <strong>Approved</strong>.</p>

                <p>Click the link below to access:</p>

               

                <div style="margin:25px 0; text-align:center;">

                    <a href="${TOOL_LINK}" style="background:#124191; color:#ffffff; padding:14px 30px; text-decoration:none; font-weight:bold; border-radius:5px; display:inline-block; font-size:16px;">

                        Login to ${TOOL_NAME}

                    </a>

                </div>



                <table style="width:100%; border-collapse:collapse; margin-top:20px;">

                    <tr>

                        <td style="padding:10px; border:1px solid #ddd; background:#f5f5f5; width:35%;"><strong>Customer Name</strong></td>

                        <td style="padding:10px; border:1px solid #ddd;">${request.requested_customers}</td>

                    </tr>

                    <tr>

                        <td style="padding:10px; border:1px solid #ddd; background:#f5f5f5;"><strong>Business Unit</strong></td>

                        <td style="padding:10px; border:1px solid #ddd;">${request.bu}</td>

                    </tr>

                </table>

                <p style="margin-top:25px;">Best Regards,<br><strong>NI INDIA PMO Team</strong></p>

            </div>

        </div>`

    };

    return transporter.sendMail(mailOptions);

};



const sendDeclineMail = async (request) => {

    const mailOptions = {

        from: '"NI INDIA Financial Cost Tracker" <care.ni_india@nokia.com>',

        to: request.email,

        cc: ["shraddha.dubey@nokia.com", "neha.sain.ext@nokia.com", "mohsin.1.khan.ext@nokia.com"],

        subject: `Access Request Update - NI INDIA Financial Cost Tracker`,

        html: `

        <div style="font-family: Calibri, Arial, sans-serif; max-width:700px; margin:auto; border:1px solid #dcdcdc; border-radius:8px; overflow:hidden;">

            <div style="background:#666666; color:#ffffff; padding:16px 24px;">

                <h2 style="margin:0;">Request Declined</h2>

            </div>

            <div style="padding:24px; color:#333333;">

                <p>Hello,</p>

                <p>We regret to inform you that your access request for the following entity has been <strong>Declined</strong>.</p>

                <table style="width:100%; border-collapse:collapse; margin:20px 0;">

                    <tr>

                        <td style="padding:10px; border:1px solid #ddd; background:#f5f5f5; width:35%;"><strong>Customer Account</strong></td>

                        <td style="padding:10px; border:1px solid #ddd;">${request.requested_customers}</td>

                    </tr>

                </table>

                <p>If you believe this is an error, please visit the portal to re-apply or contact the administrators.</p>

               

                <p style="margin-top:20px;">

                    Portal Link: <a href="${TOOL_LINK}" style="color:#124191; font-weight:bold;">${TOOL_NAME}</a>

                </p>



                <p style="margin-top:25px;">Best Regards,<br><strong>NI INDIA PMO Team</strong></p>

            </div>

        </div>`

    };

    return transporter.sendMail(mailOptions);

};

const sendOTPMail = async (email, otp) => {
    const mailOptions = {
        from: '"NI INDIA Financial Cost Tracker" <care.ni_india@nokia.com>',
        to: email,
        subject: `Password Reset OTP - NI INDIA Financial Cost Tracker`,
        html: `
        <div style="font-family: Calibri, Arial, sans-serif; max-width:600px; margin:auto; border:1px solid #eee; border-radius:10px; padding:20px;">
            <h2 style="color: #124191;">Password Reset Request</h2>
            <p>Hello,</p>
            <p>You requested to reset your password. Use the following OTP to proceed. This OTP is valid for 10 minutes only.</p>
            <div style="background: #f4f4f4; padding: 15px; text-align: center; font-size: 24px; font-weight: bold; letter-spacing: 5px; color: #124191; border-radius: 8px;">
                ${otp}
            </div>
            <p style="margin-top: 20px;">If you did not request this, please ignore this email or contact admin.</p>
            <p>Regards,<br><strong>NI INDIA PMO Team</strong></p>
        </div>`
    };
    return transporter.sendMail(mailOptions);
};


const sendCustomerUtilizationAlert = async (recipient, data) => {
    const mailOptions = {
        from: '"NI INDIA Cost Tracker Alert" <care.ni_india@nokia.com>',
        to: recipient.email,
        subject: `⚠️ URGENT: Delivery Alert - Budget Threshold Exceeded (${data.customer})`,
        html: `
        <div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:auto; border:1px solid #ddd; border-radius:10px; overflow:hidden;">
            <div style="background:#005AFF; color:#ffffff; padding:20px;">
                <h2 style="margin:0;">Stakeholder Financial Alert</h2>
                <p style="margin:5px 0 0;">Role: Business Group Delivery Manager (BGDM)</p>
            </div>
            <div style="padding:25px; color:#333333;">
                <p>Dear Team,</p>
                <p>This is an automated notification regarding the financial health of your assigned customer account. The following metrics have crossed the defined safety thresholds:</p>
                
                <table style="width:100%; margin:20px 0; border-collapse:collapse; background:#f9f9f9; border: 1px solid #eee;">
                    <tr>
                        <td style="padding:12px; border-bottom:1px solid #eee; font-weight:bold;">Customer:</td>
                        <td style="padding:12px; border-bottom:1px solid #eee;">${data.customer}</td>
                    </tr>
                    <tr>
                        <td style="padding:12px; border-bottom:1px solid #eee; font-weight:bold;">Business Unit:</td>
                        <td style="padding:12px; border-bottom:1px solid #eee;">${data.bu}</td>
                    </tr>
                    
                    <tr style="background:#fff1f1;">
                        <td style="padding:12px; border-bottom:1px solid #eee; font-weight:bold; color:#d32f2f;">PTD Utilization:</td>
                        <td style="padding:12px; border-bottom:1px solid #eee; color:#d32f2f; font-weight:bold;">${data.ptdPerc}%</td>
                    </tr>
                    <tr style="background:#fff1f1;">
                        <td style="padding:12px; border-bottom:1px solid #eee; font-weight:bold; color:#d32f2f;">EAC vs ASBL:</td>
                        <td style="padding:12px; border-bottom:1px solid #eee; color:#d32f2f; font-weight:bold;">${data.eacPerc}%</td>
                    </tr>
                </table>

                <p style="font-size:14px; line-height:1.6;"><strong>Action Required:</strong> Please coordinate with the Project Managers (PMs) to review the 'Non-Committed' cost entries and ensure that the project is within the approved budget (ASBL).</p>
                
                <div style="text-align:center; margin:35px 0;">
                    <a href="${TOOL_LINK}" style="background:#124191; color:#ffffff; padding:14px 35px; text-decoration:none; font-weight:bold; border-radius:8px; display:inline-block; font-size:16px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
                        Login to Tool
                    </a>
                </div>

                <p style="margin-top:30px; border-top:1px solid #eee; pt-15px;">Best Regards,<br><strong>NI INDIA Financial Control Team</strong></p>
            </div>
            <div style="background:#f4f4f4; padding:15px; text-align:center; font-size:11px; color:#999;">
                This is a system-generated alert for BGDM role only. Please do not reply to this mailbox.
            </div>
        </div>`
    };
    return transporter.sendMail(mailOptions);
};

// const sendPTDUpdateAlert = async (recipientEmails, periodCode) => {
//     // 🔥 Current month name and year (e.g., August 2026)
//     const now = new Date();
//     const monthName = now.toLocaleString('en-US', { month: 'long' });
//     const deadlineDate = `15th ${monthName}`; // e.g., 15th August

//     const mailOptions = {
//         from: '"NI INDIA Financial Cost Tracker" <care.ni_india@nokia.com>',
//         // to: recipientEmails, // Array of all users + admin
//         to: recipientEmails,
//         cc: ["neha.sain.ext@nokia.com"],
//         subject: `NOTIFICATION: PTD for ${periodCode} Updated - NI INDIA Financial Cost Tracker`,
//         html: `
//         <div style="font-family: Calibri, Arial, sans-serif; font-size: 15px; color: #333; line-height: 1.6;">
//             <p>Dear Team,</p>
            
//             <p>
//                 PTD for <strong>${periodCode}</strong> has been updated in PBI. 
//                 Please check and provide forecast data to complete cost by <strong>${deadlineDate}</strong>. 
//                 Below is the link for FTC inputs Tool.
//             </p>

//             <p style="margin: 25px 0;">
//                 <strong>Tool Link:</strong> <a href="${TOOL_LINK}" style="color: #124191; font-weight: bold; text-decoration: underline;">${TOOL_LINK}</a>
//             </p>

//             <p>Best Regards,<br><strong>Neha Sain</strong></p>

//             <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #eee; font-size: 12px; color: #777;">
//                 <p><strong>Note:</strong> This is an automatically generated email. Please do not reply directly to this message. 
//                 For any query/comment/suggestion, please send an email to <a href="mailto:neha.sain.ext@nokia.com">neha.sain.ext@nokia.com</a>.</p>
//             </div>
//         </div>`
//     };

//     return transporter.sendMail(mailOptions);
// };

// server/services/mailService.js

const sendPTDUpdateAlert = async (recipientEmails, periodCode) => {
    // 🔥 logic: Hamesha current month ki 15th date dikhayega
    const now = new Date();
    const monthName = now.toLocaleString('en-US', { month: 'long' });
    const deadlineDate = `15th ${monthName}`; 

    const mailOptions = {
        from: '"NI INDIA Financial Cost Tracker" <care.ni_india@nokia.com>',
        // 🔥 TESTING MODE: Sending only to you
        to: "neha.sain.ext@nokia.com",
        cc: "neha.sain.ext@nokia.com",
        // 🔥 BCC added
        bcc: "care.ni_india@nokia.com", 
        
        subject: `NOTIFICATION: PTD for ${periodCode} Updated - NI INDIA Financial Cost Tracker`,
        html: `
        <div style="font-family: Calibri, Arial, sans-serif; font-size: 15px; color: #333; line-height: 1.6;">
            <p>Dear Team,</p>
            
            <p>
                PTD for <strong>${periodCode}</strong> has been updated in PBI. 
                Please check and provide forecast data to complete cost by <strong>${deadlineDate}</strong>. 
                Below is the link for FTC inputs Tool.
            </p>

            <p style="margin: 25px 0;">
                <strong>Tool Link:</strong> <a href="${TOOL_LINK}" style="color: #124191; font-weight: bold; text-decoration: underline;">${TOOL_LINK}</a>
            </p>

            <p>Best Regards,<br><strong>Neha Sain</strong></p>

            <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #eee; font-size: 12px; color: #777;">
                <p><strong>Note:</strong> This is an automatically generated email. Please do not reply directly to this message. 
                For any query/comment/suggestion, please send an email to <a href="mailto:neha.sain.ext@nokia.com">neha.sain.ext@nokia.com</a>.</p>
            </div>
        </div>`
    };

    return transporter.sendMail(mailOptions);
};


module.exports = {

    sendAccessRequestMail,

    sendApprovalMail,

    sendDeclineMail,

    sendOTPMail,

    sendCustomerUtilizationAlert,

    sendPTDUpdateAlert

};