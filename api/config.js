export default function handler(req, res) {
    if (req.method === 'GET') {
        const config = {
            infosecEmail: process.env.INFOSEC_EMAIL || "infosec@company.com",
            spamReportEmail: process.env.SPAM_REPORT_EMAIL || "spam-report@company.com",
            supportEmail: process.env.SUPPORT_EMAIL || "support@company.com",
            gophishUrl: process.env.GOPHISH_URL || "https://saskaatoon.ca",
            version: "1.0.0"
        };
        res.status(200).json(config);
    } else {
        res.status(405).json({ error: 'Method Not Allowed' });
    }
}
