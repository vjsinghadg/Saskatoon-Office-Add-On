module.exports = (req, res) => {
    if (req.method === 'GET') {
        res.setHeader('Content-Type', 'application/json');
        const config = {
            infosecEmail: process.env.INFOSEC_EMAIL || "infosec@company.com",
            spamReportEmail: process.env.SPAM_REPORT_EMAIL || "spam-report@company.com",
            supportEmail: process.env.SUPPORT_EMAIL || "support@company.com",
            gophishUrl: process.env.GOPHISH_URL || "https://saskaatoon.ca",
            version: "1.0.0"
        };
        res.status(200).end(JSON.stringify(config));
    } else {
        res.status(405).end(JSON.stringify({ error: 'Method Not Allowed' }));
    }
};
