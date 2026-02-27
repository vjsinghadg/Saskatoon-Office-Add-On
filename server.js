const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Serve manifest.xml
app.get('/manifest.xml', (req, res) => {
    res.setHeader('Content-Type', 'application/xml');
    res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Serve function-file.html
app.get('/function-file/function-file.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'function-file.html'));
});

// Serve function-file.js
app.get('/scripts/function-file.js', (req, res) => {
    res.setHeader('Content-Type', 'application/javascript');
    res.sendFile(path.join(__dirname, 'function-file.js'));
});

// Serve configuration
app.get('/api/config', (req, res) => {
    const config = {
        infosecEmail: process.env.INFOSEC_EMAIL || "infosec@company.com",
        spamReportEmail: process.env.SPAM_REPORT_EMAIL || "spam-report@company.com",
        supportEmail: process.env.SUPPORT_EMAIL || "support@company.com",
        version: "1.0.0"
    };
    res.json(config);
});

// Health check
app.get('/health', (req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({ error: 'Internal Server Error', message: err.message });
});

// Start server
const server = https.createServer(
    {
        key: fs.readFileSync(path.join(__dirname, 'certs', 'key.pem')),
        cert: fs.readFileSync(path.join(__dirname, 'certs', 'cert.pem'))
    },
    app
);

server.listen(PORT, () => {
    console.log(`ADGSentinel Add-in Server running at https://localhost:${PORT}`);
    console.log('Manifest URL: https://localhost:3000/manifest.xml');
    console.log('\nTo use in Outlook:');
    console.log('1. Open Outlook Web');
    console.log('2. Go to Settings > Get Add-ins');
    console.log('3. Choose "My Add-ins" > "Upload My Add-in"');
    console.log('4. Upload the manifest.xml or provide URL to manifest');
});
