const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();

// Middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Security headers
app.use((req, res, next) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'SAMEORIGIN');
    res.setHeader('Content-Security-Policy', "default-src 'self' https://appsforoffice.microsoft.com https://office.com; script-src 'self' https://appsforoffice.microsoft.com");
    next();
});

// Serve manifest.xml
app.get('/manifest.xml', (req, res) => {

    console.log(1111);
    try {
        res.setHeader('Content-Type', 'application/xml');
        const manifestPath = path.join(__dirname, 'manifest.xml');
        const manifest = fs.readFileSync(manifestPath, 'utf8');
        res.send(manifest);
    } catch (error) {
        console.error('Error reading manifest:', error);
        res.status(500).json({ error: 'Failed to load manifest' });
    }
});

// Serve function-file.html
app.get('/function-file/function-file.html', (req, res) => {
    try {
        const htmlPath = path.join(__dirname, 'function-file.html');
        res.sendFile(htmlPath);
    } catch (error) {
        console.error('Error reading function-file.html:', error);
        res.status(500).json({ error: 'Failed to load function file' });
    }
});

// Serve function-file.js
app.get('/scripts/function-file.js', (req, res) => {
    try {
        res.setHeader('Content-Type', 'application/javascript');
        const jsPath = path.join(__dirname, 'function-file.js');
        res.sendFile(jsPath);
    } catch (error) {
        console.error('Error reading function-file.js:', error);
        res.status(500).json({ error: 'Failed to load script' });
    }
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
    console.error('Error:', err.stack);
    res.status(500).json({ error: 'Internal Server Error', message: err.message });
});

// For local development (runs with npm start)
if (require.main === module) {
    const PORT = process.env.PORT || 3000;
    try {
        const https = require('https');
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
    } catch (error) {
        console.error('Error starting HTTPS server:', error.message);
        console.log('Falling back to HTTP...');
        app.listen(PORT, () => {
            console.log(`ADGSentinel Add-in Server running at http://localhost:${PORT}`);
        });
    }
}

// Export for Vercel serverless function
module.exports = app;
