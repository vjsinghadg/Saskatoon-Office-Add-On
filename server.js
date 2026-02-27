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
    try {
        res.setHeader('Content-Type', 'application/xml');

        // Try multiple possible paths
        const possiblePaths = [
            path.join(__dirname, 'manifest.xml'),
            path.join(process.cwd(), 'manifest.xml'),
            './manifest.xml'
        ];

        let manifest = null;
        let foundPath = null;

        for (const manifestPath of possiblePaths) {
            try {
                if (fs.existsSync(manifestPath)) {
                    manifest = fs.readFileSync(manifestPath, 'utf8');
                    foundPath = manifestPath;
                    break;
                }
            } catch (e) {
                continue;
            }
        }

        if (!manifest) {
            console.error('Manifest.xml not found in any of:', possiblePaths);
            console.error('Current working directory:', process.cwd());
            console.error('__dirname:', __dirname);
            return res.status(404).json({
                error: 'Manifest not found',
                cwd: process.cwd(),
                dirname: __dirname
            });
        }

        console.log(`Serving manifest from: ${foundPath}`);
        res.send(manifest);
    } catch (error) {
        console.error('Error reading manifest:', error);
        res.status(500).json({ error: 'Failed to load manifest', message: error.message });
    }
});

// Serve function-file.html
app.get('/function-file/function-file.html', (req, res) => {
    try {
        const possiblePaths = [
            path.join(__dirname, 'function-file.html'),
            path.join(process.cwd(), 'function-file.html'),
            './function-file.html'
        ];

        for (const htmlPath of possiblePaths) {
            if (fs.existsSync(htmlPath)) {
                console.log(`Serving function-file.html from: ${htmlPath}`);
                return res.sendFile(path.resolve(htmlPath));
            }
        }

        console.error('function-file.html not found');
        res.status(404).json({ error: 'Function file not found' });
    } catch (error) {
        console.error('Error reading function-file.html:', error);
        res.status(500).json({ error: 'Failed to load function file', message: error.message });
    }
});

// Serve function-file.js
app.get('/scripts/function-file.js', (req, res) => {
    try {
        res.setHeader('Content-Type', 'application/javascript');
        const possiblePaths = [
            path.join(__dirname, 'function-file.js'),
            path.join(process.cwd(), 'function-file.js'),
            './function-file.js'
        ];

        for (const jsPath of possiblePaths) {
            if (fs.existsSync(jsPath)) {
                console.log(`Serving function-file.js from: ${jsPath}`);
                return res.sendFile(path.resolve(jsPath));
            }
        }

        console.error('function-file.js not found');
        res.status(404).json({ error: 'Script not found' });
    } catch (error) {
        console.error('Error reading function-file.js:', error);
        res.status(500).json({ error: 'Failed to load script', message: error.message });
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

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({
        status: 'ok',
        timestamp: new Date().toISOString(),
        cwd: process.cwd(),
        dirname: __dirname
    });
});

// Root endpoint
app.get('/', (req, res) => {
    res.json({
        message: 'ADGSentinel Outlook Add-in Server',
        version: '1.0.0',
        endpoints: {
            manifest: '/manifest.xml',
            healthCheck: '/health',
            functionFile: '/function-file/function-file.html',
            script: '/scripts/function-file.js',
            config: '/api/config'
        }
    });
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Error:', err.stack);
    res.status(500).json({ error: 'Internal Server Error', message: err.message });
});

// For local development and server deployments (Railway, Heroku, etc)
if (require.main === module) {
    const PORT = process.env.PORT || 3000;

    console.log('Starting ADGSentinel Outlook Add-in Server...');
    console.log('PORT:', PORT);
    console.log('NODE_ENV:', process.env.NODE_ENV);
    console.log('CWD:', process.cwd());
    console.log('__dirname:', __dirname);

    // Check if certificate files exist (for local HTTPS development)
    const certPath = path.join(__dirname, 'certs', 'key.pem');
    const hasCerts = fs.existsSync(certPath);

    if (hasCerts && process.env.NODE_ENV !== 'production') {
        // Use HTTPS for local development
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
                console.log(`✓ ADGSentinel Add-in Server running at https://localhost:${PORT}`);
                console.log('Manifest URL: https://localhost:3000/manifest.xml');
            });
        } catch (error) {
            console.error('Error starting HTTPS server:', error.message);
            console.log('Falling back to HTTP...');
            app.listen(PORT, () => {
                console.log(`✓ ADGSentinel Add-in Server running at http://localhost:${PORT}`);
            });
        }
    } else {
        // Use HTTP for production (Railway, etc.) - HTTPS handled by platform
        app.listen(PORT, '0.0.0.0', () => {
            console.log(`✓ ADGSentinel Add-in Server running on port ${PORT}`);
            console.log(`✓ Manifest URL: https://your-railway-domain.railway.app/manifest.xml`);
        });
    }
}

// Export for Vercel serverless function
module.exports = app;
