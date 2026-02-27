const fs = require('fs');
const path = require('path');

module.exports = (req, res) => {
    if (req.method === 'GET') {
        try {
            res.setHeader('Content-Type', 'text/html; charset=utf-8');
            res.setHeader('Cache-Control', 'public, max-age=3600');
            const htmlPath = path.join(process.cwd(), 'function-file.html');
            const html = fs.readFileSync(htmlPath, 'utf8');
            res.status(200).end(html);
        } catch (error) {
            console.error('Error reading function-file.html:', error);
            res.status(500).end(JSON.stringify({ error: 'Failed to load function file', message: error.message }));
        }
    } else {
        res.status(405).end(JSON.stringify({ error: 'Method Not Allowed' }));
    }
};
