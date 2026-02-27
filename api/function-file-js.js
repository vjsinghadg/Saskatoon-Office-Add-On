const fs = require('fs');
const path = require('path');

module.exports = (req, res) => {
    if (req.method === 'GET') {
        try {
            res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
            res.setHeader('Cache-Control', 'public, max-age=3600');
            const jsPath = path.join(process.cwd(), 'function-file.js');
            const js = fs.readFileSync(jsPath, 'utf8');
            res.status(200).end(js);
        } catch (error) {
            console.error('Error reading function-file.js:', error);
            res.status(500).end(JSON.stringify({ error: 'Failed to load script', message: error.message }));
        }
    } else {
        res.status(405).end(JSON.stringify({ error: 'Method Not Allowed' }));
    }
};
