const fs = require('fs');
const path = require('path');

export default function handler(req, res) {
    if (req.method === 'GET') {
        try {
            res.setHeader('Content-Type', 'application/xml');
            res.setHeader('Cache-Control', 'public, max-age=3600');
            const manifestPath = path.join(process.cwd(), 'manifest.xml');
            const manifest = fs.readFileSync(manifestPath, 'utf8');
            res.status(200).send(manifest);
        } catch (error) {
            console.error('Error reading manifest:', error);
            res.status(500).json({ error: 'Failed to load manifest', message: error.message });
        }
    } else {
        res.status(405).json({ error: 'Method Not Allowed' });
    }
}
