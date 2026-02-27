export default function handler(req, res) {
    if (req.method === 'GET') {
        res.status(200).json({
            status: 'ok',
            timestamp: new Date().toISOString(),
            version: '1.0.0',
            environment: 'serverless'
        });
    } else {
        res.status(405).json({ error: 'Method Not Allowed' });
    }
}
