module.exports = (req, res) => {
    if (req.method === 'GET') {
        res.setHeader('Content-Type', 'application/json');
        res.status(200).end(JSON.stringify({
            status: 'ok',
            timestamp: new Date().toISOString(),
            version: '1.0.0',
            environment: 'serverless'
        }));
    } else {
        res.status(405).end(JSON.stringify({ error: 'Method Not Allowed' }));
    }
};
