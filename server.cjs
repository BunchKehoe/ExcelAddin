const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '127.0.0.1';

// CORS headers for Excel Add-in compatibility (must be before routes)
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, PATCH, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'X-Requested-With, content-type, Authorization');
    
    // Handle preflight requests
    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }
    
    next();
});

// Serve static files from dist directory
app.use('/excellence', express.static(path.join(__dirname, 'dist')));

// Serve assets at root level for development compatibility  
app.use('/assets', express.static(path.join(__dirname, 'dist/assets')));

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'healthy', 
        service: 'exceladdin-frontend',
        timestamp: new Date().toISOString(),
        environment: process.env.NODE_ENV || 'production',
        port: PORT,
        host: HOST,
        distPath: path.join(__dirname, 'dist'),
        distExists: fs.existsSync(path.join(__dirname, 'dist'))
    });
});

// Debug endpoint to check asset availability
app.get('/debug/assets', (req, res) => {
    const distPath = path.join(__dirname, 'dist');
    const assetsPath = path.join(distPath, 'assets');
    
    const distExists = fs.existsSync(distPath);
    const assetsExist = fs.existsSync(assetsPath);
    
    let assetsList = [];
    if (assetsExist) {
        try {
            assetsList = fs.readdirSync(assetsPath);
        } catch (err) {
            assetsList = [`Error reading assets: ${err.message}`];
        }
    }
    
    res.json({
        distPath,
        assetsPath,
        distExists,
        assetsExist,
        assetsList,
        cwd: process.cwd()
    });
});
// Serve manifest files at root level
app.get('/manifest*.xml', (req, res) => {
    const manifestPath = path.join(__dirname, 'dist', req.path);
    if (fs.existsSync(manifestPath)) {
        res.setHeader('Content-Type', 'application/xml');
        res.sendFile(manifestPath);
    } else {
        // Try public directory as fallback
        const publicManifestPath = path.join(__dirname, 'public', req.path);
        if (fs.existsSync(publicManifestPath)) {
            res.setHeader('Content-Type', 'application/xml');
            res.sendFile(publicManifestPath);
        } else {
            res.status(404).send('Manifest not found');
        }
    }
});

// Serve functions.json
app.get('/functions.json', (req, res) => {
    const functionsPath = path.join(__dirname, 'dist/functions.json');
    if (fs.existsSync(functionsPath)) {
        res.sendFile(functionsPath);
    } else {
        // Try public directory as fallback
        const publicFunctionsPath = path.join(__dirname, 'public/functions.json');
        if (fs.existsSync(publicFunctionsPath)) {
            res.sendFile(publicFunctionsPath);
        } else {
            res.status(404).send('functions.json not found');
        }
    }
});

// Excel Add-in specific endpoints (required by Excel)
app.get('/excellence/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist/taskpane.html'));
});

app.get('/excellence/commands.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist/commands.html'));
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'healthy', 
        service: 'exceladdin-frontend',
        timestamp: new Date().toISOString(),
        environment: process.env.NODE_ENV || 'production',
        port: PORT,
        host: HOST
    });
});

// Fallback for any other Excel Add-in requests
app.get('/excellence/*', (req, res) => {
    const filePath = path.join(__dirname, 'dist', req.path.replace('/excellence/', ''));
    if (fs.existsSync(filePath)) {
        res.sendFile(filePath);
    } else {
        // For SPA routing, serve taskpane.html as fallback
        const taskpanePath = path.join(__dirname, 'dist/taskpane.html');
        if (fs.existsSync(taskpanePath)) {
            res.sendFile(taskpanePath);
        } else {
            // Try root taskpane.html as last resort
            const rootTaskpanePath = path.join(__dirname, 'taskpane.html');
            if (fs.existsSync(rootTaskpanePath)) {
                res.sendFile(rootTaskpanePath);
            } else {
                res.status(404).send('File not found');
            }
        }
    }
});

// CORS headers for Excel Add-in compatibility
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, PATCH, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'X-Requested-With, content-type, Authorization');
    next();
});

// Start server
const server = app.listen(PORT, HOST, () => {
    console.log(`Excel Add-in Frontend Server running on http://${HOST}:${PORT}`);
    console.log(`Environment: ${process.env.NODE_ENV || 'production'}`);
    console.log(`Serving from: ${path.join(__dirname, 'dist')}`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
    console.log('SIGTERM received, shutting down gracefully');
    server.close(() => {
        console.log('Frontend server stopped');
        process.exit(0);
    });
});

process.on('SIGINT', () => {
    console.log('SIGINT received, shutting down gracefully');
    server.close(() => {
        console.log('Frontend server stopped');
        process.exit(0);
    });
});