const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '127.0.0.1';

// Serve static files from dist directory
app.use('/excellence', express.static(path.join(__dirname, 'dist')));

// Serve manifest files and assets at root level too (for backward compatibility)
app.use('/assets', express.static(path.join(__dirname, 'dist/assets')));
app.get('/manifest*.xml', (req, res) => {
    const manifestPath = path.join(__dirname, 'dist', req.path);
    if (fs.existsSync(manifestPath)) {
        res.setHeader('Content-Type', 'application/xml');
        res.sendFile(manifestPath);
    } else {
        res.status(404).send('Manifest not found');
    }
});

// Serve functions.json
app.get('/functions.json', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist/functions.json'));
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
        res.sendFile(path.join(__dirname, 'dist/taskpane.html'));
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