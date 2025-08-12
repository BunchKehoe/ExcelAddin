/**
 * Simple Node.js HTTP server for serving ExcelAddin frontend
 * Windows-compatible replacement for PM2 + serve
 * Uses only Node.js built-in modules to avoid Unix shebang issues
 */

const http = require('http');
const path = require('path');
const fs = require('fs');
const url = require('url');

// Configuration
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '127.0.0.1';
const PROJECT_ROOT = path.resolve(__dirname, '../..');
const STATIC_DIR = path.join(PROJECT_ROOT, 'dist');

// MIME types for common web files
const MIME_TYPES = {
  '.html': 'text/html',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.eot': 'application/vnd.ms-fontobject'
};

/**
 * Get MIME type for a file extension
 */
function getMimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return MIME_TYPES[ext] || 'application/octet-stream';
}

/**
 * Serve static files
 */
function serveStaticFile(filePath, res) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      console.error('Error reading file:', filePath, err.message);
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('404 Not Found');
      return;
    }

    const mimeType = getMimeType(filePath);
    res.writeHead(200, { 
      'Content-Type': mimeType,
      'Cache-Control': 'public, max-age=3600'
    });
    res.end(data);
  });
}

/**
 * Check if a file exists
 */
function fileExists(filePath) {
  try {
    return fs.statSync(filePath).isFile();
  } catch (err) {
    return false;
  }
}

/**
 * Main request handler
 */
function handleRequest(req, res) {
  const parsedUrl = url.parse(req.url);
  const pathname = parsedUrl.pathname;
  
  console.log(`${new Date().toISOString()} ${req.method} ${pathname}`);

  // Security: prevent directory traversal
  if (pathname.includes('..')) {
    res.writeHead(400, { 'Content-Type': 'text/plain' });
    res.end('400 Bad Request');
    return;
  }

  // Determine file path
  let filePath;
  if (pathname === '/') {
    filePath = path.join(STATIC_DIR, 'index.html');
  } else {
    filePath = path.join(STATIC_DIR, pathname);
  }

  // Check if file exists
  if (fileExists(filePath)) {
    serveStaticFile(filePath, res);
  } else {
    // For SPA (Single Page Application) routing - serve index.html for non-file requests
    // This handles React Router routes
    const ext = path.extname(pathname);
    if (!ext && pathname !== '/') {
      // This looks like a SPA route, serve index.html
      const indexPath = path.join(STATIC_DIR, 'index.html');
      if (fileExists(indexPath)) {
        serveStaticFile(indexPath, res);
      } else {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('404 Not Found - index.html missing');
      }
    } else {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('404 Not Found');
    }
  }
}

// Verify static directory exists
if (!fs.existsSync(STATIC_DIR)) {
  console.error(`Static directory does not exist: ${STATIC_DIR}`);
  console.error('Please ensure the frontend has been built first (npm run build:staging)');
  process.exit(1);
}

// Create HTTP server
const server = http.createServer(handleRequest);

// Start server
server.listen(PORT, HOST, () => {
  console.log(`ExcelAddin Frontend Server started successfully`);
  console.log(`Server: http://${HOST}:${PORT}`);
  console.log(`Static files: ${STATIC_DIR}`);
  console.log(`Process ID: ${process.pid}`);
  console.log(`Node version: ${process.version}`);
  console.log('Press Ctrl+C to stop');
});

// Graceful shutdown handlers
function shutdown(signal) {
  console.log(`\nReceived ${signal}, shutting down gracefully...`);
  server.close(() => {
    console.log('HTTP server closed');
    process.exit(0);
  });

  // Force close after 5 seconds
  setTimeout(() => {
    console.error('Force closing server...');
    process.exit(1);
  }, 5000);
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));

// Handle uncaught exceptions
process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
  process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  process.exit(1);
});