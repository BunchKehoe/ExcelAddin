/**
 * Windows-compatible frontend server wrapper
 * This file avoids Unix shebang issues on Windows systems
 */

const { spawn } = require('child_process');
const path = require('path');

// Get serve from node_modules
const servePath = path.join(__dirname, '../../node_modules/serve/build/main.js');

// Start serve with the same arguments that PM2 would use
const args = ['-s', 'dist', '-l', '3000'];

console.log('Starting frontend server...');
console.log('Serve path:', servePath);
console.log('Arguments:', args);

// Spawn node with serve script
const child = spawn('node', [servePath, ...args], {
    stdio: 'inherit',
    cwd: path.join(__dirname, '../../')
});

child.on('error', (error) => {
    console.error('Failed to start frontend server:', error);
    process.exit(1);
});

child.on('exit', (code) => {
    console.log(`Frontend server exited with code ${code}`);
    process.exit(code);
});

// Handle process termination
process.on('SIGTERM', () => {
    console.log('Received SIGTERM, shutting down gracefully...');
    child.kill('SIGTERM');
});

process.on('SIGINT', () => {
    console.log('Received SIGINT, shutting down gracefully...');
    child.kill('SIGINT');
});