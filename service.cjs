const Service = require('node-windows').Service;
const path = require('path');

// Create a new service object
const svc = new Service({
  name: 'ExcelAddin Frontend',
  description: 'Excel Add-in Frontend Web Server (Vite + Express)',
  script: path.join(__dirname, 'server.cjs'),
  nodeOptions: [
    '--max_old_space_size=4096'
  ],
  env: [{
    name: "NODE_ENV",
    value: process.env.NODE_ENV || "production"
  },{
    name: "PORT", 
    value: process.env.PORT || "3000"
  },{
    name: "HOST",
    value: process.env.HOST || "127.0.0.1"
  }]
});

// Listen for the "install" event, which indicates the process is available as a service
svc.on('install', () => {
  console.log('ExcelAddin Frontend service installed successfully');
  console.log('Starting service...');
  svc.start();
});

// Listen for the "start" event and let us know the service started
svc.on('start', () => {
  console.log('ExcelAddin Frontend service started successfully');
  console.log('Service is now running and will start automatically on boot');
});

// Listen for the "stop" event and let us know the service stopped
svc.on('stop', () => {
  console.log('ExcelAddin Frontend service stopped');
});

// Listen for the "uninstall" event so we know when it's removed
svc.on('uninstall', () => {
  console.log('ExcelAddin Frontend service uninstalled successfully');
});

// Handle command line arguments
const action = process.argv[2];

switch (action) {
  case 'install':
    console.log('Installing ExcelAddin Frontend service...');
    svc.install();
    break;
    
  case 'uninstall':
    console.log('Uninstalling ExcelAddin Frontend service...');
    svc.uninstall();
    break;
    
  case 'start':
    console.log('Starting ExcelAddin Frontend service...');
    svc.start();
    break;
    
  case 'stop':
    console.log('Stopping ExcelAddin Frontend service...');
    svc.stop();
    break;
    
  case 'restart':
    console.log('Restarting ExcelAddin Frontend service...');
    svc.restart();
    break;
    
  default:
    console.log('ExcelAddin Frontend Service Manager');
    console.log('');
    console.log('Usage: node service.js <action>');
    console.log('');
    console.log('Actions:');
    console.log('  install   - Install the service');
    console.log('  uninstall - Remove the service');
    console.log('  start     - Start the service');
    console.log('  stop      - Stop the service');
    console.log('  restart   - Restart the service');
    console.log('');
    console.log('Once installed, the service will:');
    console.log('- Start automatically on system boot');
    console.log('- Run as a Windows service');
    console.log('- Be manageable via Windows Services (services.msc)');
    console.log('- Log to Windows Event Log');
    break;
}

module.exports = svc;