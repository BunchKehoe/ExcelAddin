/**
 * Quick Certificate Fix Script
 * 
 * This script automatically diagnoses and fixes common certificate issues
 * for Excel Add-in local development.
 * 
 * Usage: node fix-certificates.js
 */

const { spawn, exec } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

console.log('🔧 Excel Add-in Certificate Diagnostic Tool\n');

async function runCommand(command, description) {
    return new Promise((resolve, reject) => {
        console.log(`⏳ ${description}...`);
        const process = spawn(command.split(' ')[0], command.split(' ').slice(1), {
            stdio: 'pipe',
            shell: true
        });
        
        let output = '';
        let error = '';
        
        process.stdout.on('data', (data) => {
            output += data.toString();
        });
        
        process.stderr.on('data', (data) => {
            error += data.toString();
        });
        
        process.on('close', (code) => {
            if (code === 0) {
                console.log(`✅ ${description} completed\n`);
                resolve(output);
            } else {
                console.log(`❌ ${description} failed: ${error}\n`);
                resolve(null);
            }
        });
    });
}

function checkCertificateFiles() {
    const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
    const certFile = path.join(certDir, 'localhost.crt');
    const keyFile = path.join(certDir, 'localhost.key');
    
    console.log('📁 Checking certificate files...');
    console.log(`   Directory: ${certDir}`);
    
    const dirExists = fs.existsSync(certDir);
    const certExists = fs.existsSync(certFile);
    const keyExists = fs.existsSync(keyFile);
    
    console.log(`   Directory exists: ${dirExists ? '✅' : '❌'}`);
    console.log(`   Certificate exists: ${certExists ? '✅' : '❌'}`);
    console.log(`   Private key exists: ${keyExists ? '✅' : '❌'}\n`);
    
    return dirExists && certExists && keyExists;
}

async function main() {
    console.log('Starting certificate diagnosis...\n');
    
    // Step 1: Check if certificate files exist
    const filesExist = checkCertificateFiles();
    
    // Step 2: Verify certificate status
    console.log('🔍 Verifying certificate installation status...');
    const verifyResult = await runCommand('npm run cert:verify', 'Certificate verification');
    
    const needsInstall = !verifyResult || verifyResult.includes('You need to install') || !filesExist;
    
    if (needsInstall) {
        console.log('🚨 Certificates need to be installed or refreshed\n');
        
        // Step 3: Uninstall old certificates if they exist
        if (filesExist) {
            console.log('🧹 Removing old certificates...');
            await runCommand('npm run cert:uninstall', 'Certificate uninstallation');
        }
        
        // Step 4: Install fresh certificates
        console.log('📜 Installing fresh certificates...');
        const installResult = await runCommand('npm run cert:install', 'Certificate installation');
        
        if (installResult) {
            console.log('✅ Certificate installation completed!');
            
            // Step 5: Verify installation
            console.log('🔍 Verifying new certificate installation...');
            await runCommand('npm run cert:verify', 'Final verification');
            
            console.log('\n🎉 Certificate fix completed!');
            console.log('\n📋 Next steps:');
            console.log('   1. Restart Excel completely');
            console.log('   2. Run: npm run dev');
            console.log('   3. Load your add-in in Excel');
            console.log('\nIf you still see certificate errors, please check the Certificate Guide:\n   📖 CERTIFICATE_GUIDE.md\n');
        } else {
            console.log('\n❌ Certificate installation failed!');
            console.log('\n🔧 Manual steps to try:');
            console.log('   1. Run PowerShell as Administrator');
            console.log('   2. cd to your project directory');
            console.log('   3. npm run cert:install');
            console.log('\nFor more help, see: CERTIFICATE_GUIDE.md\n');
        }
    } else {
        console.log('✅ Certificates are already properly installed!');
        console.log('\n🤔 If you\'re still seeing certificate errors in Excel:');
        console.log('   1. Restart Excel completely');
        console.log('   2. Clear browser cache (Ctrl+Shift+Delete)');
        console.log('   3. Check Windows Certificate Store: certlm.msc');
        console.log('\nFor advanced troubleshooting, see: CERTIFICATE_GUIDE.md\n');
    }
}

// Handle errors gracefully
process.on('unhandledRejection', (error) => {
    console.error('\n❌ Unexpected error:', error.message);
    console.log('\nPlease try running these commands manually:');
    console.log('   npm run cert:uninstall');
    console.log('   npm run cert:install');
    console.log('\nFor more help, see: CERTIFICATE_GUIDE.md\n');
    process.exit(1);
});

main().catch(console.error);