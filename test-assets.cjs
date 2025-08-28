/**
 * Test script to verify environment-specific asset handling
 * This can be run to test that assets are accessible in all environments
 */

const environments = [
  {
    name: 'Development',
    assetBaseUrl: '/assets',
    testUrl: 'http://127.0.0.1:3000'
  },
  {
    name: 'Staging/Production',
    assetBaseUrl: '/excellence/assets',
    testUrl: 'http://127.0.0.1:3000'
  }
];

const testAssets = [
  'images/logos/PCAG_white_trans.svg',
  'images/icons/icon-32.png'
];

async function testAssetAccess() {
  console.log('🧪 Testing Environment-Specific Asset Access...\n');
  
  for (const env of environments) {
    console.log(`📂 Testing ${env.name} Environment:`);
    console.log(`   Asset Base URL: ${env.assetBaseUrl}`);
    
    for (const asset of testAssets) {
      const fullUrl = `${env.testUrl}${env.assetBaseUrl}/${asset}`;
      
      try {
        const response = await fetch(fullUrl, { method: 'HEAD' });
        const status = response.ok ? '✅ SUCCESS' : `❌ FAILED (${response.status})`;
        const contentType = response.headers.get('content-type') || 'unknown';
        
        console.log(`   ${status} ${asset} (${contentType})`);
      } catch (error) {
        console.log(`   ❌ ERROR ${asset} - ${error.message}`);
      }
    }
    console.log('');
  }
  
  console.log('🎯 Test completed. Both environments should show ✅ SUCCESS for all assets.\n');
  console.log('💡 If you see any failures, check:');
  console.log('   1. Server is running (node server.cjs)');
  console.log('   2. Build was successful (npm run build)');
  console.log('   3. Asset files exist in public/assets/images/');
}

// Run tests if this script is executed directly
if (typeof window === 'undefined' && typeof global !== 'undefined') {
  // Node.js environment
  const fetch = require('node-fetch');
  testAssetAccess().catch(console.error);
}

// Export for use in browser/module environments
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { testAssetAccess };
}