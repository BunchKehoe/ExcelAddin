const fs = require('fs')
const { exec } = require('child_process')
const path = require('path')

// Build custom functions as IIFE for Excel compatibility
function buildCustomFunctions() {
  console.log('Building custom functions for Excel Custom Functions runtime...')
  
  // Use TypeScript compiler to compile custom functions as IIFE
  const tsConfigForCustomFunctions = {
    compilerOptions: {
      target: 'ES2015',
      module: 'None',  // Don't use modules
      lib: ['ES2015', 'DOM'],
      strict: true,
      esModuleInterop: true,
      skipLibCheck: true,
      forceConsistentCasingInFileNames: true,
      declaration: false,
      outDir: './dist'
    },
    include: ['src/commands/customfunctions.ts']
  }
  
  // Write temporary tsconfig
  fs.writeFileSync('./tsconfig.customfunctions.json', JSON.stringify(tsConfigForCustomFunctions, null, 2))
  
  // Compile TypeScript to regular JavaScript
  exec('npx tsc -p tsconfig.customfunctions.json --outFile dist/customfunctions.js', (error, stdout, stderr) => {
    if (error) {
      console.error('Error compiling custom functions:', error)
      return
    }
    
    console.log('Custom functions compiled successfully!')
    
    // Clean up temporary tsconfig
    if (fs.existsSync('./tsconfig.customfunctions.json')) {
      fs.unlinkSync('./tsconfig.customfunctions.json')
    }
    
    // Wrap the output in IIFE if needed
    const outputPath = './dist/customfunctions.js'
    if (fs.existsSync(outputPath)) {
      let content = fs.readFileSync(outputPath, 'utf8')
      
      // Ensure it's wrapped in IIFE and doesn't have import/export
      if (!content.includes('(function()') && !content.includes('(() =>')) {
        content = `(function() {\n${content}\n})();`
        fs.writeFileSync(outputPath, content)
      }
      
      console.log('Custom functions written to dist/customfunctions.js')
    }
  })
}

buildCustomFunctions()