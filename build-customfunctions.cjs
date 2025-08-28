const fs = require('fs')
const { exec } = require('child_process')
const path = require('path')

// Build custom functions for Excel compatibility following Microsoft guidance
function buildCustomFunctions() {
  console.log('Building custom functions for Excel Custom Functions runtime...')
  
  // Create dist directory if it doesn't exist
  if (!fs.existsSync('./dist')) {
    fs.mkdirSync('./dist')
  }

  // Create public directory if it doesn't exist
  if (!fs.existsSync('./public')) {
    fs.mkdirSync('./public')
  }
  
  // Use TypeScript compiler to compile custom functions as regular JS (no modules)
  const tsConfigForCustomFunctions = {
    compilerOptions: {
      target: 'ES2015',
      module: 'None',  // Don't use modules - Excel Custom Functions requirement
      lib: ['ES2015', 'DOM'],
      strict: true,
      esModuleInterop: true,
      skipLibCheck: true,
      forceConsistentCasingInFileNames: true,
      declaration: false,
      outDir: './dist'
    },
    include: ['src/functions/functions.ts']
  }
  
  // Write temporary tsconfig
  fs.writeFileSync('./tsconfig.customfunctions.json', JSON.stringify(tsConfigForCustomFunctions, null, 2))
  
  // Compile TypeScript to regular JavaScript
  exec('npx tsc -p tsconfig.customfunctions.json --outFile dist/customfunctions.js', (error, stdout, stderr) => {
    if (error) {
      console.error('Error compiling custom functions:', error)
      console.log('Using pre-compiled JavaScript version...')
      
      // If TypeScript compilation fails, copy existing JS version
      const fallbackPath = './public/customfunctions.js'
      if (fs.existsSync(fallbackPath)) {
        console.log('Using existing customfunctions.js')
      }
    } else {
      console.log('Custom functions compiled successfully from TypeScript!')
      
      // Copy the compiled file to public directory for web serving
      const distPath = './dist/customfunctions.js'
      const publicPath = './public/customfunctions.js'
      
      if (fs.existsSync(distPath)) {
        fs.copyFileSync(distPath, publicPath)
        console.log('Custom functions copied to public/customfunctions.js')
      }
    }
    
    // Clean up temporary tsconfig
    if (fs.existsSync('./tsconfig.customfunctions.json')) {
      fs.unlinkSync('./tsconfig.customfunctions.json')
    }
    
    // Copy functions.json metadata to public directory
    const functionsJsonSrc = './src/functions/functions.json'
    const functionsJsonDest = './public/functions.json'
    
    if (fs.existsSync(functionsJsonSrc)) {
      fs.copyFileSync(functionsJsonSrc, functionsJsonDest)
      console.log('Functions metadata copied to public/functions.json')
    }
  })
}

buildCustomFunctions()