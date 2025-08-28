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
  
  // Read the TypeScript source
  const functionsTs = fs.readFileSync('./src/functions/functions.ts', 'utf8')
  
  // Create standalone version without exports and type annotations for Excel Custom Functions runtime
  let standaloneVersion = functionsTs
    .replace(/export function/g, 'function')  // Remove export keywords
    .replace(/export \{[^}]*\};?/g, '')       // Remove export statements
    .replace(/import[^;]+;/g, '')             // Remove any import statements
    .replace(/: number/g, '')                 // Remove number type annotations
    .replace(/: string/g, '')                 // Remove string type annotations  
    .replace(/: any\[\]\[\]/g, '')            // Remove any[][] type annotations
    .replace(/: string\[\]/g, '')             // Remove string[] type annotations
    .replace(/\/\/ Type declaration for CustomFunctions[\s\S]*?\} \| undefined;/g, '') // Remove type declarations block
  
  // Write standalone JavaScript version to public directory
  const standaloneJs = `"use strict";\n${standaloneVersion}`
  
  fs.writeFileSync('./public/customfunctions.js', standaloneJs)
  console.log('Custom functions standalone version created at public/customfunctions.js')
  
  // Also create functions.js (referenced by functions.html)
  fs.writeFileSync('./public/functions.js', standaloneJs)
  console.log('Custom functions standalone version created at public/functions.js')
  
  // Copy functions.json metadata to public directory
  const functionsJsonSrc = './src/functions/functions.json'
  const functionsJsonDest = './public/functions.json'
  
  if (fs.existsSync(functionsJsonSrc)) {
    fs.copyFileSync(functionsJsonSrc, functionsJsonDest)
    console.log('Functions metadata copied to public/functions.json')
  }
  
  // Copy functions.html to public directory
  const functionsHtmlSrc = './src/functions/functions.html'
  const functionsHtmlDest = './public/functions.html'
  
  if (fs.existsSync(functionsHtmlSrc)) {
    fs.copyFileSync(functionsHtmlSrc, functionsHtmlDest)
    console.log('Functions HTML copied to public/functions.html')
  }
}

buildCustomFunctions()