import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'path'
import fs from 'fs'
import { fileURLToPath, URL } from 'node:url'

// https://vitejs.dev/config/
export default defineConfig(({ command, mode }) => {
  // Load environment variables based on mode
  const env = loadEnv(mode, process.cwd(), '')
  
  // Determine if this is production/staging build
  const isProduction = mode === 'production'
  const isStaging = mode === 'staging'
  const isDevelopment = mode === 'development'
  
  // Base path for different environments
  let basePath = '/'
  if (isStaging || isProduction) {
    basePath = '/excellence/'
  }
  
  // HTTPS configuration for development
  let httpsConfig = undefined
  if (isDevelopment) {
    // Check for Office Add-in certificates
    const certPath = process.env.HOME + '/.office-addin-dev-certs/localhost.crt'
    const keyPath = process.env.HOME + '/.office-addin-dev-certs/localhost.key'
    
    if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
      httpsConfig = {
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certPath)
      }
    } else {
      console.warn('⚠️  Office Add-in certificates not found. Run "npm run cert:install" for proper HTTPS support.')
      httpsConfig = true // Use Vite's default self-signed certificates
    }
  }

  return {
    plugins: [react()],
    base: basePath,
    
    // Configure multiple entry points for Excel Add-in
    build: {
      outDir: 'dist',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, 'taskpane.html'),
          commands: resolve(__dirname, 'commands.html')
        },
        output: {
          // Organize output files  
          chunkFileNames: 'assets/js/[name]-[hash].js',
          entryFileNames: 'assets/js/[name]-[hash].js',
          assetFileNames: (assetInfo) => {
            const info = assetInfo.name.split('.')
            const ext = info[info.length - 1]
            if (/png|jpe?g|svg|gif|tiff|bmp|ico/i.test(ext)) {
              return `assets/images/[name]-[hash][extname]`
            }
            if (/css/i.test(ext)) {
              return `assets/css/[name]-[hash][extname]`
            }
            return `assets/[name]-[hash][extname]`
          }
        }
      },
      // Copy static assets and manifest files
      copyPublicDir: true
    },
    
    // Public directory for static assets
    publicDir: 'public',
    
    // Development server configuration
    server: {
      host: '0.0.0.0',
      port: 3000,
      https: httpsConfig,
      cors: true,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
        'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
      }
    },
    
    // Preview server configuration (for production preview)
    preview: {
      host: '0.0.0.0',
      port: 3000,
      https: isProduction || isStaging,
      cors: true
    },
    
    // Resolve configuration
    resolve: {
      alias: {
        '@': fileURLToPath(new URL('./src', import.meta.url))
      }
    },
    
    // Environment variable prefix
    envPrefix: 'VITE_',
    
    // Define global constants
    define: {
      __MODE__: JSON.stringify(mode),
      __BASE_PATH__: JSON.stringify(basePath)
    }
  }
})