import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteStaticCopy } from 'vite-plugin-static-copy';
import { resolve } from 'path';
import { readFileSync, existsSync } from 'fs';
import { homedir } from 'os';
import { fileURLToPath, URL } from 'node:url';

export default defineConfig(({ mode }) => {
  const isProduction = mode === 'production';
  const isStaging = mode === 'staging';
  const isDevelopment = mode === 'development';

  // Determine public path based on environment
  let publicPath = '/';
  if (isStaging || isProduction) {
    publicPath = '/excellence/';
  }

  // SSL certificate paths for dev server
  const certPath = homedir() + '/.office-addin-dev-certs/localhost.crt';
  const keyPath = homedir() + '/.office-addin-dev-certs/localhost.key';

  // Determine which manifest to copy based on environment
  let manifestSource = './manifest.xml';
  if (isProduction) {
    manifestSource = './manifest-prod.xml';
  } else if (isStaging) {
    manifestSource = './manifest-staging.xml';
  }

  return {
    plugins: [
      react(),
      viteStaticCopy({
        targets: [
          {
            src: './src/commands/functions.json',
            dest: '.'
          },
          {
            src: './assets',
            dest: '.',
            options: { overwrite: true }
          },
          {
            src: manifestSource,
            dest: '.',
            rename: 'manifest.xml'
          }
        ]
      })
    ],
    
    base: publicPath,
    
    build: {
      outDir: 'dist',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          taskpane: resolve(fileURLToPath(new URL('.', import.meta.url)), 'taskpane.html'),
          commands: resolve(fileURLToPath(new URL('.', import.meta.url)), 'commands.html')
        },
        output: {
          // Use content hash for production builds
          entryFileNames: isProduction ? '[name].[hash].js' : '[name].js',
          chunkFileNames: isProduction ? '[name].[hash].js' : '[name].js',
          assetFileNames: isProduction ? '[name].[hash].[ext]' : '[name].[ext]'
        }
      },
      // Code splitting configuration for production
      chunkSizeWarningLimit: 600,
      target: 'es2015'
    },

    server: {
      port: 3000,
      open: false,
      https: (() => {
        // Check if Office Add-in certificates exist for HTTPS
        if (existsSync(certPath) && existsSync(keyPath)) {
          return {
            key: readFileSync(keyPath),
            cert: readFileSync(certPath)
          };
        } else {
          console.warn('⚠️  Office Add-in certificates not found. Run "npm run cert:install" to install certificates for proper HTTPS support.');
          console.warn('📍 Using default HTTPS configuration as fallback.');
          return true;
        }
      })(),
      cors: {
        origin: '*',
        methods: ['GET', 'POST', 'PUT', 'DELETE', 'PATCH', 'OPTIONS'],
        allowedHeaders: ['X-Requested-With', 'content-type', 'Authorization']
      }
    },

    resolve: {
      extensions: ['.ts', '.tsx', '.js', '.jsx']
    }
  };
});