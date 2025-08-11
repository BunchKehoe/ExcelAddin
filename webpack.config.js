const path = require('path');
const { readFileSync, existsSync } = require('fs');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  mode: 'development',
  entry: {
    taskpane: './src/taskpane/taskpane.tsx',
    commands: './src/commands/commands.ts'
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
    clean: true
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx']
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: {
          loader: 'ts-loader',
          options: {
            compilerOptions: {
              noEmit: false,
              declaration: false,
              declarationMap: false,
              sourceMap: false
            }
          }
        },
        exclude: /node_modules/
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/taskpane/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane']
    }),
    new HtmlWebpackPlugin({
      template: './src/commands/commands.html',
      filename: 'commands.html',
      chunks: ['commands']
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: './src/commands/functions.json', to: 'functions.json' },
        {
          from: './assets',
          to: 'assets',
          noErrorOnMissing: true
        }
      ]
    })
  ],
  devServer: {
    static: './dist',
    port: 3000,
    open: false,
    hot: true,
    server: (() => {
      const certPath = require('os').homedir() + '/.office-addin-dev-certs/localhost.crt';
      const keyPath = require('os').homedir() + '/.office-addin-dev-certs/localhost.key';
      
      // Check if Office Add-in certificates exist
      if (existsSync(certPath) && existsSync(keyPath)) {
        return {
          type: 'https',
          options: {
            key: readFileSync(keyPath),
            cert: readFileSync(certPath)
          }
        };
      } else {
        console.warn('⚠️  Office Add-in certificates not found. Run "npm run cert:install" to install certificates for proper HTTPS support.');
        console.warn('📍 Using default HTTPS configuration as fallback.');
        return 'https';
      }
    })(),
    allowedHosts: 'all',
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
      'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
    }
  }
};