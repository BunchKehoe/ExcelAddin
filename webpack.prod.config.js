const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
  const isProduction = argv.mode === 'production';
  
  return {
    mode: isProduction ? 'production' : 'development',
    entry: {
      taskpane: './src/taskpane/taskpane.tsx',
      commands: './src/commands/commands.ts'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: isProduction ? '[name].[contenthash].js' : '[name].js',
      clean: true,
      publicPath: '/excellence/'
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
                sourceMap: !isProduction
              }
            }
          },
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg|ico)$/i,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext]'
          }
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane'],
        minify: isProduction ? {
          collapseWhitespace: true,
          removeComments: true,
          removeRedundantAttributes: true,
          removeScriptTypeAttributes: true,
          removeStyleLinkTypeAttributes: true,
          useShortDoctype: true
        } : false
      }),
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: ['commands'],
        minify: isProduction ? {
          collapseWhitespace: true,
          removeComments: true,
          removeRedundantAttributes: true,
          removeScriptTypeAttributes: true,
          removeStyleLinkTypeAttributes: true,
          useShortDoctype: true
        } : false
      }),
      new CopyWebpackPlugin({
        patterns: [
          { 
            from: './src/commands/functions.json', 
            to: 'functions.json' 
          },
          {
            from: './assets',
            to: 'assets',
            noErrorOnMissing: true
          },
          {
            from: './manifest-staging.xml',
            to: 'manifest.xml',
            noErrorOnMissing: true
          }
        ]
      })
    ],
    optimization: {
      splitChunks: isProduction ? {
        chunks: 'all',
        cacheGroups: {
          vendor: {
            test: /[\\/]node_modules[\\/]/,
            name: 'vendors',
            chunks: 'all',
          },
        },
      } : false,
      minimize: isProduction,
      usedExports: true,
      sideEffects: false
    },
    devServer: {
      static: './dist',
      port: 3000,
      open: false,
      hot: true,
      server: 'https',
      allowedHosts: 'all',
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
        'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
      },
      // Development only - use self-signed certificates
      https: !isProduction ? {
        key: './certs/localhost-key.pem',
        cert: './certs/localhost.pem'
      } : true
    },
    // Performance hints for production builds
    performance: {
      hints: isProduction ? 'warning' : false,
      maxAssetSize: 1024 * 1024, // 1MB
      maxEntrypointSize: 1024 * 1024 // 1MB
    },
    // Source maps for debugging
    devtool: isProduction ? false : 'eval-source-map'
  };
};