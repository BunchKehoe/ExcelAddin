const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
  const isProduction = argv.mode === 'production';
  const isStaging = env && env.staging;
  const isProd = env && env.production;
  
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
          // Copy the appropriate manifest file based on environment
          isProd ? {
            from: './manifest-prod.xml',
            to: 'manifest.xml',
            noErrorOnMissing: true
          } : isStaging ? {
            from: './manifest-staging.xml',
            to: 'manifest.xml',
            noErrorOnMissing: true
          } : {
            from: './manifest.xml',
            to: 'manifest.xml',
            noErrorOnMissing: true
          }
        ].filter(Boolean)
      })
    ],
    optimization: {
      splitChunks: isProduction ? {
        chunks: 'all',
        minSize: 20000,
        maxSize: 200000,
        cacheGroups: {
          vendor: {
            test: /[\\/]node_modules[\\/]/,
            name: 'vendors',
            chunks: 'all',
            priority: 10,
          },
          mui: {
            test: /[\\/]node_modules[\\/]@mui[\\/]/,
            name: 'mui',
            chunks: 'all',
            priority: 20,
          },
          react: {
            test: /[\\/]node_modules[\\/](react|react-dom)[\\/]/,
            name: 'react',
            chunks: 'all',
            priority: 20,
          },
          recharts: {
            test: /[\\/]node_modules[\\/]recharts[\\/]/,
            name: 'recharts',
            chunks: 'all',
            priority: 20,
          },
          common: {
            minChunks: 2,
            priority: 5,
            reuseExistingChunk: true,
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
      maxAssetSize: 500 * 1024, // 500KB (increased from 1MB)
      maxEntrypointSize: 500 * 1024 // 500KB (increased from 1MB)
    },
    // Source maps for debugging
    devtool: isProduction ? false : 'eval-source-map'
  };
};