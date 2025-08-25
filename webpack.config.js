/**
 * Configuration Webpack pour Office Add-in
 * 
 * Cette configuration gère la compilation et le bundling de l'add-in Office.
 * Elle inclut la gestion des certificats de développement, la compilation Babel,
 * et la génération des fichiers HTML nécessaires.
 * 
 * @author Microsoft Corporation (modifié)
 * @license MIT
 * @version 1.0.0
 */

/* eslint-disable no-undef */

const devCerts = require('office-addin-dev-certs');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const path = require('path');

/**
 * Configuration des URLs de développement et production
 */
const URLS = {
  DEV: 'https://localhost:3001/',
  PROD: 'https://www.contoso.com/' // À modifier pour votre déploiement de production
};

/**
 * Configuration des ports et chemins
 */
const PATHS = {
  SRC: path.resolve(__dirname, 'src'),
  DIST: path.resolve(__dirname, 'dist'),
  ASSETS: path.resolve(__dirname, 'assets'),
  MANIFEST_OUT: path.resolve(__dirname, 'manifest-xml-out')
};

/**
 * Obtient les options HTTPS pour le serveur de développement
 * Utilise les certificats de développement Office Add-in
 * 
 * @returns {Promise<Object>} Options HTTPS avec certificats
 */
async function getHttpsOptions() {
  try {
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return { 
      ca: httpsOptions.ca, 
      key: httpsOptions.key, 
      cert: httpsOptions.cert 
    };
  } catch (error) {
    console.warn('⚠️ Impossible de charger les certificats HTTPS:', error.message);
    console.warn('⚠️ Le serveur de développement utilisera HTTP');
    return {};
  }
}

/**
 * Configuration principale de Webpack
 * 
 * @param {Object} env - Variables d'environnement
 * @param {Object} options - Options de compilation
 * @returns {Promise<Object>} Configuration Webpack
 */
module.exports = async (env, options) => {
  // Déterminer le mode de compilation
  const isDevelopment = options.mode === 'development';
  
  console.log(`🔧 Configuration Webpack en mode: ${options.mode}`);
  console.log(`📁 Dossier source: ${PATHS.SRC}`);
  console.log(`📁 Dossier de sortie: ${PATHS.DIST}`);

  // Configuration de base
  const config = {
    // Source maps pour le débogage
    devtool: isDevelopment ? 'source-map' : false,
    
    // Points d'entrée de l'application
    entry: {
      // Polyfills pour la compatibilité des navigateurs
      polyfill: [
        'core-js/stable', 
        'regenerator-runtime/runtime'
      ],
      
      // Interface principale de l'add-in
      taskpane: [
        './src/taskpane/taskpane.js',
        './src/taskpane/taskpane.html'
      ],
      
      // Commandes du ruban Office
      commands: './src/commands/commands.js'
    },
    
    // Configuration de sortie
    output: {
      path: PATHS.DIST,
      filename: '[name].js',
      clean: true, // Nettoie le dossier de sortie avant chaque build
      publicPath: '/'
    },
    
    // Résolution des modules
    resolve: {
      extensions: ['.html', '.js', '.json'],
      alias: {
        '@': PATHS.SRC,
        '@assets': PATHS.ASSETS
      }
    },
    
    // Règles de traitement des fichiers
    module: {
      rules: [
        // Traitement des fichiers JavaScript avec Babel
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
            options: {
              presets: ['@babel/preset-env'],
              plugins: [
                // Support des fonctionnalités ES6+
                '@babel/plugin-proposal-class-properties',
                '@babel/plugin-proposal-object-rest-spread'
              ]
            }
          }
        },
        
        // Traitement des fichiers HTML
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: {
            loader: 'html-loader',
            options: {
              minimize: !isDevelopment,
              removeComments: !isDevelopment
            }
          }
        },
        
        // Gestion des assets (images, icônes)
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext][query]'
          }
        },
        
        // Gestion des polices
        {
          test: /\.(woff|woff2|eot|ttf|otf)$/,
          type: 'asset/resource',
          generator: {
            filename: 'fonts/[name][ext][query]'
          }
        },
        
        // Gestion des fichiers CSS
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        }
      ]
    },
    
    // Plugins de compilation
    plugins: [
      // Plugin principal pour la page taskpane
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['polyfill', 'taskpane'],
        minify: !isDevelopment,
        inject: 'body'
      }),
      
      // Page d'authentification MSAL (ouverte en Dialog)
      new HtmlWebpackPlugin({
        filename: 'auth.html',
        template: './src/taskpane/auth.html',
        chunks: [], // Pas de chunks pour cette page
        minify: !isDevelopment,
        inject: 'body'
      }),
      
      // Page de callback pour le consentement administrateur
      new HtmlWebpackPlugin({
        filename: 'consent/callback.html',
        template: './src/taskpane/consent-callback.html',
        chunks: [],
        minify: !isDevelopment,
        inject: 'body'
      }),
      
      // Page des commandes du ruban Office
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['polyfill', 'commands'],
        minify: !isDevelopment,
        inject: 'body'
      }),
      
      // Plugin de copie des assets
      new CopyWebpackPlugin({
        patterns: [
          {
            from: PATHS.ASSETS,
            to: 'assets',
            globOptions: {
              ignore: ['**/.*'] // Ignorer les fichiers cachés
            }
          }
        ]
      })
    ],
    
    // Configuration du serveur de développement
    devServer: {
      // Headers CORS pour le développement
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
        'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
      },
      
      // Configuration du serveur HTTPS
      server: {
        type: 'https',
        options: env.WEBPACK_BUILD || options.https !== undefined 
          ? options.https 
          : await getHttpsOptions()
      },
      
      // Port du serveur de développement
      port: process.env.npm_package_config_dev_server_port || 3001,
      
      // Configuration pour gérer les routes personnalisées
      historyApiFallback: {
        rewrites: [
          {
            from: /^\/consent\/callback/,
            to: '/consent/callback.html'
          },
          {
            from: /^\/auth/,
            to: '/auth.html'
          }
        ]
      },
      
      // Configuration des fichiers statiques
      static: {
        directory: PATHS.DIST,
        publicPath: '/',
        watch: true
      },
      
      // Configuration du hot reload
      hot: isDevelopment,
      liveReload: isDevelopment,
      
      // Configuration des logs
      client: {
        logging: isDevelopment ? 'info' : 'warn',
        overlay: {
          errors: true,
          warnings: false
        }
      },
      
      // Configuration de la compression
      compress: true,
      
      // Configuration des timeouts
      timeout: 30000,
      
      // Configuration des erreurs
      onErrors: (devServer) => {
        console.error('❌ Erreur du serveur de développement:', devServer);
      }
    },
    
    // Configuration des optimisations
    optimization: {
      // Minimisation en production
      minimize: !isDevelopment,
      
      // Séparation des chunks
      splitChunks: {
        chunks: 'all',
        cacheGroups: {
          vendor: {
            test: /[\\/]node_modules[\\/]/,
            name: 'vendors',
            chunks: 'all'
          }
        }
      }
    },
    
    // Configuration des performances
    performance: {
      hints: isDevelopment ? false : 'warning',
      maxEntrypointSize: 512000,
      maxAssetSize: 512000
    }
  };

  // Ajouter des plugins spécifiques au mode développement
  if (isDevelopment) {
    console.log('🚀 Mode développement activé');
    
    // Ajouter des plugins de développement si nécessaire
    // config.plugins.push(new webpack.HotModuleReplacementPlugin());
  } else {
    console.log('🏭 Mode production activé');
    
    // Optimisations spécifiques à la production
    config.optimization.minimize = true;
  }

  console.log('✅ Configuration Webpack générée avec succès');
  return config;
};
