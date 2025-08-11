/**
 * Environment configuration for the Excel Add-in
 * This file handles dynamic environment detection and URL configuration
 */

export interface EnvironmentConfig {
  apiBaseUrl: string;
  environment: 'development' | 'staging' | 'production';
  manifestUrl: string;
  assetBaseUrl: string;
}

/**
 * Determine the current environment based on hostname and URL patterns
 */
function detectEnvironment(): EnvironmentConfig['environment'] {
  if (typeof window === 'undefined') {
    // Server-side or Node.js environment - default to development
    return 'development';
  }

  const hostname = window.location.hostname;
  const origin = window.location.origin;
  
  // Development environment
  if (hostname === 'localhost' || hostname === '127.0.0.1') {
    return 'development';
  }
  
  // Production environment
  if (hostname === 'server-vs84.intranet.local') {
    return 'production';
  }
  
  // Staging environment
  if (hostname === 'server-vs81t.intranet.local') {
    return 'staging';
  }
  
  // Default to development for unknown hosts
  console.warn(`Unknown hostname: ${hostname}, defaulting to development environment`);
  return 'development';
}

/**
 * Get environment-specific configuration
 */
function getEnvironmentConfig(): EnvironmentConfig {
  const env = detectEnvironment();
  
  switch (env) {
    case 'development':
      return {
        apiBaseUrl: 'http://localhost:5000/api',
        environment: 'development',
        manifestUrl: 'https://localhost:3000/manifest.xml',
        assetBaseUrl: 'https://localhost:3000/assets'
      };
      
    case 'staging':
      return {
        apiBaseUrl: 'https://server-vs81t.intranet.local:9443/excellence/api',
        environment: 'staging',
        manifestUrl: 'https://server-vs81t.intranet.local:9443/excellence/manifest.xml',
        assetBaseUrl: 'https://server-vs81t.intranet.local:9443/excellence/assets'
      };
      
    case 'production':
      return {
        apiBaseUrl: 'https://server-vs84.intranet.local:9443/excellence/api',
        environment: 'production',
        manifestUrl: 'https://server-vs84.intranet.local:9443/excellence/manifest.xml',
        assetBaseUrl: 'https://server-vs84.intranet.local:9443/excellence/assets'
      };
      
    default:
      throw new Error(`Unknown environment: ${env}`);
  }
}

// Export the configuration
export const config = getEnvironmentConfig();

// Export individual properties for convenience
export const { apiBaseUrl, environment, manifestUrl, assetBaseUrl } = config;

// Development helper to log current environment
if (environment === 'development') {
  console.log('🔧 Excel Add-in Environment Configuration:', {
    environment,
    apiBaseUrl,
    manifestUrl,
    assetBaseUrl,
    hostname: typeof window !== 'undefined' ? window.location.hostname : 'N/A'
  });
}