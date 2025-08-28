/**
 * Asset utility functions for environment-aware asset loading
 * Provides consistent asset URL generation across all environments
 */

import { assetBaseUrl, environment } from '../config/environment';

/**
 * Generate a complete asset URL for the current environment
 * @param assetPath - Relative path to the asset (e.g., 'images/logos/PCAG_white_trans.png')
 * @returns Complete asset URL for the current environment
 */
export function getAssetUrl(assetPath: string): string {
  // Remove leading slash if present to avoid double slashes
  const cleanPath = assetPath.startsWith('/') ? assetPath.substring(1) : assetPath;
  
  const fullUrl = `${assetBaseUrl}/${cleanPath}`;
  
  // Development logging for asset URL generation
  if (environment === 'development') {
    console.debug(`🖼️ Asset URL generated:`, {
      environment,
      assetBaseUrl,
      assetPath: cleanPath,
      fullUrl,
      hostname: typeof window !== 'undefined' ? window.location.hostname : 'N/A'
    });
  }
  
  return fullUrl;
}

/**
 * Generate asset URL specifically for logo images
 * @param logoFilename - Logo filename (e.g., 'PCAG_white_trans.png')
 * @returns Complete logo URL for the current environment
 */
export function getLogoUrl(logoFilename: string): string {
  return getAssetUrl(`images/logos/${logoFilename}`);
}

/**
 * Generate asset URL specifically for icon images
 * @param iconFilename - Icon filename (e.g., 'icon-32.png')
 * @returns Complete icon URL for the current environment
 */
export function getIconUrl(iconFilename: string): string {
  return getAssetUrl(`images/icons/${iconFilename}`);
}

/**
 * Verify if asset loading is working correctly for the current environment
 * This function can be used for debugging asset loading issues
 */
export async function verifyAssetLoading(): Promise<{
  environment: string;
  assetBaseUrl: string;
  testResults: Array<{ url: string; status: 'success' | 'error'; error?: string }>;
}> {
  const testAssets = [
    'images/logos/PCAG_white_trans.svg',
    'images/icons/icon-32.png'
  ];
  
  const testResults = await Promise.all(
    testAssets.map(async (asset) => {
      const url = getAssetUrl(asset);
      try {
        const response = await fetch(url, { method: 'HEAD' });
        return {
          url,
          status: response.ok ? 'success' as const : 'error' as const,
          error: response.ok ? undefined : `HTTP ${response.status}`
        };
      } catch (error) {
        return {
          url,
          status: 'error' as const,
          error: error instanceof Error ? error.message : 'Unknown error'
        };
      }
    })
  );
  
  return {
    environment,
    assetBaseUrl,
    testResults
  };
}