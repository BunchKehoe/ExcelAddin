/**
 * Debug component for verifying asset loading in different environments
 * Only renders in development mode
 */

import React, { useEffect, useState } from 'react';
import { Box, Typography, Alert, Collapse, Button } from '@mui/material';
import { environment } from '../config/environment';
import { verifyAssetLoading } from '../utils/assetUtils';

interface AssetTestResult {
  environment: string;
  assetBaseUrl: string;
  testResults: Array<{ url: string; status: 'success' | 'error'; error?: string }>;
}

const AssetDebugPanel: React.FC = () => {
  const [testResults, setTestResults] = useState<AssetTestResult | null>(null);
  const [isExpanded, setIsExpanded] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  const runAssetTest = async () => {
    setIsLoading(true);
    try {
      const results = await verifyAssetLoading();
      setTestResults(results);
    } catch (error) {
      console.error('Asset verification failed:', error);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    // Auto-run test in development
    if (environment === 'development') {
      runAssetTest();
    }
  }, []);

  // Only render in development environment
  if (environment !== 'development') {
    return null;
  }

  const hasErrors = testResults?.testResults.some(result => result.status === 'error');

  return (
    <Box sx={{ position: 'fixed', bottom: 16, right: 16, zIndex: 9999, maxWidth: 400 }}>
      <Button
        variant="outlined"
        size="small"
        onClick={() => setIsExpanded(!isExpanded)}
        sx={{ mb: 1 }}
      >
        🔧 Asset Debug {hasErrors ? '⚠️' : '✅'}
      </Button>
      
      <Collapse in={isExpanded}>
        <Alert 
          severity={hasErrors ? 'warning' : 'success'}
          sx={{ mb: 2 }}
        >
          <Typography variant="subtitle2" gutterBottom>
            Environment: {testResults?.environment || 'Unknown'}
          </Typography>
          <Typography variant="body2" gutterBottom>
            Asset Base URL: {testResults?.assetBaseUrl || 'Not loaded'}
          </Typography>
          
          {testResults && (
            <Box sx={{ mt: 1 }}>
              <Typography variant="subtitle2" gutterBottom>
                Asset Loading Test Results:
              </Typography>
              {testResults.testResults.map((result, index) => (
                <Box key={index} sx={{ fontSize: '0.875rem', mt: 0.5 }}>
                  <Typography 
                    variant="body2" 
                    color={result.status === 'success' ? 'success.main' : 'error.main'}
                  >
                    {result.status === 'success' ? '✅' : '❌'} {result.url}
                    {result.error && ` - ${result.error}`}
                  </Typography>
                </Box>
              ))}
            </Box>
          )}
          
          <Button
            size="small"
            onClick={runAssetTest}
            disabled={isLoading}
            sx={{ mt: 1 }}
          >
            {isLoading ? 'Testing...' : 'Retest Assets'}
          </Button>
        </Alert>
      </Collapse>
    </Box>
  );
};

export default AssetDebugPanel;