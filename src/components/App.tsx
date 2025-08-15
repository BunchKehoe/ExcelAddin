import React, { useState, Suspense, lazy } from 'react';
import {
  AppBar,
  Toolbar,
  Typography,
  Container,
  Button,
  Stack,
  Box,
  CircularProgress
} from '@mui/material';
import { ThemeProvider } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import { LocalizationProvider } from '@mui/x-date-pickers/LocalizationProvider';
import { AdapterDayjs } from '@mui/x-date-pickers/AdapterDayjs';
import theme from '../theme';
import { assetBaseUrl } from '../config/environment';

// Lazy load page components to reduce initial bundle size
const DatabasePage = lazy(() => import('./pages/DatabasePage'));
const MarketDataPage = lazy(() => import('./pages/MarketDataPage'));
const ApplicationsPage = lazy(() => import('./pages/ApplicationsPage'));
const DashboardsPage = lazy(() => import('./pages/DashboardsPage'));
const ExcelFunctionsPage = lazy(() => import('./pages/ExcelFunctionsPage'));
const DataUploadPage = lazy(() => import('./pages/DataUploadPage'));

type Page = 'home' | 'data-upload' | 'database' | 'market-data' | 'applications' | 'dashboards' | 'excel-functions';

const App: React.FC = () => {
  const [currentPage, setCurrentPage] = useState<Page>('home');

  const renderPage = () => {
    switch (currentPage) {
      case 'data-upload':
        return (
          <Suspense fallback={<Container><CircularProgress /></Container>}>
            <DataUploadPage />
          </Suspense>
        );
      case 'database':
        return (
          <Suspense fallback={<Container><CircularProgress /></Container>}>
            <DatabasePage />
          </Suspense>
        );
      case 'market-data':
        return (
          <Suspense fallback={<Container><CircularProgress /></Container>}>
            <MarketDataPage />
          </Suspense>
        );
      case 'applications':
        return (
          <Suspense fallback={<Container><CircularProgress /></Container>}>
            <ApplicationsPage />
          </Suspense>
        );
      case 'dashboards':
        return (
          <Suspense fallback={<Container><CircularProgress /></Container>}>
            <DashboardsPage />
          </Suspense>
        );
      case 'excel-functions':
        return (
          <Suspense fallback={<Container><CircularProgress /></Container>}>
            <ExcelFunctionsPage />
          </Suspense>
        );
      default:
        return (
          <Container maxWidth="sm" sx={{ mt: 4 }}>
            <Typography variant="h4" component="h1" gutterBottom align="center">
              Welcome to<br />
              Prime Capital
            </Typography>
            <Typography variant="body1" paragraph align="center">
              Select a function from the buttons below to get started.
            </Typography>
            <Stack spacing={2} sx={{ mt: 4 }}>
              <Button
                variant="contained"
                fullWidth
                size="large"
                onClick={() => setCurrentPage('data-upload')}
              >
                Data Upload
              </Button>
              <Button
                variant="contained"
                fullWidth
                size="large"
                onClick={() => setCurrentPage('database')}
              >
                Raw Data
              </Button>
              <Button
                variant="contained"
                fullWidth
                size="large"
                onClick={() => setCurrentPage('market-data')}
              >
                Market Data
              </Button>
              <Button
                variant="contained"
                fullWidth
                size="large"
                onClick={() => setCurrentPage('applications')}
              >
                Applications
              </Button>
              <Button
                variant="contained"
                fullWidth
                size="large"
                onClick={() => setCurrentPage('dashboards')}
              >
                Dashboards
              </Button>
              <Button
                variant="contained"
                fullWidth
                size="large"
                onClick={() => setCurrentPage('excel-functions')}
              >
                Excel Functions
              </Button>
            </Stack>
          </Container>
        );
    }
  };

  return (
    <ThemeProvider theme={theme}>
    <LocalizationProvider dateAdapter={AdapterDayjs}>
      <CssBaseline />
      <Box sx={{ flexGrow: 1 }}>
        <AppBar position="static">
          <Toolbar sx={{ minHeight: 65 }}>
            <Box sx={{ display: 'flex', alignItems: 'center', flexGrow: 1 }}>
              <Box
                component="img"
                src={`${assetBaseUrl}/PCAG_white_trans.png`}
                alt="Prime Capital Logo"
                sx={{
                  height: 24,
                  width: 24,
                  mr: 1
                }}
              />
              <Typography variant="h6" component="div" sx={{ display: 'flex', alignItems: 'center' }}>
                <Box component="span" sx={{ fontWeight: 'bold' }}>
                  Prime Capital
                </Box>
                Excellence
              </Typography>
            </Box>
            {currentPage !== 'home' && (
              <Button color="inherit" onClick={() => setCurrentPage('home')}>
                Home
              </Button>
            )}
          </Toolbar>
        </AppBar>
        <Box sx={{ minHeight: '100vh-100px' }}>
          {renderPage()}
        </Box>
      </Box>
     </LocalizationProvider>
    </ThemeProvider>
  );
};

export default App;
