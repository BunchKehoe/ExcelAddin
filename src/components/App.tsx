import React, { useState } from 'react';
import {
  AppBar,
  Toolbar,
  Typography,
  Container,
  Button,
  Stack,
  Box
} from '@mui/material';
import { ThemeProvider, createTheme } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import { LocalizationProvider } from '@mui/x-date-pickers/LocalizationProvider';
import { AdapterDayjs } from '@mui/x-date-pickers/AdapterDayjs';
import 'dayjs/locale/de';
import theme from '../theme';

import DatabasePage from './pages/DatabasePage';
import MarketDataPage from './pages/MarketDataPage';
import ApplicationsPage from './pages/ApplicationsPage';
import DashboardsPage from './pages/DashboardsPage';
import ExcelFunctionsPage from './pages/ExcelFunctionsPage';

type Page = 'home' | 'database' | 'market-data' | 'applications' | 'dashboards' | 'excel-functions';

const App: React.FC = () => {
  const [currentPage, setCurrentPage] = useState<Page>('home');

  const renderPage = () => {
    switch (currentPage) {
      case 'database':
        return <DatabasePage />;
      case 'market-data':
        return <MarketDataPage />;
      case 'applications':
        return <ApplicationsPage />;
      case 'dashboards':
        return <DashboardsPage />;
      case 'excel-functions':
        return <ExcelFunctionsPage />;
      default:
        return (
          <Container maxWidth="sm" sx={{ mt: 4 }}>
            <Typography variant="h4" component="h1" gutterBottom align="center">
              Welcome to<br />
              Prime Excellence
            </Typography>
            <Typography variant="body1" paragraph align="center">
              Select a function from the buttons below to get started.
            </Typography>
            <Stack spacing={2} sx={{ mt: 4 }}>
              <Button
                variant="contained"
                fullWidth
                size="large"
                onClick={() => setCurrentPage('database')}
              >
                Raw Admin Data
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
    <LocalizationProvider dateAdapter={AdapterDayjs} adapterLocale="de">
      <CssBaseline />
      <Box sx={{ flexGrow: 1 }}>
        <AppBar position="static">
          <Toolbar sx={{ minHeight: 65 }}>
            <Box sx={{ display: 'flex', alignItems: 'center', flexGrow: 1 }}>
              <Typography variant="h6" component="div" sx={{ display: 'flex', alignItems: 'center' }}>
                <Box component="span" sx={{ fontWeight: 'bold' }}>
                  Excel
                </Box>
                lence
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
