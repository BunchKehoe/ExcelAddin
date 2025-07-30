/* global Office, Excel */

import React, { useState, useEffect } from 'react';
import {
  Container,
  Typography,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  TextField,
  Button,
  Stack,
  Box,
  SelectChangeEvent,
  Alert,
  Snackbar
} from '@mui/material';
import { getMarketDataSecurities, getMarketDataFields, downloadMarketData } from '../api/apiClient';

const MarketDataPage: React.FC = () => {
  const [securities, setSecurities] = useState<string[]>([]);
  const [fields, setFields] = useState<string[]>([]);
  const [selectedSecurity, setSelectedSecurity] = useState<string>('');
  const [selectedField, setSelectedField] = useState<string>('');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  
  // Notification state
  const [notification, setNotification] = useState<{
    open: boolean;
    message: string;
    severity: 'error' | 'warning' | 'info' | 'success';
  }>({
    open: false,
    message: '',
    severity: 'info'
  });

  const showNotification = (message: string, severity: 'error' | 'warning' | 'info' | 'success' = 'info') => {
    setNotification({
      open: true,
      message,
      severity
    });
  };

  const handleCloseNotification = () => {
    setNotification(prev => ({ ...prev, open: false }));
  };

  useEffect(() => {
    // Load securities when component mounts
    loadSecurities();
  }, []);

  const loadSecurities = async () => {
    try {
      const response = await getMarketDataSecurities();
      if (response.success) {
        setSecurities(response.data);
      } else {
        setError('Failed to load securities');
      }
    } catch (error) {
      console.error('Error loading securities:', error);
      setError('Failed to connect to backend. Make sure the Flask server is running on localhost:5000');
      // Fallback to mock data
      setSecurities(['AAPL US Equity', 'MSFT US Equity', 'GOOGL US Equity']);
    }
  };

  const loadFields = async (security: string) => {
    try {
      setLoading(true);
      const response = await getMarketDataFields(security);
      if (response.success) {
        setFields(response.data);
      } else {
        setError('Failed to load fields');
      }
    } catch (error) {
      console.error('Error loading fields:', error);
      setError('Failed to load fields for selected security');
      // Fallback to mock data
      setFields(['PX_LAST', 'PX_OPEN', 'PX_HIGH', 'PX_LOW', 'PX_VOLUME']);
    } finally {
      setLoading(false);
    }
  };

  const handleSecurityChange = (event: SelectChangeEvent<string>) => {
    const security = event.target.value;
    setSelectedSecurity(security);
    setSelectedField(''); // Reset field selection
    setFields([]); // Clear fields
    setError('');
    
    if (security) {
      loadFields(security);
    }
  };

  const handleFieldChange = (event: SelectChangeEvent<string>) => {
    setSelectedField(event.target.value);
  };

  const handleDownload = async () => {
    if (!selectedSecurity || !selectedField || !startDate || !endDate) {
      showNotification('Please fill in all fields', 'warning');
      return;
    }

    setLoading(true);
    setError('');
    
    try {
      const data = await downloadMarketData({
        security: selectedSecurity,
        field: selectedField,
        start_date: startDate,
        end_date: endDate
      });
      
      if (data.success) {
        // Insert data into Excel
        await insertDataIntoExcel(data.data);
      } else {
        setError(data.error || 'Failed to download data');
      }
      
    } catch (error) {
      console.error('Error downloading data:', error);
      setError('Error downloading data from backend');
      // Fallback to mock data for demo
      const mockData = [
        { date: '2024-01-01', security: selectedSecurity, field: selectedField, value: 150.00 },
        { date: '2024-01-02', security: selectedSecurity, field: selectedField, value: 152.50 },
        { date: '2024-01-03', security: selectedSecurity, field: selectedField, value: 148.75 }
      ];
      await insertDataIntoExcel(mockData);
    } finally {
      setLoading(false);
    }
  };

  const insertDataIntoExcel = async (data: any[]) => {
    try {
      // Check if Excel is available
      if (typeof Excel === 'undefined' || !Excel.run) {
        showNotification(`Excel integration not available in development mode. Would insert ${data.length} market data records into Excel in production.`, 'info');
        console.log('Market data to be inserted:', data);
        return;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        if (data.length === 0) {
          showNotification('No data to insert', 'warning');
          return;
        }

        // Get headers from first data object
        const headers = Object.keys(data[0]);
        const rows = data.map(row => headers.map(header => {
          const value = row[header];
          // Format dates and values appropriately
          if (header === 'date' && value) {
            return new Date(value).toISOString().split('T')[0];
          }
          return value || '';
        }));
        
        // Insert headers
        const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`);
        headerRange.values = [headers];
        
        // Insert data
        if (rows.length > 0) {
          const dataRange = sheet.getRange(`A2:${String.fromCharCode(64 + headers.length)}${rows.length + 1}`);
          dataRange.values = rows;
        }
        
        // Format headers
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = '#2E7D32';
        headerRange.format.font.color = 'white';
        
        // Auto-fit columns
        sheet.getUsedRange().format.autofitColumns();
        
        await context.sync();
      });
      
      showNotification(`Successfully inserted ${data.length} market data records into Excel!`, 'success');
    } catch (error) {
      console.error('Error inserting data into Excel:', error);
      showNotification('Error inserting data into Excel', 'error');
    }
  };

  return (
    <Container maxWidth="sm" sx={{ mt: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom align="center">
        Market Data
      </Typography>
      
      {error && (
        <Alert severity="error" sx={{ mb: 2 }}>
          {error}
        </Alert>
      )}
      
      <Stack spacing={3} sx={{ mt: 4 }}>
        <FormControl fullWidth>
          <InputLabel>Select Security</InputLabel>
          <Select
            value={selectedSecurity}
            label="Select Security"
            onChange={handleSecurityChange}
          >
            {securities.map((security) => (
              <MenuItem key={security} value={security}>
                {security}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        <FormControl fullWidth disabled={!selectedSecurity || loading}>
          <InputLabel>Select Field</InputLabel>
          <Select
            value={selectedField}
            label="Select Field"
            onChange={handleFieldChange}
          >
            {fields.map((field) => (
              <MenuItem key={field} value={field}>
                {field}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {selectedField && (
          <Typography variant="h6" component="h2" sx={{ mt: 2 }}>
            {selectedSecurity} - {selectedField}
          </Typography>
        )}

        <TextField
          fullWidth
          label="Start Date"
          type="date"
          value={startDate}
          onChange={(e) => setStartDate(e.target.value)}
          InputLabelProps={{ shrink: true }}
        />

        <TextField
          fullWidth
          label="End Date"
          type="date"
          value={endDate}
          onChange={(e) => setEndDate(e.target.value)}
          InputLabelProps={{ shrink: true }}
        />

        <Button
          variant="contained"
          fullWidth
          size="large"
          onClick={handleDownload}
          disabled={loading || !selectedSecurity || !selectedField || !startDate || !endDate}
          sx={{ mt: 3, backgroundColor: '#2E7D32', '&:hover': { backgroundColor: '#1B5E20' } }}
        >
          {loading ? 'Downloading...' : 'Download Market Data'}
        </Button>
      </Stack>

      {/* Notification Snackbar */}
      <Snackbar
        open={notification.open}
        autoHideDuration={6000}
        onClose={handleCloseNotification}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
      >
        <Alert
          onClose={handleCloseNotification}
          severity={notification.severity}
          sx={{ width: '100%' }}
        >
          {notification.message}
        </Alert>
      </Snackbar>
    </Container>
  );
};

export default MarketDataPage;