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
  Alert
} from '@mui/material';
import { getRawDataCategories, getRawDataFunds, downloadRawData } from '../api/apiClient';

const DatabasePage: React.FC = () => {
  const [categories, setCategories] = useState<string[]>([]);
  const [funds, setFunds] = useState<string[]>([]);
  const [selectedCategory, setSelectedCategory] = useState<string>('');
  const [selectedFund, setSelectedFund] = useState<string>('');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [fundFilteringAvailable, setFundFilteringAvailable] = useState<boolean>(true);

  useEffect(() => {
    // Load categories when component mounts
    loadCategories();
  }, []);

  const loadCategories = async () => {
    try {
      const response = await getRawDataCategories();
      if (response.success) {
        setCategories(response.data);
      } else {
        setError('Failed to load categories');
      }
    } catch (error) {
      console.error('Error loading categories:', error);
      setError('Failed to connect to backend. Make sure the Flask server is running on localhost:5000');
      // Fallback to mock data
      setCategories(['SAMPLE_CATEGORY_1', 'SAMPLE_CATEGORY_2']);
    }
  };

  const loadFunds = async (category: string) => {
    try {
      setLoading(true);
      const response = await getRawDataFunds(category);
      if (response.success) {
        setFunds(response.data);
        setFundFilteringAvailable(response.fund_filtering_available !== false);
      } else {
        setError('Failed to load funds');
      }
    } catch (error) {
      console.error('Error loading funds:', error);
      setError('Failed to load funds for selected category');
      // Fallback to mock data
      setFunds(['SAMPLE_FUND_1', 'SAMPLE_FUND_2']);
      setFundFilteringAvailable(true);
    } finally {
      setLoading(false);
    }
  };

  const handleCategoryChange = (event: SelectChangeEvent<string>) => {
    const category = event.target.value;
    setSelectedCategory(category);
    setSelectedFund(''); // Reset fund selection
    setFunds([]); // Clear funds
    setFundFilteringAvailable(true); // Reset fund filtering availability
    setError('');
    
    if (category) {
      loadFunds(category);
    }
  };

  const handleFundChange = (event: SelectChangeEvent<string>) => {
    setSelectedFund(event.target.value);
  };

  const handleDownload = async () => {
    if (!selectedCategory || !startDate || !endDate) {
      alert('Please fill in all required fields');
      return;
    }

    if (fundFilteringAvailable && !selectedFund) {
      alert('Please select a fund for this category');
      return;
    }

    setLoading(true);
    setError('');
    
    try {
      const requestData: any = {
        catalog: selectedCategory,
        start_date: startDate,
        end_date: endDate
      };

      // Only include fund if filtering is available
      if (fundFilteringAvailable) {
        requestData.fund = selectedFund;
      }

      const data = await downloadRawData(requestData);
      
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
        { date: '2024-01-01', value: 100.00, fund: selectedFund || 'N/A' },
        { date: '2024-01-02', value: 101.50, fund: selectedFund || 'N/A' },
        { date: '2024-01-03', value: 99.75, fund: selectedFund || 'N/A' }
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
        alert(`Excel integration not available in development mode. Would insert ${data.length} records into Excel in production.`);
        console.log('Data to be inserted:', data);
        return;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        if (data.length === 0) {
          alert('No data to insert');
          return;
        }

        // Get headers from first data object
        const headers = Object.keys(data[0]);
        const rows = data.map(row => headers.map(header => row[header] || ''));
        
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
        headerRange.format.fill.color = '#4472C4';
        headerRange.format.font.color = 'white';
        
        // Auto-fit columns
        sheet.getUsedRange().format.autofitColumns();
        
        await context.sync();
      });
      
      alert(`Successfully inserted ${data.length} records into Excel!`);
    } catch (error) {
      console.error('Error inserting data into Excel:', error);
      alert('Error inserting data into Excel');
    }
  };

  return (
    <Container maxWidth="sm" sx={{ mt: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom align="center">
        Raw Database Tables
      </Typography>
      
      {error && (
        <Alert severity="error" sx={{ mb: 2 }}>
          {error}
        </Alert>
      )}
      
      <Stack spacing={3} sx={{ mt: 4 }}>
        <FormControl fullWidth>
          <InputLabel>Select Category</InputLabel>
          <Select
            value={selectedCategory}
            label="Select Category"
            onChange={handleCategoryChange}
          >
            {categories.map((category) => (
              <MenuItem key={category} value={category}>
                {category}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        <FormControl fullWidth disabled={!selectedCategory || loading || !fundFilteringAvailable}>
          <InputLabel>
            {fundFilteringAvailable ? 'Select Fund' : 'Fund filtering not available for this category'}
          </InputLabel>
          <Select
            value={selectedFund}
            label={fundFilteringAvailable ? 'Select Fund' : 'Fund filtering not available for this category'}
            onChange={handleFundChange}
          >
            {funds.map((fund) => (
              <MenuItem key={fund} value={fund}>
                {fund}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {selectedCategory && fundFilteringAvailable && selectedFund && (
          <Typography variant="h6" component="h2" sx={{ mt: 2 }}>
            {selectedFund} - {selectedCategory}
          </Typography>
        )}

        {selectedCategory && !fundFilteringAvailable && (
          <Typography variant="h6" component="h2" sx={{ mt: 2, color: 'text.secondary' }}>
            {selectedCategory} (No fund filtering available)
          </Typography>
        )}

        <TextField
          fullWidth
          label="Delivery Start"
          type="date"
          value={startDate}
          onChange={(e) => setStartDate(e.target.value)}
          InputLabelProps={{ shrink: true }}
        />

        <TextField
          fullWidth
          label="Delivery End"
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
          disabled={loading || !selectedCategory || !startDate || !endDate || (fundFilteringAvailable && !selectedFund)}
          sx={{ mt: 3 }}
        >
          {loading ? 'Downloading...' : 'Download Data'}
        </Button>
      </Stack>
    </Container>
  );
};

export default DatabasePage;