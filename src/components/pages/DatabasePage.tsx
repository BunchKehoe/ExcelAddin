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
  SelectChangeEvent
} from '@mui/material';
import { downloadData } from '../api/apiClient';

const DatabasePage: React.FC = () => {
  const [funds, setFunds] = useState<string[]>([]);
  const [selectedFund, setSelectedFund] = useState<string>('');
  const [dataType, setDataType] = useState<string>('');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);

  useEffect(() => {
    // Mock data for funds - in real implementation, this would come from SQL database
    const mockFunds = [
      'Global Equity Fund',
      'Fixed Income Fund',
      'Emerging Markets Fund',
      'Technology Fund',
      'Real Estate Fund'
    ];
    setFunds(mockFunds);
  }, []);

  const handleFundChange = (event: SelectChangeEvent<string>) => {
    setSelectedFund(event.target.value);
  };

  const handleDataTypeChange = (event: SelectChangeEvent<string>) => {
    setDataType(event.target.value);
  };

  const handleDownload = async () => {
    if (!selectedFund || !dataType || !startDate || !endDate) {
      alert('Please fill in all fields');
      return;
    }

    setLoading(true);
    try {
      const data = await downloadData({
        fund: selectedFund,
        dataType,
        startDate,
        endDate
      });
      
      // Insert data into Excel
      await insertDataIntoExcel(data);
      
    } catch (error) {
      console.error('Error downloading data:', error);
      alert('Error downloading data');
    } finally {
      setLoading(false);
    }
  };

  const insertDataIntoExcel = async (data: any) => {
    try {
      // Check if Excel is available
      if (typeof Excel === 'undefined' || !Excel.run) {
        alert('Excel integration not available in development mode. Data would be inserted into Excel in production.');
        return;
      }

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange('A1');
        
        // Mock financial data structure
        const headers = ['Date', 'Value', 'Change', 'Percentage'];
        const sampleData = [
          ['2024-01-01', '100.00', '0.00', '0.00%'],
          ['2024-01-02', '101.50', '1.50', '1.50%'],
          ['2024-01-03', '99.75', '-1.75', '-1.72%'],
          ['2024-01-04', '102.25', '2.50', '2.51%'],
          ['2024-01-05', '98.90', '-3.35', '-3.28%']
        ];

        // Insert headers
        sheet.getRange('A1:D1').values = [headers];
        
        // Insert data
        const dataRange = sheet.getRange(`A2:D${sampleData.length + 1}`);
        dataRange.values = sampleData;
        
        // Format headers
        const headerRange = sheet.getRange('A1:D1');
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = '#4472C4';
        headerRange.format.font.color = 'white';
        
        // Auto-fit columns
        sheet.getUsedRange().format.autofitColumns();
        
        await context.sync();
      });
      
      alert('Data successfully inserted into Excel!');
    } catch (error) {
      console.error('Error inserting data into Excel:', error);
      alert('Error inserting data into Excel');
    }
  };

  return (
    <Container maxWidth="sm" sx={{ mt: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom align="center">
        KVG Data
      </Typography>
      
      <Stack spacing={3} sx={{ mt: 4 }}>
        <FormControl fullWidth>
          <InputLabel>Select Fund</InputLabel>
          <Select
            value={selectedFund}
            label="Select Fund"
            onChange={handleFundChange}
          >
            {funds.map((fund) => (
              <MenuItem key={fund} value={fund}>
                {fund}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {selectedFund && (
          <Typography variant="h6" component="h2" sx={{ mt: 2 }}>
            {selectedFund}
          </Typography>
        )}

        <FormControl fullWidth>
          <InputLabel>Data Type</InputLabel>
          <Select
            value={dataType}
            label="Data Type"
            onChange={handleDataTypeChange}
          >
            <MenuItem value="NAV">NAV</MenuItem>
            <MenuItem value="Balance Sheet">Balance Sheet</MenuItem>
            <MenuItem value="Transactions">Transactions</MenuItem>
          </Select>
        </FormControl>

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
          disabled={loading}
          sx={{ mt: 3 }}
        >
          {loading ? 'Downloading...' : 'Download Data'}
        </Button>
      </Stack>
    </Container>
  );
};

export default DatabasePage;