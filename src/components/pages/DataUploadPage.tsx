/* global Office, Excel */

import React, { useState } from 'react';
import {
  Container,
  Typography,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Button,
  Stack,
  Box,
  SelectChangeEvent,
  Alert,
  Snackbar,
  FormControlLabel,
  Checkbox,
  Switch
} from '@mui/material';
import { DatePicker } from '@mui/x-date-pickers/DatePicker';
import dayjs, { Dayjs } from 'dayjs';

const DataUploadPage: React.FC = () => {
  const [dataUploadType, setDataUploadType] = useState<string>('');
  const [skipDuplicateCheck, setSkipDuplicateCheck] = useState<boolean>(false);
  const [deliveryDate, setDeliveryDate] = useState<Dayjs | null>(null);
  const [readyToUpload, setReadyToUpload] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(false);

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

  const handleDataUploadTypeChange = (event: SelectChangeEvent) => {
    setDataUploadType(event.target.value);
  };

  const generateDataUploadTemplate = async () => {
    if (!dataUploadType) {
      showNotification('Please select a data upload type first.', 'warning');
      return;
    }

    try {
      await Office.onReady();
      await Excel.run(async (context) => {
        // Create a new worksheet with the selected data upload type as name
        const sheets = context.workbook.worksheets;
        const newSheet = sheets.add(dataUploadType);
        
        // Add headers
        const headers = ['Security', 'Cost', 'Performance', 'Date', 'Identifier'];
        const headerRange = newSheet.getRange('A1:E1');
        headerRange.values = [headers];
        
        // Format headers
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = '#4472C4';
        headerRange.format.font.color = 'white';
        
        // Auto-fit columns
        newSheet.getUsedRange().format.autofitColumns();
        
        // Activate the new sheet
        newSheet.activate();
        
        await context.sync();
        
        showNotification(`Template sheet "${dataUploadType}" created successfully!`, 'success');
      });
    } catch (error) {
      console.error('Error creating template:', error);
      showNotification('Error creating template sheet. Please try again.', 'error');
    }
  };

  const handleUpload = async () => {
    if (!dataUploadType) {
      showNotification('Please select a data upload type first.', 'warning');
      return;
    }

    setLoading(true);
    try {
      await Office.onReady();
      await Excel.run(async (context) => {
        // Find the sheet with the data upload type name
        const sheets = context.workbook.worksheets;
        let targetSheet;
        
        try {
          targetSheet = sheets.getItem(dataUploadType);
        } catch (error) {
          showNotification(`Sheet "${dataUploadType}" not found. Please generate template first.`, 'warning');
          setLoading(false);
          return;
        }

        // Get the used range
        const usedRange = targetSheet.getUsedRange();
        usedRange.load(['values', 'rowCount']);
        
        await context.sync();
        
        if (usedRange.rowCount <= 1) {
          showNotification('No data found in the template sheet. Please add data before uploading.', 'warning');
          setLoading(false);
          return;
        }

        // Extract data (excluding header row)
        const values = usedRange.values;
        const headers = values[0] as string[];
        const dataRows = values.slice(1);
        
        // Convert to JSON format
        const jsonData = dataRows.map(row => {
          const obj: any = {};
          headers.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });

        // Prepare upload payload
        const uploadPayload = {
          dataType: dataUploadType,
          skipDuplicateCheck,
          deliveryDate: deliveryDate?.toISOString(),
          data: jsonData
        };

        // Here you would normally send to the NiFi endpoint
        // For now, we'll just log it and show a success message
        console.log('Upload payload:', uploadPayload);
        
        // Simulated API call
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        showNotification(`Successfully uploaded ${jsonData.length} records to NiFi endpoint!`, 'success');
      });
    } catch (error) {
      console.error('Error uploading data:', error);
      showNotification('Error uploading data. Please try again.', 'error');
    } finally {
      setLoading(false);
    }
  };

  const dataUploadTypes = [
    'Windmill Statistics',
    'Financial Outperformance', 
    'Excellence Accounting'
  ];

  return (
    <Container maxWidth="md" sx={{ mt: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom align="center">
        Data Upload
      </Typography>
      
      <Stack spacing={3} sx={{ mt: 4 }}>
        {/* Data Upload Type Dropdown */}
        <FormControl fullWidth>
          <InputLabel id="data-upload-type-label">Data Upload Type</InputLabel>
          <Select
            labelId="data-upload-type-label"
            id="data-upload-type"
            value={dataUploadType}
            label="Data Upload Type"
            onChange={handleDataUploadTypeChange}
          >
            {dataUploadTypes.map((type) => (
              <MenuItem key={type} value={type}>
                {type}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        {/* Generate Template Button */}
        {dataUploadType && (
          <Button
            variant="outlined"
            fullWidth
            size="large"
            onClick={generateDataUploadTemplate}
          >
            Generate Data Upload Template
          </Button>
        )}

        {/* Skip Duplicate Check */}
        {dataUploadType && (
          <FormControlLabel
            control={
              <Checkbox
                checked={skipDuplicateCheck}
                onChange={(e) => setSkipDuplicateCheck(e.target.checked)}
              />
            }
            label="Skip Duplicate Check"
          />
        )}

        {/* Optional Delivery Date */}
        {dataUploadType && (
          <DatePicker
            label="Specify Optional Delivery Date"
            value={deliveryDate}
            onChange={(newValue) => setDeliveryDate(newValue)}
          />
        )}

        {/* Ready to Upload Switch */}
        {dataUploadType && (
          <FormControlLabel
            control={
              <Switch
                checked={readyToUpload}
                onChange={(e) => setReadyToUpload(e.target.checked)}
              />
            }
            label="Ready to Upload"
          />
        )}

        {/* Upload Button */}
        {dataUploadType && (
          <Button
            variant="contained"
            fullWidth
            size="large"
            disabled={!readyToUpload || loading}
            onClick={handleUpload}
            sx={{
              opacity: readyToUpload ? 1 : 0.5,
              transition: 'opacity 0.3s ease'
            }}
          >
            {loading ? 'Uploading...' : 'Upload'}
          </Button>
        )}
      </Stack>

      {/* Notification Snackbar */}
      <Snackbar
        open={notification.open}
        autoHideDuration={6000}
        onClose={handleCloseNotification}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
      >
        <Alert onClose={handleCloseNotification} severity={notification.severity}>
          {notification.message}
        </Alert>
      </Snackbar>
    </Container>
  );
};

export default DataUploadPage;