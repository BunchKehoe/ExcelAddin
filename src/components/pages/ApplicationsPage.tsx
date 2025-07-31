import React, { useState } from 'react';
import {
  Container,
  Typography,
  Button,
  Stack,
  Snackbar,
  Alert
} from '@mui/material';

const ApplicationsPage: React.FC = () => {
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

  const handleButtonClick = (appName: string) => {
    showNotification(`${appName} functionality coming soon!`, 'info');
  };

  return (
    <Container maxWidth="sm" sx={{ mt: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom align="center">
        Applications
      </Typography>
      
      <Stack spacing={3} sx={{ mt: 4 }}>
        <Button
          variant="contained"
          fullWidth
          size="large"
          onClick={() => handleButtonClick('Kassandra')}
        >
          Kassandra
        </Button>
        
        <Button
          variant="contained"
          fullWidth
          size="large"
          onClick={() => handleButtonClick('Infinity')}
        >
          Infinity
        </Button>
        
        <Button
          variant="contained"
          fullWidth
          size="large"
          onClick={() => handleButtonClick('Pandora')}
        >
          Pandora
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

export default ApplicationsPage;