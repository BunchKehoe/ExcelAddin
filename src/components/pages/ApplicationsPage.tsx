import React from 'react';
import {
  Container,
  Typography,
  Button,
  Stack
} from '@mui/material';

const ApplicationsPage: React.FC = () => {
  const handleButtonClick = (appName: string) => {
    alert(`${appName} functionality coming soon!`);
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
    </Container>
  );
};

export default ApplicationsPage;