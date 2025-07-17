import React from 'react';
import {
  Container,
  Typography,
  Box
} from '@mui/material';
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer
} from 'recharts';

const DashboardsPage: React.FC = () => {
  // Mock data for the line chart
  const data = [
    { name: 'Jan', series1: 4000, series2: 2400, series3: 1800 },
    { name: 'Feb', series1: 3000, series2: 1398, series3: 2100 },
    { name: 'Mar', series1: 2000, series2: 9800, series3: 2900 },
    { name: 'Apr', series1: 2780, series2: 3908, series3: 2000 },
    { name: 'May', series1: 1890, series2: 4800, series3: 2181 },
    { name: 'Jun', series1: 2390, series2: 3800, series3: 2500 },
    { name: 'Jul', series1: 3490, series2: 4300, series3: 2100 },
  ];

  return (
    <Container maxWidth="md" sx={{ mt: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom align="center">
        Windpark A
      </Typography>
      
      <Box sx={{ mt: 4, height: 400 }}>
        <ResponsiveContainer width="100%" height="100%">
          <LineChart data={data}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" />
            <YAxis />
            <Tooltip />
            <Legend />
            <Line 
              type="monotone" 
              dataKey="series1" 
              stroke="#8884d8" 
              strokeWidth={2}
              name="Power Generation"
            />
            <Line 
              type="monotone" 
              dataKey="series2" 
              stroke="#82ca9d" 
              strokeWidth={2}
              name="Wind Speed"
            />
            <Line 
              type="monotone" 
              dataKey="series3" 
              stroke="#ffc658" 
              strokeWidth={2}
              name="Efficiency"
            />
          </LineChart>
        </ResponsiveContainer>
      </Box>
    </Container>
  );
};

export default DashboardsPage;