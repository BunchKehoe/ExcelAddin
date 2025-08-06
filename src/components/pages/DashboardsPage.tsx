import React from 'react';
import {
  Container,
  Typography,
  Box,
  Paper,
  Stack
} from '@mui/material';

// Simple SVG-based chart component instead of heavy Recharts
const SimpleLineChart: React.FC<{ data: any[] }> = ({ data }) => {
  const width = 400;
  const height = 200;
  const padding = 40;
  
  const maxValue = Math.max(...data.map(d => Math.max(d.series1, d.series2, d.series3)));
  const minValue = 0;
  
  const getX = (index: number) => padding + (index * (width - 2 * padding)) / (data.length - 1);
  const getY = (value: number) => height - padding - ((value - minValue) * (height - 2 * padding)) / (maxValue - minValue);
  
  const createPath = (seriesKey: string) => {
    return data.map((d, i) => `${i === 0 ? 'M' : 'L'} ${getX(i)} ${getY(d[seriesKey])}`).join(' ');
  };
  
  return (
    <Paper sx={{ p: 2, height: '100%' }}>
      <Typography variant="h6" gutterBottom>Windpark Performance</Typography>
      <svg width={width} height={height} style={{ border: '1px solid #eee' }}>
        {/* Grid lines */}
        {[0, 1, 2, 3, 4].map(i => (
          <line
            key={`grid-${i}`}
            x1={padding}
            y1={padding + i * (height - 2 * padding) / 4}
            x2={width - padding}
            y2={padding + i * (height - 2 * padding) / 4}
            stroke="#f0f0f0"
            strokeWidth={1}
          />
        ))}
        
        {/* Data lines */}
        <path d={createPath('series1')} stroke="#8884d8" strokeWidth={2} fill="none" />
        <path d={createPath('series2')} stroke="#82ca9d" strokeWidth={2} fill="none" />
        <path d={createPath('series3')} stroke="#ffc658" strokeWidth={2} fill="none" />
        
        {/* Data points */}
        {data.map((d, i) => (
          <g key={i}>
            <circle cx={getX(i)} cy={getY(d.series1)} r={3} fill="#8884d8" />
            <circle cx={getX(i)} cy={getY(d.series2)} r={3} fill="#82ca9d" />
            <circle cx={getX(i)} cy={getY(d.series3)} r={3} fill="#ffc658" />
          </g>
        ))}
        
        {/* X-axis labels */}
        {data.map((d, i) => (
          <text
            key={i}
            x={getX(i)}
            y={height - 10}
            textAnchor="middle"
            fontSize="12"
            fill="#666"
          >
            {d.name}
          </text>
        ))}
      </svg>
      
      {/* Legend */}
      <Box sx={{ mt: 2, display: 'flex', gap: 2, justifyContent: 'center' }}>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
          <Box sx={{ width: 16, height: 3, bgcolor: '#8884d8' }} />
          <Typography variant="body2">Power Generation</Typography>
        </Box>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
          <Box sx={{ width: 16, height: 3, bgcolor: '#82ca9d' }} />
          <Typography variant="body2">Wind Speed</Typography>
        </Box>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
          <Box sx={{ width: 16, height: 3, bgcolor: '#ffc658' }} />
          <Typography variant="body2">Efficiency</Typography>
        </Box>
      </Box>
    </Paper>
  );
};

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
        Windpark A Dashboard
      </Typography>
      
      <Box sx={{ mt: 2 }}>
        <SimpleLineChart data={data} />
      </Box>
      
      {/* Additional dashboard metrics */}
      <Stack direction={{ xs: 'column', md: 'row' }} spacing={2} sx={{ mt: 3 }}>
        <Paper sx={{ p: 2, textAlign: 'center', flex: 1 }}>
          <Typography variant="h6" color="primary">Current Output</Typography>
          <Typography variant="h4">2.4 MW</Typography>
        </Paper>
        <Paper sx={{ p: 2, textAlign: 'center', flex: 1 }}>
          <Typography variant="h6" color="primary">Efficiency</Typography>
          <Typography variant="h4">87%</Typography>
        </Paper>
        <Paper sx={{ p: 2, textAlign: 'center', flex: 1 }}>
          <Typography variant="h6" color="primary">Availability</Typography>
          <Typography variant="h4">99.2%</Typography>
        </Paper>
      </Stack>
    </Container>
  );
};

export default DashboardsPage;