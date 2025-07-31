import React from 'react';
import {
  Container,
  Typography,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  Stack
} from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';

const ExcelFunctionsPage: React.FC = () => {
  const functions = [
    {
      title: 'Download Market Data',
      description: 'Input-Bloomberg Ticker-DateFrom-DateTo Returns-Daily Close Prices',
      details: 'This function allows you to download historical market data from Bloomberg. Simply provide the ticker symbol, start date, and end date to retrieve daily closing prices for the specified period.'
    },
    {
      title: 'Download Static Data',
      description: 'Input-Bloomberg Ticker-Date',
      details: 'Download static reference data for a specific security as of a particular date. This includes fundamental data points such as company information, sector classification, and other static attributes.'
    },
    {
      title: 'Aggregate IRR',
      description: '=PC.AGGIRR(expected_future_value, original_beginning_value)',
      details: 'Calculate the Internal Rate of Return (IRR) for a portfolio or aggregate of investments. This function takes the expected future value and original beginning value, then divides the former by the latter to compute the aggregate return rate.'
    },
    {
      title: 'Weighted IRR',
      description: 'This Function calculates Weighted IRR',
      details: 'Calculate the Weighted Internal Rate of Return for multiple investments or portfolios. This function considers the relative size or importance of each investment when computing the overall return rate.'
    },
    {
      title: 'Join Cells',
      description: '=PC.JOINCELLS(range, [delimiter])',
      details: 'Transform a range of cells into a single delimited string. Takes a range like A1:A23 and an optional delimiter (default is comma), then combines all cell values into a single string separated by the delimiter, e.g., "A1, A2, A3, ..."'
    }
  ];

  return (
    <Container maxWidth="md" sx={{ mt: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom align="center">
        Excel Functions
      </Typography>
      
      <Stack spacing={2} sx={{ mt: 4 }}>
        {functions.map((func, index) => (
          <Accordion key={index}>
            <AccordionSummary
              expandIcon={<ExpandMoreIcon />}
              aria-controls={`panel${index}-content`}
              id={`panel${index}-header`}
            >
              <Typography variant="h6">
                {func.title}: {func.description}
              </Typography>
            </AccordionSummary>
            <AccordionDetails>
              <Typography>
                {func.details}
              </Typography>
            </AccordionDetails>
          </Accordion>
        ))}
      </Stack>
    </Container>
  );
};

export default ExcelFunctionsPage;