import { createTheme } from '@mui/material/styles';

// Colors taken from logo - please correct if not accurate
export const colors = {
  primeBlue: '#005F9C',
  primeGray: '#686A69',
  primeLux: '#009FE3',
  white: '#FFFFFF',
  black: '#000000',
};

const theme = createTheme({
  palette: {
    primary: {
      main: colors.primeBlue,
      contrastText: colors.white,
    },
    secondary: {
      main: colors.primeLux,
      contrastText: colors.white,
    },
    background: {
      default: colors.white,
    },
    text: {
      primary: colors.black,
      secondary: colors.primeGray,
    },
  },
  typography: {
    fontFamily: '"Avenir LT 55 Roman", "Segoe UI", "Roboto", "Arial", sans-serif',
  },
});

export default theme;
