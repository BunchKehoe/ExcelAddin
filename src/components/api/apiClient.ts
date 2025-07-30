import axios from 'axios';

const API_BASE_URL = 'http://localhost:5000/api';

const apiClient = axios.create({
  baseURL: API_BASE_URL,
  timeout: 10000,
  headers: {
    'Content-Type': 'application/json',
  },
});

interface DownloadDataParams {
  fund: string;
  dataType: string;
  startDate: string;
  endDate: string;
}

interface RawDataDownloadParams {
  catalog: string;
  fund: string;
  start_date: string;
  end_date: string;
}

interface MarketDataDownloadParams {
  security: string;
  field: string;
  start_date: string;
  end_date: string;
}

// Raw Data Endpoints
export const getRawDataCategories = async () => {
  try {
    const response = await apiClient.get('/raw-data/categories');
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    throw error;
  }
};

export const getRawDataFunds = async (catalog: string) => {
  try {
    const response = await apiClient.get(`/raw-data/funds/${catalog}`);
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    throw error;
  }
};

export const downloadRawData = async (params: RawDataDownloadParams) => {
  try {
    const response = await apiClient.post('/raw-data/download', params);
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    throw error;
  }
};

// Market Data Endpoints
export const getMarketDataSecurities = async () => {
  try {
    const response = await apiClient.get('/market-data/securities');
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    throw error;
  }
};

export const getMarketDataFields = async (security: string) => {
  try {
    const response = await apiClient.get(`/market-data/fields/${security}`);
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    throw error;
  }
};

export const downloadMarketData = async (params: MarketDataDownloadParams) => {
  try {
    const response = await apiClient.post('/market-data/download', params);
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    throw error;
  }
};

// Legacy endpoints for backward compatibility
export const downloadData = async (params: DownloadDataParams) => {
  try {
    const response = await apiClient.post('/download-data', params);
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    // Return mock data for demo purposes
    return {
      data: [
        { date: '2024-01-01', value: 100.00, change: 0.00, percentage: '0.00%' },
        { date: '2024-01-02', value: 101.50, change: 1.50, percentage: '1.50%' },
        { date: '2024-01-03', value: 99.75, change: -1.75, percentage: '-1.72%' },
        { date: '2024-01-04', value: 102.25, change: 2.50, percentage: '2.51%' },
        { date: '2024-01-05', value: 98.90, change: -3.35, percentage: '-3.28%' }
      ]
    };
  }
};

export const getFunds = async () => {
  try {
    const response = await apiClient.get('/funds');
    return response.data;
  } catch (error) {
    console.error('API Error:', error);
    // Return mock data for demo purposes
    return [
      'Global Equity Fund',
      'Fixed Income Fund',
      'Emerging Markets Fund',
      'Technology Fund',
      'Real Estate Fund'
    ];
  }
};

export default apiClient;