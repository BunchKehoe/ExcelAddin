# Excel Add-in Backend API

Python Flask backend service providing REST API endpoints for the Excel add-in frontend.

## Architecture

Domain Driven Architecture structure:
```
backend/
├── src/                   # Business logic and API layers
├── app.py                 # Main Flask application
├── run.py                 # Development server entry point
├── service_wrapper.py     # Windows service wrapper
├── pyproject.toml         # Python dependencies and Poetry configuration
└── database.cfg          # Database configuration
```

## Quick Start

### Local Development
```bash
cd backend
poetry install
poetry shell
python run.py
```
API available at `http://localhost:5000`

### Windows Service (Production)
```powershell
# Use the deployment script
.\deployment\scripts\setup-backend-service.ps1
```

## Configuration

The backend uses **environment-aware configuration** that automatically adapts to different deployment scenarios.

### Environment Detection

The application automatically detects its environment based on:
- `ENVIRONMENT` environment variable (`development`, `staging`, `production`)
- Hostname patterns (development machines, vs81t for staging, vs84 for production)
- Other runtime indicators

### Database Configuration

The backend uses the `database.cfg` file with environment-specific sections for clean and simple configuration:

```ini
# Database Configuration for Excel Add-in Backend
# Environment-specific database configurations
# The application automatically detects the environment and uses the appropriate section

[development]
# For local development - leave empty to use mock data
# Uncomment and configure if you want to use a real database for local development
# url = sqlite:///./development.db

[staging] 
url = mssql+pyodbc://user:pass@server-vs81t.intranet.local/test?driver=ODBC+Driver+17+for+SQL+Server

[production]
url = mssql+pyodbc://user:pass@server-vs84.intranet.local/test?driver=ODBC+Driver+17+for+SQL+Server
```

#### Environment Behavior:
- **Development**: Uses mock data by default (no database connection required)
- **Staging**: Uses staging SQL Server database  
- **Production**: Uses production SQL Server database

#### Local Database Override:
To use a local database for development, uncomment and configure the URL in the `[development]` section of `database.cfg`.

### Development Features

- **Mock Data Fallback**: Local development automatically uses mock data when no database is configured
- **Clean Environment Detection**: Automatically detects environment and uses appropriate configuration section
- **No Hardcoded URLs**: All database URLs are in the configuration file

### Application Settings

Additional settings can be configured via `.env` files:
- `DEBUG` - Enable debug mode
- `HOST` / `PORT` - Server binding
- `CORS_ORIGINS` - Allowed CORS origins  
- `NIFI_ENDPOINT` - NiFi server URL for data uploads

## API Endpoints

### Health Check
- `GET /api/health` - Health check endpoint

### Raw Database Tables
- `GET /api/raw-data/categories` - Get file categories for dropdown
- `GET /api/raw-data/funds/{catalog}` - Get funds for a specific catalog
- `POST /api/raw-data/download` - Download raw data

### Market Data
- `GET /api/market-data/securities` - Get securities for dropdown
- `GET /api/market-data/fields/{security}` - Get fields for a specific security
- `POST /api/market-data/download` - Download market data

## Database Queries

The backend executes the following SQL queries:

### Raw Data
1. **Categories:** `SELECT FILE_CATEGORY FROM test.dbo.DELIVERY_CATALOG d WHERE GETDATE() > d.VALID_FROM AND GETDATE() < d.VALID_TO`
2. **Funds:** `SELECT DISTINCT FUND FROM test.dbo.<catalog>`
3. **Data:** `SELECT * FROM test.dbo.<catalog> c WHERE c.fund = :fund AND c.START_DATE BETWEEN :start AND :end`

### Market Data
1. **Securities:** `SELECT DISTINCT b.security FROM BLOOMBERG_ODD_MONTHLY b`
2. **Fields:** `SELECT DISTINCT b.field FROM BLOOMBERG_ODD_MONTHLY b WHERE b.security = :security`
3. **Data:** `SELECT * FROM BLOOMBERG_ODD_MONTHLY b WHERE b.security = :security AND b.field = :field AND b.date BETWEEN :start AND :end`

## Development

### Testing the structure:
```bash
python test_structure.py
```

### Running in debug mode:
Set the `DEBUG` environment variable:
```bash
export DEBUG=true
python run.py
```

### CORS Configuration:
The backend is configured to allow requests from `http://localhost:3000` (the frontend development server). To add more origins, set the `CORS_ORIGINS` environment variable:
```bash
export CORS_ORIGINS="http://localhost:3000,https://your-domain.com"
```

## Data Processing

The backend includes optimizations for large datasets:
- Batching support for large data downloads
- Streaming capabilities for real-time data processing
- Proper error handling and logging
- Data transformation optimized for Excel integration

## Frontend Integration

The frontend TypeScript code should use the API endpoints as follows:

```typescript
// Raw data categories
const response = await fetch('http://localhost:5000/api/raw-data/categories');

// Download raw data
const response = await fetch('http://localhost:5000/api/raw-data/download', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    catalog: 'CATEGORY_NAME',
    fund: 'FUND_NAME',
    start_date: '2024-01-01',
    end_date: '2024-12-31'
  })
});
```