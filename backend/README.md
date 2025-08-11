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
├── requirements.txt       # Python dependencies
└── database.cfg          # Database configuration
```

## Quick Start

### Local Development
```bash
cd backend
pip install -r requirements.txt
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

#### **Recommended: Environment Variables**
Create environment-specific `.env` files:

**Local Development** (`.env` or `.env.development`):
```env
ENVIRONMENT=development
DEBUG=true
# No database configuration = uses mock data
# Optional: LOCAL_DATABASE_URL=sqlite:///./local_dev.db
```

**Staging** (`.env.staging`):
```env
ENVIRONMENT=staging
STAGING_DATABASE_URL=mssql+pyodbc://user:pass@server-vs81t.intranet.local/test?driver=ODBC+Driver+17+for+SQL+Server
```

**Production** (`.env.production`):
```env
ENVIRONMENT=production
PRODUCTION_DATABASE_URL=mssql+pyodbc://user:pass@server-vs84.intranet.local/test?driver=ODBC+Driver+17+for+SQL+Server
```

#### **Legacy: Configuration Files**
The `database.cfg` file is maintained for backwards compatibility but is superseded by environment variables.

### Development Features

- **Mock Data Fallback**: Local development automatically uses mock data when no database is configured
- **Flexible Database**: Can optionally connect to local/test databases via `LOCAL_DATABASE_URL`
- **Environment Isolation**: No hardcoded staging/production URLs that break local development

### Application Settings

Additional settings can be configured via `.env` files:
- `DEBUG` - Enable debug mode
- `HOST` / `PORT` - Server binding
- `CORS_ORIGINS` - Allowed CORS origins  
- `NIFI_ENDPOINT` - NiFi server URL for data uploads

## API Endpoints

- `GET /api/health` - Health check endpoint
- Additional endpoints defined in `src/presentation/controllers/`

For complete setup, deployment, and troubleshooting information, see the main project documentation:
- [Application Guide](../APPLICATION_GUIDE.md)
- [Deployment Guide](../DEPLOYMENT_GUIDE.md)  
- [Troubleshooting Guide](../TROUBLESHOOTING_GUIDE.md)
   ```ini
   [database]
   url = mssql+pyodbc://user:pass@server-vs81t.intranet.local/test?driver=ODBC+Driver+17+for+SQL+Server
   ```

3. **Run the application:**
   ```bash
   python run.py
   ```

   The API will be available at `http://localhost:5000`

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