# Excel Add-in Backend API

This is a Python Flask backend application that provides data access APIs for the Excel add-in frontend. The backend follows Domain Driven Architecture (DDA) principles and provides endpoints for raw database tables and market data.

## Architecture

The backend is structured using Domain Driven Architecture:

```
backend/
├── src/
│   ├── domain/           # Business logic and entities
│   │   ├── entities/     # Domain entities
│   │   ├── repositories/ # Repository interfaces
│   │   └── services/     # Domain services
│   ├── infrastructure/   # External concerns
│   │   ├── config/       # Configuration management
│   │   └── database/     # Database implementations
│   ├── application/      # Application services
│   │   ├── services/     # Application services
│   │   └── dtos/         # Data Transfer Objects
│   └── presentation/     # Controllers and API endpoints
│       ├── controllers/  # Flask controllers
│       └── middleware/   # Middleware components
├── app.py               # Main Flask application
├── run.py              # Application runner
├── requirements.txt    # Python dependencies
└── database.cfg       # Database configuration
```

## Requirements

- Python 3.12
- Flask 3.0.2
- SQLAlchemy 2.0.27
- pyodbc (for SQL Server connectivity)
- Flask-CORS (for frontend integration)

## Setup

1. **Install dependencies:**
   ```bash
   cd backend
   pip install -r requirements.txt
   ```

2. **Configure database connection:**
   Edit `database.cfg` with your SQL Server connection details:
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