"""
Simple test to validate the Flask application structure.
"""
import sys
import os

# Add the backend directory to Python path
backend_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, backend_dir)

def test_app_creation():
    """Test if the Flask app can be created."""
    try:
        from app import create_app
        app = create_app()
        print("✓ Flask app created successfully")
        return True
    except Exception as e:
        print(f"✗ Failed to create Flask app: {e}")
        return False

def test_imports():
    """Test if all modules can be imported."""
    try:
        from src.infrastructure.config.app_config import AppConfig, DatabaseConfig
        print("✓ Config imports successful")
        
        from src.domain.entities.raw_data import FileCategory, Fund, RawDataRecord
        from src.domain.entities.market_data import Security, DataField, MarketDataRecord
        print("✓ Entity imports successful")
        
        from src.application.services.raw_data_service import RawDataService
        from src.application.services.market_data_service import MarketDataService
        print("✓ Service imports successful")
        
        return True
    except Exception as e:
        print(f"✗ Import failed: {e}")
        return False

def test_config():
    """Test configuration loading."""
    try:
        from src.infrastructure.config.app_config import DatabaseConfig
        config = DatabaseConfig()
        url = config.database_url
        print(f"✓ Database config loaded: {url[:50]}...")
        return True
    except Exception as e:
        print(f"✗ Config test failed: {e}")
        return False

if __name__ == '__main__':
    print("Running backend validation tests...")
    print("-" * 50)
    
    tests = [
        test_imports,
        test_config,
        test_app_creation
    ]
    
    passed = 0
    for test in tests:
        if test():
            passed += 1
        else:
            break
    
    print("-" * 50)
    print(f"Tests passed: {passed}/{len(tests)}")
    
    if passed == len(tests):
        print("All tests passed! Backend structure is valid.")
        sys.exit(0)
    else:
        print("Some tests failed. Please check the issues above.")
        sys.exit(1)