#!/usr/bin/env python3
"""
Simple runner script for the Flask backend.
"""
import os
import sys

# Add the backend directory to Python path
backend_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, backend_dir)

from app import main

if __name__ == '__main__':
    main()