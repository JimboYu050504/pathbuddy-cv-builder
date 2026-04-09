import sys
import os

# Make the project root importable so `from app import app` works
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app

# Vercel looks for a module-level name called `app` (WSGI callable)
# Nothing else needed — Vercel's Python runtime handles the rest.
