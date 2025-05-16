import sys
import os

# Add the root directory to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from src.main.main import main

if __name__ == "__main__":
    main()
