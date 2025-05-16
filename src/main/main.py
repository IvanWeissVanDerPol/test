from src.utils.logging import setup_logging
from src.pipeline.pipeline import Pipeline

logger = setup_logging()

def main():
    """Main function to execute the complete process."""
    pipeline = Pipeline()
    pipeline.run()

if __name__ == "__main__":
    main()
