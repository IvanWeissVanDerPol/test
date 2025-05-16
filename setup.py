from setuptools import setup, find_packages

setup(
    name="excel_processor",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pandas>=1.0.0",
        "openpyxl>=3.0.0"
    ],
    python_requires='>=3.7',
    entry_points={
        'console_scripts': [
            'excel_processor=src.main.main:main'
        ]
    }
)
