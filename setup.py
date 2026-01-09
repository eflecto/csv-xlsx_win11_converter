"""
Setup script for CSV to XLSX Converter
For creating distributable packages
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="csv-to-xlsx-converter",
    version="1.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="A beautiful Windows 11 utility for converting CSV files to Excel format",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/csv-to-xlsx-converter",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business :: Office Suites",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows :: Windows 11",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    python_requires=">=3.9",
    install_requires=[
        "customtkinter>=5.2.0",
        "pandas>=2.0.0",
        "openpyxl>=3.1.0",
    ],
    entry_points={
        "console_scripts": [
            "csv2xlsx=main:main",
        ],
        "gui_scripts": [
            "csv2xlsx-gui=main:main",
        ],
    },
    keywords="csv, xlsx, excel, converter, windows, gui",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/csv-to-xlsx-converter/issues",
        "Source": "https://github.com/yourusername/csv-to-xlsx-converter",
    },
)
