# Cryptocurrency Live Data Fetcher

## Overview
This Python script fetches real-time cryptocurrency data from the CoinGecko API and automatically updates an Excel spreadsheet with the top 50 cryptocurrencies by market capitalization.

## Features
- Fetches top 50 cryptocurrencies by market cap
- Real-time data updates 
- Automatic Excel spreadsheet population
- Logging for tracking script activities
- Error handling and retry mechanisms

## Prerequisites

### System Requirements
- Windows Operating System (due to Windows-specific Excel COM library)
- Python 3.7+
- Microsoft Excel installed

### Dependencies
- requests
- pandas
- pywin32
- pythoncom

## Installation Steps

### 1. Clone the Repository
```bash
git clone https://github.com/yourusername/cryptocurrency-live-data-fetcher.git
cd cryptocurrency-live-data-fetcher
```

### 2. Create a Virtual Environment (Recommended)
```bash
python -m venv venv
venv\Scripts\activate  # On Windows
```

### 3. Install Required Dependencies
```bash
pip install -r requirements.txt
```

### 4. Create requirements.txt
Create a `requirements.txt` file with the following contents:
```
requests
pandas
pywin32
pythoncom
```

## Configuration

### Excel File Path
- By default, the script creates `cryptocurrency_live_data.xlsx` in the script's directory
- You can specify a custom path in the `main()` function

## Running the Script
```bash
python cryptocurrency_fetcher.py
```

## Customization
- Modify `time.sleep()` duration to control update frequency
- Adjust API parameters in `fetch_top_50_cryptocurrencies()` method

## Logging
- Logs are printed to console
- Tracks successful data fetches, updates, and potential errors

## Notes
- Requires active internet connection
- Uses free CoinGecko API with rate limits
- Designed for Windows with Excel integration

## Troubleshooting
- Ensure all dependencies are installed
- Check that Excel is not blocked by antivirus
- Verify Python and Excel are 64-bit or both 32-bit versions

# Demo
Check out the price of bitcoin in this demo video to see the results.
https://github.com/user-attachments/assets/f266af1e-0b33-43ec-a41e-083d1a8f256b

