import requests
import pandas as pd
import time
from datetime import datetime
import os
import logging
import pythoncom
import win32com.client
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


class CryptocurrencyDataFetcher:
    def __init__(self, excel_file_path=None):
        """
        Initialize the cryptocurrency data fetcher.

        Args:
            excel_file_path (str, optional): Path to the Excel file to update.
        """
        self.base_url = "https://api.coingecko.com/api/v3/coins/markets"  
        # using coingecko api cause its free
        # If no file path is provided we will create a default file in the current directory
        if excel_file_path is None:
            excel_file_path = "cryptocurrency_live_data.xlsx"

        self.excel_file_path = os.path.abspath(excel_file_path)
        self.app = None
        self.wb = None
        self.sheet = None

    def fetch_top_50_cryptocurrencies(self):
        """
        Fetch top 50 cryptocurrencies by market capitalization.

        Returns:
            pandas.DataFrame: DataFrame with cryptocurrency data.
        """
        params = {
            "vs_currency": "usd",
            "order": "market_cap_desc",
            "per_page": 50,
            "page": 1,
            "sparkline": False,
        }

        try:
            response = requests.get(self.base_url, params=params)
            response.raise_for_status()
            data = response.json()

            logger.info(f"Successfully fetched {len(data)} cryptocurrencies")

            #This is to extract required fields
            crypto_data = []
            for coin in data:
                crypto_data.append(
                    {
                        "Name": coin["name"],
                        "Symbol": coin["symbol"].upper(),
                        "Current Price (USD)": coin["current_price"],
                        "Market Capitalization": coin["market_cap"],
                        "24h Trading Volume": coin["total_volume"],
                        "24h Price Change (%)": coin["price_change_percentage_24h"],
                    }
                )

            df = pd.DataFrame(crypto_data)
            logger.info(f"DataFrame created with {len(df)} rows")
            return df

        except requests.RequestException as e:
            logger.error(f"Error fetching cryptocurrency data: {e}")
            return pd.DataFrame()

    def initialize_excel_connection(self):
        """
        Initialize connection to Excel workbook, creating if not exists.
        """
        try:
            # Ensure COM is initialized for the current thread
            pythoncom.CoInitialize()

            # Create Excel application
            self.app = win32com.client.Dispatch("Excel.Application")
            self.app.Visible = True
            self.app.DisplayAlerts = False

            # Open or create workbook
            if os.path.exists(self.excel_file_path):
                self.wb = self.app.Workbooks.Open(self.excel_file_path)
            else:
                self.wb = self.app.Workbooks.Add()
                self.wb.SaveAs(self.excel_file_path)

            # Select or create sheet
            try:
                self.sheet = self.wb.Sheets("CryptocurrencyData")
            except:
                self.sheet = self.wb.Sheets.Add()
                self.sheet.Name = "CryptocurrencyData"

            logger.info("Excel connection initialized successfully")

        except Exception as e:
            logger.error(f"Error initializing Excel connection: {e}")
            traceback.print_exc()
            raise

    def update_excel_sheet(self, df):
        """
        Update Excel sheet with live cryptocurrency data.

        Args:
            df (pandas.DataFrame): DataFrame with cryptocurrency data.
        """
        try:
            # Ensure DataFrame is not empty
            if df.empty:
                logger.warning("DataFrame is empty. Skipping Excel update.")
                return

            # Ensure Excel connection is established
            if self.app is None or self.wb is None or self.sheet is None:
                self.initialize_excel_connection()

            # Save DataFrame to CSV as a backup/debugging method
            csv_path = os.path.splitext(self.excel_file_path)[0] + ".csv"
            df.to_csv(csv_path, index=False)
            logger.info(f"Data also saved to CSV: {csv_path}")

            # Clear existing content in the sheet
            self.sheet.Cells.Clear()

            # Write headers
            for col_idx, col_name in enumerate(df.columns, start=1):
                self.sheet.Cells(1, col_idx).Value = col_name

            # Write data
            for row_idx, row in df.iterrows():
                for col_idx, value in enumerate(row, start=1):
                    self.sheet.Cells(row_idx + 2, col_idx).Value = value

            # Save workbook
            self.wb.Save()
            logger.info("Excel sheet updated successfully!")

        except Exception as e:
            logger.error(f"Error updating Excel sheet: {e}")
            traceback.print_exc()

    def close_excel_connection(self):
        """
        Close Excel workbook and application.
        """
        try:
            if hasattr(self, "wb") and self.wb:
                self.wb.Save()
                self.wb.Close()

            if hasattr(self, "app") and self.app:
                self.app.Quit()

        except Exception as e:
            logger.error(f"Error closing Excel connection: {e}")

        finally:
            # Reset all references
            self.app = None
            self.wb = None
            self.sheet = None

            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass


def main():
    fetcher = CryptocurrencyDataFetcher(
        excel_file_path=r"C:\Users\DELL\Downloads\work\cryptocurrency_live_data.xlsx"
    )

    try:
        # Initialize Excel connection at the start
        fetcher.initialize_excel_connection()

        while True:
            try:
                # Fetch live data
                df = fetcher.fetch_top_50_cryptocurrencies()

                # Update Excel sheet
                fetcher.update_excel_sheet(df)

                # Wait for 30 sec before next update
                logger.info("Waiting 5 minutes before next update...")
                time.sleep(30)  # 30sec minutes

            except Exception as e:
                logger.error(f"Error in main loop: {e}")
                traceback.print_exc()
                time.sleep(60)  # Wait a bit before retrying

    except Exception as e:
        logger.error(f"Initialization error: {e}")
        traceback.print_exc()

    finally:
        # Ensure Excel connection is closed
        fetcher.close_excel_connection()


if __name__ == "__main__":
    main()
