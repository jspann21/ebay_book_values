"""
eBay Book Price Searcher

This script provides a graphical user interface (GUI) application that allows users to search for
book prices on eBay. The application leverages Selenium for web scraping and PyQt5 for the GUI.

Key Features:
- Load a list of books (title, author, ISBN) from an Excel file.
- Perform automated searches on eBay to retrieve book pricing information.
- Display the search results in a table with options to export the results to an Excel file.
- Includes logging functionality to provide feedback during the search process.

Modules:
- Worker: A QThread subclass that handles the background search process.
- EbaySearcher: The main QWidget subclass that manages the GUI and coordinates the search process.

Dependencies:
- PyQt5: For the GUI components.
- Selenium: For web scraping eBay.
- pandas: For handling Excel files and data management.

Usage:
- Run the script to launch the application.
- Use the "Browse" button to select an Excel file with the columns 'Title', 'Author', and 'ISBN'.
- Click "Search" to start the search process, or "Cancel" to stop an ongoing search.
- Results can be exported to an Excel file using the "Export Results" button.
"""

import sys
import time
import random
import os
import json
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QTableWidget, QTableWidgetItem,
                             QLabel, QPlainTextEdit, QHeaderView, QFileDialog, QLineEdit)
from PyQt5.QtCore import QThread, pyqtSignal, QCoreApplication
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException


class Worker(QThread):
    """
    Worker thread to perform searches on eBay using Selenium.

    Attributes:
        update_result (pyqtSignal): Signal to update the result table with search results.
        log_message (pyqtSignal): Signal to log messages to the log window.
        search_complete (pyqtSignal): Signal emitted when the search is complete.
    """

    update_result = pyqtSignal(list)
    log_message = pyqtSignal(str)
    search_complete = pyqtSignal()

    def __init__(self, driver, data):
        """
        Initializes the worker thread.

        Args:
            driver (webdriver.Chrome): The Selenium WebDriver instance.
            data (list): List of search data entries (title, author, ISBN).
        """
        super().__init__()
        self.driver = driver
        self.data = data
        self._running = True

    def run(self):
        """
        Executes the search process in the background thread.
        """
        for entry in self.data:
            if not self._running:
                self.log_message.emit("Search cancelled.")
                break

            title, author, isbn = entry

            if not title and not author and not isbn:
                self.log_message.emit("Skipping empty line.")
                continue

            self.log_message.emit(f"Processing: {title}, {author}, ISBN: {isbn}")

            search_query = isbn if isbn else f"{title} {author}"
            if not self.perform_search(search_query, title, author, isbn):
                search_query = f"{title} {author}"
                self.perform_search(search_query, title, author, isbn)

            delay = random.uniform(15, 40)
            self.log_message.emit(f"Sleeping for {delay:.2f} seconds to avoid being flagged.")
            time.sleep(delay)

            QCoreApplication.processEvents()

        self.search_complete.emit()

    def stop(self):
        """Stops the search process."""
        self._running = False

    def perform_search(self, search_query, title, author, isbn):
        """
        Performs a search on eBay using the specified search query.

        Args:
            search_query (str): The search query string.
            title (str): The book title.
            author (str): The book author.
            isbn (str): The book ISBN.

        Returns:
            bool: True if results were found, False otherwise.
        """
        url = (f"https://www.ebay.com/sh/research?marketplace=EBAY-US"
               f"&keywords={search_query}&dayRange=365&endDate=1724798549762"
               f"&startDate=1693262549762&categoryId=0&offset=0&limit=50"
               f"&tabName=SOLD&tz=America%2FNew_York")
        self.driver.get(url)

        try:
            self.log_message.emit(f"Navigating to URL: {url}")

            WebDriverWait(self.driver, 19).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.aggregates"))
            )
            self.log_message.emit("Aggregates section found.")

            avg_price = self.get_aggregate_value(
                "//section[contains(., 'Avg sold price')]/div[@class='metric-value']")
            price_range = self.get_aggregate_value(
                "//section[contains(., 'Sold price range')]/div[@class='metric-value']")
            avg_shipping = self.get_aggregate_value(
                "//section[contains(., 'Avg shipping')]/div[@class='metric-value']")
            total_sellers = self.get_aggregate_value(
                "//section[contains(., 'Total sellers')]/div[@class='metric-value']")

            self.log_message.emit(
                f"Data extracted: Avg Price: {avg_price}, Price Range: {price_range}, "
                f"Avg Shipping: {avg_shipping}, Total Sellers: {total_sellers}")

            if not any([avg_price, price_range, avg_shipping, total_sellers]):
                self.log_message.emit(f"No valid data found for {search_query}. Skipping.")
                return False

            self.log_message.emit(
                f"Found results for {title} by {author} (ISBN: {isbn})")

            self.update_result.emit([title, author, isbn, avg_price, price_range,
                                     avg_shipping, total_sellers])
            return True

        except TimeoutException:
            self.log_message.emit(f"No results found for {search_query}. Skipping.")
            return False

    def get_aggregate_value(self, xpath):
        """
        Retrieves the value of an aggregate field from the eBay page.

        Args:
            xpath (str): The XPath of the element to retrieve.

        Returns:
            str: The text value of the element, or an empty string if not found.
        """
        try:
            element = WebDriverWait(self.driver, 1).until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            )
            return element.text.strip()
        except TimeoutException:
            delay = random.uniform(3, 9)
            self.log_message.emit(
                f"Sleeping for {delay:.2f} seconds to avoid being flagged.")
            time.sleep(delay)
            self.log_message.emit(
                f"Could not find element with xpath: {xpath}. Returning empty value.")
            return ""


class EbaySearcher(QWidget):
    """
    Main application window for the eBay Book Price Searcher.

    Attributes:
        driver (webdriver.Chrome): The Selenium WebDriver instance.
        worker (Worker): The worker thread for performing searches.
        data (list): List of search data entries.
    """

    def __init__(self):
        """Initializes the eBaySearcher application."""
        super().__init__()

        self.settings_file = "settings.json"
        self.settings = self.load_settings()

        self.initUI()
        self.setup_selenium()
        self.worker = None
        self.data = []

    def initUI(self):
        """Initializes the user interface components."""
        layout = QVBoxLayout()

        # User Data Directory Input
        user_data_dir_layout = QVBoxLayout()
        user_data_dir_label = QLabel("Location of Chrome user-data-dir for Selenium", self)
        user_data_dir_layout.addWidget(user_data_dir_label)

        user_data_dir_input_layout = QHBoxLayout()
        self.user_data_dir_input = QLineEdit(self)
        self.user_data_dir_input.setText(self.settings.get("user_data_dir", ""))
        user_data_dir_input_layout.addWidget(self.user_data_dir_input)

        user_data_dir_button = QPushButton("Browse", self)
        user_data_dir_button.clicked.connect(self.browse_user_data_dir)
        user_data_dir_input_layout.addWidget(user_data_dir_button)
        user_data_dir_layout.addLayout(user_data_dir_input_layout)
        layout.addLayout(user_data_dir_layout)

        # Profile Directory Input
        profile_dir_layout = QVBoxLayout()
        profile_dir_label = QLabel("Enter name of Chrome profile-directory for Selenium", self)
        profile_dir_layout.addWidget(profile_dir_label)

        profile_dir_input_layout = QHBoxLayout()
        self.profile_dir_input = QLineEdit(self)
        self.profile_dir_input.setText(self.settings.get("profile_directory", ""))
        profile_dir_input_layout.addWidget(self.profile_dir_input)

        profile_dir_button = QPushButton("Browse", self)  # Optional, add functionality if needed
        profile_dir_input_layout.addWidget(profile_dir_button)
        profile_dir_layout.addLayout(profile_dir_input_layout)
        layout.addLayout(profile_dir_layout)

        # Chromedriver.exe Input
        chromedriver_layout = QVBoxLayout()
        chromedriver_label = QLabel("Location of chromedriver.exe for Selenium", self)
        chromedriver_layout.addWidget(chromedriver_label)

        chromedriver_input_layout = QHBoxLayout()
        self.chromedriver_input = QLineEdit(self)
        self.chromedriver_input.setText(self.settings.get("chromedriver_path", ""))
        chromedriver_input_layout.addWidget(self.chromedriver_input)

        chromedriver_button = QPushButton("Browse", self)
        chromedriver_button.clicked.connect(self.browse_chromedriver)
        chromedriver_input_layout.addWidget(chromedriver_button)
        chromedriver_layout.addLayout(chromedriver_input_layout)
        layout.addLayout(chromedriver_layout)

        # Excel File Input
        excel_file_layout = QVBoxLayout()
        excel_file_label = QLabel("Location of .xlsx file", self)
        excel_file_layout.addWidget(excel_file_label)

        excel_file_input_layout = QHBoxLayout()
        self.excel_file_input = QLineEdit(self)
        self.excel_file_input.setText(self.settings.get("excel_file", ""))
        excel_file_input_layout.addWidget(self.excel_file_input)

        excel_file_button = QPushButton("Browse", self)
        excel_file_button.clicked.connect(self.browse_file)
        excel_file_input_layout.addWidget(excel_file_button)
        excel_file_layout.addLayout(excel_file_input_layout)
        layout.addLayout(excel_file_layout)

        # Control Buttons
        btn_layout = QHBoxLayout()
        clear_button = QPushButton("Clear")
        clear_button.clicked.connect(self.clear_input)
        btn_layout.addWidget(clear_button)

        search_button = QPushButton("Search")
        search_button.clicked.connect(self.search_ebay)
        btn_layout.addWidget(search_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.cancel_search)
        btn_layout.addWidget(cancel_button)

        layout.addLayout(btn_layout)

        # Results Table
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(7)
        self.result_table.setHorizontalHeaderLabels(
            ["Title", "Author", "ISBN", "Avg Sold Price",
             "Sold Price Range", "Avg Shipping", "Total Sellers"])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.result_table)

        # Export Results Button
        export_button = QPushButton("Export Results")
        export_button.clicked.connect(self.export_results)
        layout.addWidget(export_button)

        # Log Window
        self.log_window = QPlainTextEdit()
        self.log_window.setReadOnly(True)
        layout.addWidget(QLabel("Log:"))
        layout.addWidget(self.log_window)

        self.setLayout(layout)
        self.setWindowTitle("eBay Book Price Searcher")
        self.setGeometry(100, 100, 800, 600)

    def load_settings(self):
        """Loads saved settings from a JSON file."""
        if os.path.exists(self.settings_file):
            with open(self.settings_file, 'r', encoding='utf-8') as file:
                return json.load(file)
        return {}

    def save_settings(self):
        """Saves current settings to a JSON file."""
        self.settings['user_data_dir'] = self.user_data_dir_input.text()
        self.settings['profile_directory'] = self.profile_dir_input.text()
        self.settings['chromedriver_path'] = self.chromedriver_input.text()
        self.settings['excel_file'] = self.excel_file_input.text()

        with open(self.settings_file, 'w', encoding='utf-8') as file:
            json.dump(self.settings, file)

    def browse_user_data_dir(self):
        """Opens a file dialog to select the Chrome user-data-dir."""
        user_data_dir = QFileDialog.getExistingDirectory(self, "Select Chrome User Data Directory")
        if user_data_dir:
            self.user_data_dir_input.setText(user_data_dir)
            self.save_settings()

    def browse_chromedriver(self):
        """Opens a file dialog to select the chromedriver.exe."""
        chromedriver_path, _ = QFileDialog.getOpenFileName(
            self, "Select Chromedriver", "",
            "Executable Files (*.exe);;All Files (*)")
        if chromedriver_path:
            self.chromedriver_input.setText(chromedriver_path)
            self.save_settings()

    def browse_file(self):
        """Opens a file dialog to select an Excel file."""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "",
            "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_path:
            self.excel_file_input.setText(file_path)
            self.log(f"Selected file: {file_path}")
            self.save_settings()
            self.process_file(file_path)

    def setup_selenium(self):
        """Sets up Selenium WebDriver with Chrome."""
        user_data_dir = self.user_data_dir_input.text()
        profile_dir = self.profile_dir_input.text()
        chromedriver_path = self.chromedriver_input.text()

        if not user_data_dir or not profile_dir or not chromedriver_path:
            self.log("Error: Please ensure all paths (user-data-dir, profile-directory, chromedriver) are provided.")
            return False

        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument(f"user-data-dir={user_data_dir}")  # Chrome user data dir
        options.add_argument(f"profile-directory={profile_dir}")  # Chrome profile directory

        self.driver = webdriver.Chrome(
            service=Service(chromedriver_path),  # Chromedriver path
            options=options
        )
        return True

    def process_file(self, file_path):
        """
        Processes the selected Excel file and loads data into the application.

        Args:
            file_path (str): Path to the Excel file.
        """
        try:
            df = pd.read_excel(file_path)

            if 'Title' not in df.columns or 'Author' not in df.columns or 'ISBN' not in df.columns:
                self.log("Excel file must contain 'Title', 'Author', and 'ISBN' columns.")
                return

            data = df[['Title', 'Author', 'ISBN']].fillna('')  # Replace NaNs with empty strings
            self.data = data.values.tolist()
            self.log(f"Loaded {len(self.data)} rows from the file.")

        except FileNotFoundError:
            self.log(f"File not found: {file_path}")

        except pd.errors.EmptyDataError:
            self.log("The file is empty. Please select a non-empty Excel file.")

        except ValueError:
            self.log("Invalid file format. Please select a valid Excel file.")

        except (KeyError, AttributeError) as e:
            # Handle specific exceptions related to DataFrame operations
            self.log(f"Data processing error: {str(e)}")

        except Exception as e:
            # Log and re-raise unexpected exceptions
            self.log(f"An unexpected error occurred: {str(e)}")
            raise

    def clear_input(self):
        """Clears the input data and log window."""
        self.result_table.setRowCount(0)
        self.log_window.clear()
        self.data = []

    def log(self, message):
        """
        Logs a message to the log window.

        Args:
            message (str): The message to log.
        """
        self.log_window.appendPlainText(message)

    def search_ebay(self):
        """Starts the eBay search process."""
        if not self.data:
            self.log("No data to process. Please load an Excel file first.")
            return

        if not self.setup_selenium():
            self.log("Selenium setup failed. Please check the paths and try again.")
            return

        self.worker = Worker(self.driver, self.data)
        self.worker.update_result.connect(self.update_table)
        self.worker.log_message.connect(self.log)
        self.worker.search_complete.connect(self.on_search_complete)
        self.worker.start()


    def cancel_search(self):
        """Cancels the ongoing search process."""
        if self.worker:
            self.worker.stop()

    def update_table(self, result):
        """
        Updates the result table with search results.

        Args:
            result (list): List of search results to add to the table.
        """
        self.log(f"Updating table with result: {result}")
        row_position = self.result_table.rowCount()
        self.result_table.insertRow(row_position)

        for i, item in enumerate(result):
            self.result_table.setItem(row_position, i, QTableWidgetItem(str(item)))

        self.log(f"Row inserted at position {row_position} with data {result}")
        QCoreApplication.processEvents()

    def on_search_complete(self):
        """Handles the search complete event."""
        self.log("Search completed.")

    def export_results(self):
        """Exports the search results to an Excel file."""
        row_count = self.result_table.rowCount()
        if row_count == 0:
            self.log("No results to export.")
            return

        data = []
        for row in range(row_count):
            row_data = []
            for col in range(self.result_table.columnCount()):
                item = self.result_table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        df = pd.DataFrame(data, columns=["Title", "Author", "ISBN",
                                         "Avg Sold Price", "Sold Price Range",
                                         "Avg Shipping", "Total Sellers"])

        default_filename = f"book_values_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Results", default_filename,
            "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_path:
            df.to_excel(file_path, index=False)
            self.log(f"Results exported to {file_path}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = EbaySearcher()
    ex.show()
    sys.exit(app.exec_())
