# eBay Book Price Searcher

eBay Book Price Searcher is a personal desktop application developed to help you determine the current market value of books in your collection by retrieving pricing information from eBay. The tool utilizes Selenium for web scraping and PyQt5 for its user-friendly interface, making it easier for you to manage and assess the potential resale value of your personal library.

## Features

- **Personal Library Evaluation**: Upload an Excel file containing details of books in your personal collection (Title, Author, ISBN) and retrieve their current market value from eBay.
- **Automated eBay Search**: Gathers pricing data including average sold price, price range, average shipping cost, and total sellers for books listed on eBay.
- **User-friendly GUI**: Easy to use interface with real-time log updates and status feedback.
- **Data Export**: Export the search results to an Excel file to keep a personal record of your book collection's value.
- **Delays for Server Friendliness**: Introduces random delays between search requests (15 to 40 seconds) to avoid placing a heavy load on eBay's servers.

## Requirements

This project requires the following Python packages:
- `PyQt5`: Used to create the graphical user interface.
- `Selenium`: Used for automating the web browsing and data extraction from eBay.
- `pandas`: Used for Excel file processing and data management.
- `chromedriver`: Required to enable Selenium to control Chrome for web scraping.

You can install the required dependencies via pip:

```bash
pip install PyQt5 selenium pandas
```

Additionally, download the appropriate version of [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/) that matches your installed version of Google Chrome.

## How It Works

1. **GUI Setup**: The program starts with a graphical interface that prompts you to input paths for Chrome's user data directory, ChromeDriver, and your personal Excel file containing book details.
2. **Data Input**: The Excel file should contain columns for 'Title', 'Author', and 'ISBN' which correspond to the books in your collection.
3. **Search Process**: After clicking the `Search` button, the program automates a search process using Chrome and Selenium to retrieve book prices from eBay’s marketplace.
4. **Time Delays**: The tool includes random delays of **15 to 40 seconds** between requests to ensure it's respectful of eBay's servers.
5. **Results Display**: Results including average sold price, price range, shipping cost, and total sellers are displayed in the GUI.
6. **Export Data**: The data can be exported to an Excel file for future reference or personal documentation.

## Setup Instructions

1. Clone the repository:

    ```bash
    git clone https://github.com/yourusername/ebay-book-price-searcher.git
    cd ebay-book-price-searcher
    ```

2. Install the necessary Python dependencies:

    ```bash
    pip install -r requirements.txt
    ```

3. Download the appropriate version of ChromeDriver for your operating system and ensure it is added to your system’s PATH. Alternatively, specify its location in the application’s input fields.

4. Run the application:

    ```bash
    python ebay_book_price_searcher.py
    ```

5. In the GUI, provide the necessary paths for:
    - Chrome user data directory
    - Chrome profile directory (if applicable)
    - ChromeDriver executable
    - Excel file containing book data

## Excel File Format

The input Excel file should have the following columns:

- `Title`: The book's title
- `Author`: The author of the book
- `ISBN`: The ISBN number of the book (if available)

The program will attempt to search eBay using the ISBN first. If the ISBN is not available, it will search by title and author.

## Selenium Configuration

The application uses Selenium to automate searches on eBay. You need to configure Selenium with Chrome by providing the following:

- **User Data Directory**: Path to your Chrome profile's user data directory, allowing the scraper to use your browsing preferences and avoid login issues.
- **Profile Directory**: The name of the Chrome profile directory you want Selenium to use.
- **ChromeDriver Path**: The location of the ChromeDriver executable on your system.

### Example Paths:
- User Data Directory: `C:/Users/YourUsername/AppData/Local/Google/Chrome/User Data`
- Profile Directory: `Default`
- ChromeDriver Path: `C:/path/to/chromedriver.exe`

## Usage

1. Load the Excel file containing the books in your collection.
2. Click the `Search` button to begin the search.
3. The application will log the progress in the log window and display results in the table.
4. Once the search is complete, you can export the results to an Excel file for future reference by clicking the `Export Results` button.

## Logging

During the search, the log window provides real-time updates, including messages about the current book being processed, the search URL being visited, and any issues encountered (e.g., no results found, or timeouts).

## Example

### 1. Load Book Data
   - The user can load an Excel file containing book details such as title, author, and ISBN.
   - The application logs the number of rows loaded.

### 2. Start Search
   - The search begins after clicking the `Search` button, and real-time logs keep you informed of the process.
   - The tool waits between 15-40 seconds between searches to avoid overloading eBay's servers.

### 3. Export Results
   - After the search is complete, you can export the results into a new Excel file for your records.

## A Note on Rate Limiting

To **respect eBay's servers**, this tool introduces **delays** between each search. These delays ensure the search process is friendly to their systems. Please **limit the number of books** you search in a single session to avoid any issues.

## Disclaimer

**This tool is strictly for personal use**. It was created with the sole purpose of helping me value my personal library. It should **not** be used for commercial purposes, bulk scraping, or any other form of profit-driven activity.
