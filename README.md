
# 📈 Financial Crawler (台灣股票財報爬蟲)
This project crawls key financial data like **PER (本益比)**, **PBR (股價淨值比)**, and other details for multiple Taiwanese stock companies from online sources, and exports everything nicely into an Excel file.

---

## 📋 Prerequisites
Make sure you have installed:
- **Python 3.11+**
- **Google Chrome** browser (latest version)

---

## 🛠️ Installation
1. **Clone this repository**:
   ```bash
   git clone https://github.com/your-username/your-repo-name.git
   cd your-repo-name
   ```

2. **Install required Python packages**:
   ```bash
   pip install -r requirements.txt
   ```
   If requirements.txt is not yet created, you can install manually:
   ```bash
   pip install selenium pandas beautifulsoup4 webdriver-manager openpyxl
   ```

---

## 🚀 How to Run
### Option 1: Using Command Line
1. Open your terminal or command prompt.
2. Navigate to your project folder.
3. Run the script:
   ```bash
   python crawler.py
   ```
### Option 2: Windows Double-Click Method

- Simply double-click on the crawler.py file in Windows Explorer
- A command prompt window will open automatically


### Enter the stock codes you want to search for (separate them by spaces), like:
   ```
   2330 2303 3515
   ```

The program will:
- Fetch the latest 5 years(start from last year) of financial data (can be changed inside the script).
- Create or append results to a file called `crawler.xlsx`.
- Name each worksheet based on current time and date.

After it's done, you will see:
```
✅ Crawling finished!
```
Open crawler.xlsx to see your results!

---

## ⚙️ Configuration
- **Years of data**:
  Change this line inside `crawler.py`:
  ```python
  years = 5  # Set between 1 ~ 7
  ```

- **Headless mode**:
  If you want to see the browser working (not headless), just comment out this line:
  ```python
  options.add_argument("--headless")
  ```

---

## 🧹 Notes
- If Excel is open while the script is writing, it will cause an error.
  → Make sure to close `crawler.xlsx` before running the script.
- If your Chrome browser updates, webdriver-manager will automatically manage the correct driver version.

---

## ✨ Features
- Fetches company name, PER, PBR, and other financial indicators.
- Supports multiple stock codes at once.
- Outputs clean, organized Excel sheets.
- Fully automatic ChromeDriver management.
- BeautifulSoup + Selenium combo for reliability.
