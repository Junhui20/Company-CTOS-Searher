# CTOS Company Scraper

A desktop GUI tool for batch searching Malaysian company information from the [CTOS Business Report](https://businessreport.ctoscredit.com.my) platform. Built with Python and CustomTkinter.

<!-- 放一張主畫面截圖 -->
![Main Interface](screenshots/main.png)

## Demo

<!-- 把影片上傳到 GitHub 後，替換下面的連結 -->
https://github.com/user-attachments/assets/YOUR_VIDEO_ID

---

## Features

- **Single & Batch Search** - Search one company or load hundreds via CSV / paste list
- **Smart Auto-Matching** - Automatically resolves ambiguous results with name matching algorithms
- **Multi-Threaded** - Up to 5 browser instances in parallel with configurable delay
- **Unattended Mode** - Skip ambiguous results automatically for overnight batch jobs
- **Session History & Resume** - All sessions saved to SQLite, resume interrupted jobs anytime
- **Flexible Export** - Export to Excel (.xlsx), CSV, or copy to clipboard

---

## Installation

### Prerequisites
- Python 3.9+
- Google Chrome browser

### Setup

```bash
git clone https://github.com/your-username/Company-CTOS-Searcher.git
cd Company-CTOS-Searcher
pip install -r requirements.txt
```

### Dependencies
| Package | Purpose |
|---------|---------|
| selenium | Browser automation |
| webdriver-manager | Auto-manage ChromeDriver |
| customtkinter | Modern dark-themed GUI |
| pandas | Data handling & export |
| openpyxl | Excel file export |

---

## Usage

```bash
python main.py
```

1. Type a company name and click **Add One**, or load a CSV / paste a list
2. Click **Start** to begin scraping
3. Click **Export Results** to save as Excel / CSV

### CSV Format

```
Company Name
ABC Sdn Bhd
XYZ Enterprise
DEF Plt
```

Header row is optional. Company names should be in the first column.

---

## Project Structure

```
├── main.py              # Entry point
├── requirements.txt     # Dependencies
├── scrape_history.db    # SQLite database (auto-created)
├── app/
│   ├── gui.py           # Main window & logic
│   ├── scraper.py       # CTOS scraper (Selenium)
│   ├── history.py       # Session history (SQLite)
│   └── dialogs.py       # Dialog windows
└── tests/
    ├── test_confidence.py
    ├── test_excel.py
    └── test_history.py
```

---

## License

This project is for internal use.
