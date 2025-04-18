# Excel Data Matcher 📊  
A practical desktop tool for automating the comparison of logistics or operational data across two Excel files, with a clean UI and instant Excel report generation.

---

## 🚀 Problem

Before this tool, comparing logistics or financial data from different sources (e.g., internal system vs partner report) meant:

- Manually opening both Excel files
- Looking for matching invoice numbers, values, or dates
- Copying and comparing row by row
- Spending hours trying to spot missing or mismatched records

This process was tedious, error-prone, and highly inefficient — especially with large datasets (200+ rows).

---

## ✅ Solution

This Python-based GUI app automates the full process by:

- Letting the user select two Excel files through a simple interface
- Comparing the data based on a key column (`Invoice`)
- Generating a clean Excel report with:
  - ✅ Matched records
  - ❌ Records only found in File 1
  - ❌ Records only found in File 2

All in under **10 seconds** — no copy-pasting or manual formulas required.

---

## 📈 Impact

This automation reduced file comparison time from **2–3 hours** to **less than a minute**, even for large files.  
It eliminates human error and makes data verification easy for logistics managers, operations teams, or analysts.

> ⏱️ Time before automation: ~10 hours/week  
> ⚡ Time with this tool: ~30 minutes/week  
> 💸 Net savings: ~9.5 hours/week of manual effort eliminated

💡 I'm proud to have built a lightweight, user-friendly solution that solves a real admin pain point using Python.

---

## 🧠 Skills & Technologies

- Python 🐍  
- pandas  
- openpyxl  
- ETL-style comparison  
- Excel report generation  
- Tkinter GUI

---

## 📂 Project Structure

```
excel-data-matcher/
├── matcher.py              # Main script with GUI and logic
├── requirements.txt        # Dependencies
├── sample_data/            # Sample Excel files
│   ├── export1.xlsx
│   └── export2.xlsx
└── matched_report.xlsx     # Output with results
```

---

## 🔧 How to Use

1. Clone the repo:

```bash
git clone https://github.com/YOUR_USERNAME/excel-data-matcher.git
cd excel-data-matcher
```

2. (Optional) Create and activate a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

3. Install the dependencies:

```bash
pip install -r requirements.txt
```

4. Run the app:

```bash
python matcher.py
```

5. Select your two Excel files using the UI and click the button.  
   Output will be saved to `sample_data/matched_report.xlsx` 📁

## 🔐 Data Privacy

All sample files in this repository are fictional or anonymized.  
This tool can be used safely for demos, testing, or real operations.

---

## 🧰 Future Improvements

- Allow choosing different key columns from the UI  
- Highlight specific mismatches (e.g. values differ for same invoice)  
- Add export to PDF summary  
- Web-based version with drag-and-drop UI  
- Logging & advanced error handling

---

## ✨ Author

Created by [@Kseniatyschuk](https://github.com/Kseniatyschuk) —  
for my personal portfolio to showcase practical Python skills, GUI building, and data automation.
