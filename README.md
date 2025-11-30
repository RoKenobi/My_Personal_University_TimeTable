# NTU Timetable Optimizer

Optimizes your NUS course timetable to **minimize campus days** while avoiding time conflicts.

## 🛠️ Requirements
- Python 3.8+
- Packages: `streamlit`, `ortools`, `pandas`, `openpyxl`

## 📁 Files Needed
For this you need to scrape data and organise them into these files below 
- `Table1.xlsx` – Lecture data  
- `Table2.xlsx` – Tutorial/Lab index data

## ▶️ How to Run
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   streamlit run plan_app.py
