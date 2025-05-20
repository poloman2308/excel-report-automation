# 📊 Excel Report Automation

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Status](https://img.shields.io/badge/status-production-ready-brightgreen)

> A modular, automated Excel report generator with pivot tables, charts, conditional formatting, and CLI support — deployable via Windows Task Scheduler.

---

## 📌 Features

- ✅ Clean, reusable **object-oriented design**
- ✅ CLI-ready with `argparse` — supports dynamic input/output
- ✅ Auto-generated:
  - Summary table
  - Pivot table by Region/Product
  - Bar chart
  - Conditional formatting
- ✅ Supports logo embedding
- ✅ Autosized columns for all sheets
- ✅ Daily scheduling via Windows Task Scheduler

---

## 🗂️ Folder Structure

```
excel_report_automation/
├── data/ # CSV files (input)
├── templates/ # Logo assets
├── reports/ # Output Excel files
├── report_generator/ # Modular Python logic
│ ├── init.py
│ ├── base_report.py
│ ├── sales_report.py
│ └── utils.py
├── main.py # CLI entry point
├── run_report.bat # Windows automation script
├── requirements.txt
└── README.md
```

---

## 🚀 Getting Started

### 1️⃣ Clone the Repository

```
git clone https://github.com/YOUR_USERNAME/excel-report-automation.git
cd excel-report-automation
```

---

### 2️⃣ Install Dependencies

```
pip install -r requirements.txt
```

---

### 3️⃣ Run the Report Generator

```
python main.py \
  --input data/sales_march.csv \
  --output_dir reports \
  --logo templates/company_logo.png
```

---

## ⚙️ Automation (Optional)
### Schedule with Windows Task Scheduler

#### 1️⃣ Create a run_report.bat file

```
@echo off
cd /d C:\Path\To\excel_report_automation
python main.py --input data\sales_march.csv --output_dir reports --logo templates\company_logo.png
```

---

#### 2️⃣ Use Task Scheduler to run it daily, weekly, or on login

---

## 🔧 Customization

```
Extend BaseReport to define new report types
Add Excel formatting/styling in sales_report.py
Upload results to SharePoint, OneDrive, or email via smtplib
```

---

## 🧪 Example Output

```
RawData sheet: all records from CSV
Summary sheet: revenue by region/product
Pivot sheet: visual pivot table + chart
```

---

## 🤝 Contribution
Pull requests are welcome! For major changes, please open an issue first to discuss what you’d like to add.

```
# Format code with black
black report_generator/
```

---

## 📄 License
This project is licensed under the MIT License.

---

## 👨‍💻 Author
**Derek Acevedo**
GitHub • LinkedIn
