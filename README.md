📊 Excel Sales Report Generator

A simple Python automation tool that reads an Excel file with product sales data and generates a clean, formatted Excel report and a matching PDF version using a graphical user interface (GUI).

---

> 🚀 Features

- ✅ GUI interface to select the Excel file (using PySimpleGUI)
- ✅ Calculates and appends Total and Average sales
- ✅ Generates a professionally styled Excel report
- ✅ Creates a matching PDF report
- ✅ Automatically timestamps output files

---

> 📂 Input Format

The input Excel file must contain at least the following columns:

| Product | Sales |
|---------|-------|
| Item A  | 100   |
| Item B  | 250   |

---

> 📦 Output

After processing, two files are generated in the same folder:

- `cleaned_YYYY-MM-DD_HH-MM.xlsx` – Clean and formatted Excel report
- `report_YYYY-MM-DD_HH-MM.pdf` – PDF version of the report

---

> 🛠️ Requirements

Install dependencies with:

python -m pip uninstall PySimpleGUI
python -m pip cache purge
python -m pip install --upgrade --extra-index-url https://PySimpleGUI.net/install PySimpleGUI
pip install pandas openpyxl PySimpleGUI fpdf


> Contact

For questions or suggestions, feel free to email: saloom8143@gmail.com
