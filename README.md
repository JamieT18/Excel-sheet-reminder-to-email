# License Manager UI

A robust Python desktop application for tracking, visualizing, and managing license due/expiration dates from a large Excel file. Designed for reliability, scalability, and ease of use.

---

## ✨ Features

- **Excel-powered tracking:**  
  Reads an Excel file with hundreds or thousands of locations, each with “Due” and “Expires” dates.
- **Automatic date rollover:**  
  When a due date matches today’s month and day, the year is automatically updated for the next cycle.
- **Big data support:**  
  Clean, scrollable, sortable table UI for effortless navigation of large datasets.
- **Visual highlights:**  
  - **Yellow:** License due today  
  - **Red:** License expires today  
  - **Green:** Upcoming dates (customizable)
- **Search/filter bar:**  
  Instantly find any location by name.
- **Column sorting:**  
  Click any column header to sort by location, due date, or expiration date.
- **Status bar:**  
  Displays last refresh time and total number of locations currently shown.
- **Error handling:**  
  Friendly alerts for missing files, bad data, or incorrect date formats.
- **Data integrity:**  
  Only updates “Due” dates as needed; never deletes or overwrites data unintentionally.
- **Easy refresh:**  
  One-click refresh to reload and update from Excel.
- **Customizable:**  
  Ready for enhancements like exporting, reporting, or multi-user access.

---

## 🚀 Quick Start

1. **Prepare your Excel file**  
   - File name: `license_dates.xlsx` (or update in script)
   - Column 1: Location name  
   - Other columns: Cells with `Due MM.DD.YYYY` and/or `Expires MM.DD.YYYY`
   - Example:
     | Location Name | Due       | Expires      |
     |---------------|-----------|--------------|
     | Place A       | Due 09.15.2025 | Expires 10.31.2025 |
     | Place B       | Due 10.15.2025 | Expires 11.30.2025 |

2. **Install the dependencies**
   ```bash
   pip install pandas openpyxl
   ```

3. **Configure the script**
   - Update the `EXCEL_PATH` variable in `license_manager_ui.py` if your file has a different name or path.

4. **Run the app**
   ```bash
   python license_manager_ui.py
   ```

---

## 🛠 Customization & Advanced Usage

- **Export features:**  
  Add PDF/CSV export for reporting/sharing.
- **Date format settings:**  
  Choose between `MM.DD.YYYY` and `DD.MM.YYYY`.
- **Region/category filtering:**  
  Organize data by region or category for very large datasets.
- **User instructions/help:**  
  Add a simple menu for guidance and about info.
- **Multi-user access:**  
  Adapt for networked/shared environments.

---

## 🧑‍💻 Troubleshooting

- **File not found:**  
  Check that your Excel file path and name match what’s set in `EXCEL_PATH`.
- **Bad data/format:**  
  Ensure your dates are in the format `MM.DD.YYYY`. Fix any empty or misspelled cells.
- **Dependencies missing:**  
  If you see ImportError, run `pip install pandas openpyxl`.

---

## 🔒 Reliability & Best Practices

- **Backup regularly:**  
  Keep copies of your Excel file.
- **Validate files:**  
  Use the app’s error messages to fix formatting issues.
- **Consistent updates:**  
  Only update “Due” dates as needed; review data before making major changes.

---

## 🤝 Contributions

Pull requests, suggestions, and feature requests are welcome!  
For bugs or enhancement ideas, please open an issue.

---

## 👤 Author

JamieT18  
[GitHub Profile](https://github.com/JamieT18)

---

*License Manager UI helps you automate license tracking, avoid missed renewals, and manage large portfolios of locations with clarity and confidence.*
