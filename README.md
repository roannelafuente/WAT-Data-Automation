# WAT Data Automation Tool – v1.0.0

## 📖 Description
First release of the WAT Data Automation Tool, designed to automate semiconductor wafer acceptance test (WAT) deliverables.  
This tool converts raw `.wat` files into structured Excel workbooks, generates per‑Unit and per‑Wafer summaries, validates specifications, and produces interactive capability plots with Cp/Cpk statistics. Built with Python, Tkinter, OpenPyXL, Matplotlib, and NumPy, it streamlines engineering workflows and ensures reproducible, audit‑ready insights.

---

## 🚀 Features

### .wat to Excel Conversion
- Core functionality: converts raw `.wat` files into structured **Excel workbooks**.  
- Workbook naming is dynamic, e.g., `Wafer 1~N` where *N* is fetched from the `.wat` file.  
- This converted workbook serves as the **reference source** for generating:
  - **Per Unit Data** sheet  
  - **Per Wafer Summary** sheets  
  - **Summary Sheet**  
  - **Histogram Plot (Capability Plot)**

### Per Unit Data
- Extracts wafer IDs, parameters, and site measurements from the converted `.wat` file.  
- Outputs a structured **Per Unit Data** sheet for traceability and analysis.  

### Per Wafer
- Automatically builds a per Wafer sheet with:
  - Site values transposed per wafer
  - AVERAGE and STDEV formulas
  - Spec HI, Spec LO, and Unit mapping
- Professional formatting with merged headers, auto‑fit columns, and borders.

### Summary Sheet
- Consolidates statistics per parameter from the converted `.wat` workbook.  
- Includes:
  - **Parameter** name  
  - **Spec HI** and **Spec LO** limits  
  - **Mean** and **Standard Deviation** values  
  - **CpK**, **CpK Hi**, and **CpK Lo** capability indices  
- Professionally formatted with merged headers, auto‑fit columns, borders, and fills.  
- Provides a quick, at‑a‑glance overview of wafer performance and spec compliance.  

### Histogram Plot (Capability Plot)
- Uses the converted `.wat` workbook as the reference for plotting.  
- Interactive histogram viewer with radio buttons to browse parameters.  
- Normal curve overlay extended to ±3σ for accurate visualization.  
- Cp, Cpk, Cpk Hi, and Cpk Lo statistics displayed in a side panel.  
- GUI interface with scrollable logs, success/error messages, and polished layout.

---

## 🛠️ Tech Stack
- Python (automation & GUI)  
- Tkinter (user interface)  
- OpenPyXL (Excel file handling)  
- Matplotlib (histogram visualization)  
- NumPy (statistics)  

---

## 📦 Required Packages
The dependencies are listed in [`requirements.txt`](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/requirements.txt).

Install them with:
```bash
pip install -r requirements.txt
```

---

## ⚡ Usage Workflow
- Load Input File
- Select a raw .wat file (e.g., [Dummy data.wat](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Dummy%20data.wat) ).
- Run Automation
- Converts the .wat file into Excel deliverables.
- Automatically generates:
• Per Unit Data sheet
• Per Wafer Summary sheet
- Generate Summary Sheet
- Consolidates statistics per parameter into a Summary sheet.
- Includes Spec HI/LO, Mean, Std Dev, CpK, CpK Hi, and CpK Lo.
- Professionally formatted with merged headers, auto‑fit columns, borders, and fills.
- Explore Histogram Plot
- Use the interactive GUI to browse parameters via radio buttons.
- View histograms with normal curve overlays (±3σ).
- Cp/Cpk statistics displayed in a side panel for quick capability analysis.
- Review Outputs
- Excel workbook: [Dummy data.xlsx](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Dummy%20data.xlsx)
- Dashboard interface: [WAT Data Automation Dashboard.png](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/WAT%20Data%20Automation%20Dashboard.png)
- Capability plot: [Sample capability plot output.png](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Sample%20capability%20plot%20output.png)
 
---
 
## 📂 Sample Files
- Input: [Dummy data.wat](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Dummy%20data.wat)  
- Output: [Dummy data.xlsx](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Dummy%20data.xlsx)  
- Histogram Screenshot: [Sample capability plot output.png](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Sample%20capability%20plot%20output.png)  
- Dashboard Screenshot: [WAT Data Automation Dashboard.png](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/WAT%20Data%20Automation%20Dashboard.png)  

These files are included for demonstration purposes so recruiters and collaborators can quickly test the workflow and visualize the results.

---

## 📸 Screenshots
- **Dashboard Interface**  
  ![WAT Data Automation Dashboard](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/WAT%20Data%20Automation%20Dashboard.png)

- **Capability Plot Output**  
  ![Sample capability plot output](https://github.com/roannelafuente/WAT-Data-Automation/blob/main/Sample%20capability%20plot%20output.png)

---

## 🌟 Impact
- Reduces manual effort in semiconductor WAT deliverables preparation by automating .wat → Excel conversions.
- Ensures reproducibility and audit‑ready insights with deterministic spec mapping.
- Improves accuracy in yield and defect tracking through Cp/Cpk capability statistics.
- Enhances usability with a polished GUI, making engineering workflows faster and more transparent.

---

## 📦 Download
Release v1.0.0 with histogram viewer and spec mapping is available here:  
➡️ [Download WAT Data Automation Tool v1.0.0](https://github.com/roannelafuente/WAT-Data-Automation/releases/download/v1.0.0/WAT.Data.Automation.zip)

---

## 👩‍💻 Author
**Rose Anne Lafuente**  
Licensed Electronics Engineer | Product Engineer II | Python Automation  
GitHub: [@roannelafuente](https://github.com/roannelafuente)
