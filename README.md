# BOM Checker Tool (Syspro-Compatible)

A Python utility designed to compare two **Bill of Materials (BOM)** Excel files exported from **Syspro**. This tool quickly identifies:

- Part numbers that are missing from either BOM
- Differences in part quantities for shared items

Ideal for manufacturing engineers, purchasing teams, and anyone working with Syspro BOM exports.

---

## 🔧 Tailored for Syspro

This script is specifically configured to handle BOMs **exported directly from Syspro**, using:
- A 5-row header skip
- `Stock code` as the part number column
- `Quantity per` as the quantity field

It removes punctuation and standardizes case to avoid false mismatches.

---

## 📂 Features

- ✅ Compare part numbers across two BOM files
- ⚖️ Flag mismatched quantities for matching parts
- 🔍 Case-insensitive and punctuation-agnostic comparisons
- 🖱️ File picker interface using Tkinter
- 🔁 Option to rerun comparisons without restarting

---

## 🛠 Requirements

- Python 3.10 or higher
- Required Python libraries:
  - `pandas`
  - `openpyxl`
  - `tkinter` (included with most Python installations)

Install dependencies with:

```bash
pip install pandas openpyxl

