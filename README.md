# PM-date-app
For Josh to calculate the difference between the date of primary mets and other blood dates. (Remember to keep the headings the same)

# Weekly Events Excel Processor

This Python script processes an Excel file containing biobank sample data, calculates the weeks between key events, colors weeks on a new Excel sheet, and exports numeric week values to a text file.

---

## Features

- Calculates weeks from "Date of PM" to "Mets Development", "Date of last follow up/death", and dates embedded in "Whole blood" and "Follow up bloods" columns.
- Generates a new Excel sheet with weeks highlighted in different colors for each event type.
- Exports a text file listing the numeric week values per biobank sample.
- Does not modify the original Excel sheet.

---

## Requirements

- Python 3.7+
- Conda package manager (recommended)

---

## Installation and Setup

### 1. Install [Miniconda](https://docs.conda.io/en/latest/miniconda.html) or [Anaconda](https://www.anaconda.com/products/distribution)

- Download and install Miniconda or Anaconda for your operating system.
- Also download a version of python (just go to the python website or through install univeristy applications)

### 2. Create and activate a new conda environment

```bash
conda create -n PMcalc python=3.10 -y
conda activate PMcalc
```

```
pip install pandas openpyxl
```
### 3. Save app.py
Place this app.py and the excel sheet in the same place and open the terminal and go to this location.

# Usage
To use, in the terminal and location where your app.py and excel is use the script below:
```
python app.py excelname.xlsx
```
Easy as that

