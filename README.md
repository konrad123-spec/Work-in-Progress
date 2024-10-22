# Data Management Automation - Analysis for Microsoft Excel

This project involves a set of macros designed to automate the process of data consolidation, filtering, and organization within *Analysis for Microsoft Excel*. The macros streamline data management, making it more efficient by handling multiple tasks like adding columns, updating sheets, and processing specific data formats.

## Macro Overview

The primary macro consolidates monthly data from various files into the main workbook. It filters, formats, and organizes data, ultimately providing summarized information for the following categories:

- **DAMP**
- **GAPI**
- **DAEN**
- **LABA**

The consolidated data is placed in the main sheet for further analysis.

## Key Features

### 1. **Adding Columns Macro**
- Selects the entire **Column O** in the active sheet.
- Inserts a new column to the right of **Column O**.
- Copies the content from **Column L** and pastes it into the newly created column.

### 2. **Tabs to Analysis Macro**
This macro automates the process of updating the "Analisi" sheets across multiple tabs.

- Prompts the user to input the current period via an `InputBox` (format: "yymm").
- Iterates through the following sheet names: **DAMP**, **GAPI**, **DAEN**, and **LABA**.
- Clears specific ranges in the corresponding **Analisi** sheet.
- Copies data from the external file **"WIP.xlsx"** and pastes it into the appropriate **Analisi** sheet.
- Calculates formulas in **Columns L, M, and N** based on the copied data.
- Fills down these formulas to cover all relevant rows.
- Special cases:
  - For **DAMP** and **DAEN** sheets, negative values are replaced with **0**, and those cells are highlighted with a green background.

### 3. **Files to Sheets Macro**
This macro updates the "ODA" sheets by pulling data from an external workbook.

- Prompts the user to input the period again via an `InputBox`.
- Clears specified ranges in the **ODA** sheets corresponding to the current array item.
- Copies data from **"ODA.xlsm"** and pastes it into the **ODA** sheets.
- Calculates formulas in **Columns L and M**.
- Fills down the formulas for the correct number of rows.
- Filters and deletes rows where **Column M** contains specific values: “TERZI” and “Italia.”
- Replaces any slashes ("/") in **Column F** with empty strings.



