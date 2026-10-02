# 🏫 CBSE LOC Data Validator & Auto-Cleaner

An automated Excel VBA tool designed to clean, standardize, and validate List of Candidates (LOC) registration data across 50+ schools prior to uploading to the CBSE Board portal.

## 🎯 Problem Statement
School admin departments manually enter student enrollment data across 33+ columns. Inconsistent date formats, lower/mixed casing, and missing mandatory fields often result in portal upload rejections.

## ✨ Key Features
- **Data Standardization (Auto-Clean):** Auto-converts text fields to `UPPERCASE`, normalizes gender/category codes, and formats admission dates.
- **Dynamic Header Lookup:** Maps headers dynamically using case-insensitive search to handle shifting column layouts.
- **Automated Error Logging:** Scans mandatory fields and logs errors with exact row numbers in a dedicated `Error` sheet.
- **Portal Compliance:** Ensures a 0% rejection rate on the CBSE upload portal.

## 📂 Project Structure
- `Validate_LOC_FullCheck.bas` - VBA Macro Source Code
- `Sample_LOC_Class11.xlsx` - Test Dataset
- `README.md` - Documentation

## 🛠️ How to Use
1. Open your Excel workbook containing the `CLASSXI` sheet.
2. Open VBA Editor (`Alt + F11`), go to `File > Import File`, and select `Validate_LOC_FullCheck.bas`.
3. Run the macro `Validate_LOC_FullCheck_Optimized`.
4. Review the generated `Error` sheet for flagged issues.
