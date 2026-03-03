📊 Automatic Header Row Detection for Excel Files
🚀 Overview

This project implements a heuristic-based algorithm to automatically detect the header row in semi-structured Excel files.

In real-world business spreadsheets, column headers are often not located on the first row. Files may contain:

Titles

Metadata

Blank rows

Notes

Logos or formatting elements

This tool scans the top rows of a worksheet and determines which row most likely contains the column names.
The detected row is then automatically passed to pandas for structured data loading.

🎯 Problem Statement

The standard function:

pandas.read_excel()

assumes the header is located at row 0.

In practice, this assumption frequently fails when working with business or operational Excel files.

Objective

Automatically detect the correct header row without manual inspection.

🧠 Algorithm Design

The solution relies on a weighted heuristic scoring system based on four structural signals.

🔢 Scoring Formula
Score=3×FillRatio+4×StringRatio+3×BoldRatio−0.1×RowIndex
Score=3×FillRatio+4×StringRatio+3×BoldRatio−0.1×RowIndex

The row with the highest score is selected as the header.

📌 Feature Engineering
1️⃣ Fill Ratio (Weight = 3)

Ratio of non-empty cells in a row.

Headers usually populate most columns and are rarely sparse.

2️⃣ String Ratio (Weight = 4)

Proportion of text values among non-empty cells.

Column names are typically text-based, whereas data rows often contain numbers or mixed types.

This is the most influential feature.

3️⃣ Bold Ratio (Weight = 3)

Proportion of bold-formatted cells.

Headers are often visually emphasized using bold formatting in professional spreadsheets.

4️⃣ Depth Penalty
−0.1×RowIndex
−0.1×RowIndex

Rows appearing further down in the file are penalized, reflecting the assumption that headers are generally near the top.

⚙️ Implementation Workflow

Load Excel file using openpyxl

Scan first N rows (default = 30)

Compute feature values for each row

Calculate weighted score

Select row with maximum score

Load dataframe using pandas with detected header

🧪 Example Usage
header_line = detect_header_row("file.xlsx")
df = pd.read_excel("file.xlsx", header=header_line - 1)
⚠ Important

openpyxl uses 1-based indexing

pandas uses 0-based indexing

Therefore, the detected row index must be adjusted by subtracting 1.

📈 Time Complexity

Let:

n = number of scanned rows (≤ 30)

m = number of columns

Time complexity:

O(n × m)

The algorithm is efficient and scalable for large spreadsheets.

🏗 Architecture Pipeline
Excel File
     ↓
Row Scan (Top N Rows)
     ↓
Feature Extraction
     ↓
Weighted Scoring
     ↓
Best Row Selection
     ↓
Load DataFrame in Pandas
✅ Project Strengths

Fully interpretable heuristic model

No training data required

Fast and lightweight

Works on real-world business Excel files

Seamless integration with pandas

Formatting-aware detection (bold support)

🔬 Ideal Use Cases

Data preprocessing pipelines

Automated ETL workflows

Business data ingestion systems

Analytics projects involving messy Excel files

🔮 Possible Improvements

Multi-sheet detection support

Confidence score thresholding

Additional statistical features (variance, uniqueness ratio)

Hybrid ML-based enhancement

Logging and debugging mode

📦 Dependencies
pip install openpyxl pandas
🎓 Learning Outcomes

This project demonstrates:

Heuristic algorithm design

Feature engineering for semi-structured data

Excel metadata extraction

Practical data engineering problem-solving

Clean integration between openpyxl and pandas

📌 Technical Summary

This project implements a semi-structure-aware heuristic scoring system for dynamic header detection in spreadsheet data.

It leverages structural, statistical, and formatting signals to robustly identify column headers in real-world Excel files.

If you'd like, I can now:

Make it more “Data Engineering portfolio” oriented

Add badges and project structure

Or make a version tailored specifically for internship applications in Data / Analytics 🚀
