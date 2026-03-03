# 📊 Dynamic Header Detection for Excel Ingestion Pipelines

## 🚀 Problem identification

In production environments, Excel files rarely follow clean, standardized formats. They often include:

- Report titles
- Metadata rows
- Blank separators
- Notes and comments
- Formatting artifacts

As a result, assuming that headers are always on row 0 (the default behavior in `pandas.read_excel`) is unreliable.

This project solves that problem by introducing a lightweight, interpretable algorithm that dynamically identifies the correct header row before loading the data into a structured DataFrame.

---

## 🎯 Why This Matters in Data Engineering

In data pipelines, ingestion is often the most fragile step.

Incorrect header detection can lead to:

- Misaligned columns
- Corrupted schemas
- Silent data quality issues
- Downstream transformation errors

This project demonstrates how to build a robust preprocessing layer that increases ingestion reliability without introducing heavy machine learning dependencies.

---

## 🧠 Algorithmic Strategy

The system uses a heuristic scoring model based on structural and semantic signals extracted from each row.

Instead of assuming where the header is, the algorithm evaluates each candidate row and selects the one that maximizes a weighted score.

### 🔢 Scoring Formula

Score = 3 × FillRatio + 4 × StringRatio + 3 × BoldRatio − 0.1 × RowIndex


The row with the highest score is selected as the header.

---

## 🔍 Feature Engineering Explained

Each row is evaluated using four signals:

### 1️⃣ Fill Ratio (Structural Density)

**Definition:** Proportion of non-empty cells in the row.

**Why it works:** Header rows typically define all columns and are therefore dense.

**Engineering Insight:** Sparse rows are more likely to be metadata or decorative content.

**Weight:** 3

---

### 2️⃣ String Ratio (Semantic Signal)

**Definition:** Proportion of string values among non-empty cells.

**Why it works:** Headers are usually composed of column names (text), while data rows contain numeric or mixed values.

**Engineering Insight:** This is the strongest discriminator between schema rows and data rows.

**Weight:** 4

---

### 3️⃣ Bold Ratio (Formatting Metadata)

**Definition:** Proportion of cells formatted in bold.

**Why it works:** Professional spreadsheets often visually emphasize headers.

**Engineering Insight:** Leveraging formatting metadata improves robustness in business datasets.

**Weight:** 3

---

### 4️⃣ Depth Penalty (Positional Prior)

**Definition:** Penalty proportional to the row index.

**Why it works:** Headers are generally near the top of the file.

**Engineering Insight:** Introduces a soft prior without hardcoding assumptions.

**Penalty:** −0.1 × RowIndex

---

## 🧪 Example Usage

```python

header_line = detect_header_row("file.xlsx")
df = pd.read_excel("file.xlsx", header=header_line - 1)

```

Important Indexing Note
openpyxl uses 1-based indexing
pandas uses 0-based indexing
Therefore, the detected row index must be adjusted by subtracting 1.

## 📈 Performance Characteristics

Let:
n = number of scanned rows (≤ 30)

m = number of columns

Time Complexity: O(n × m)

Because n is bounded and small, the algorithm is computationally inexpensive and suitable for batch ingestion pipelines.


## 🏗 Engineering Design Principles

This project reflects key data engineering concepts:

Principle	Description
✅ Deterministic and Interpretable	No black-box models. Fully explainable logic.

✅ Lightweight	No training data required. No model persistence.

✅ Scalable	Efficient even for wide spreadsheets.

✅ Formatting-Aware	Uses structural and formatting metadata.

✅ Pipeline-Ready	Designed to integrate directly into ETL workflows.


## 🔬 Real-World Use Cases

Automated Excel ingestion pipelines

ETL preprocessing layers

Data lake raw ingestion standardization

Analytics workflows with heterogeneous input files

Internal business reporting automation


## 📌 Technical Summary

This project implements a semi-structure-aware heuristic scoring engine for dynamic header detection in spreadsheet data.
By combining structural density, semantic signals, formatting metadata, and positional priors, the system reliably identifies schema rows in non-standard Excel files.
It demonstrates how thoughtful feature engineering can significantly improve data ingestion robustness without introducing unnecessary model complexity.
