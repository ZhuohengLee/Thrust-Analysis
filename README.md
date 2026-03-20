# Thrust Analysis

This repository contains a single desktop GUI tool for cyclic thrust-data analysis:

- `thrust_analysis.py`

## What It Does

The script loads `.xlsx`, `.xls`, or `.csv` thrust datasets and performs:

- sampling-rate detection from the time column
- baseline drift removal
- low-pass filtering
- FFT-based cycle-period estimation
- thrust cycle detection
- per-file summary metric export

## Run

```bash
python thrust_analysis.py
```

## Requirements

Install the required Python packages:

```bash
pip install pandas numpy matplotlib scipy openpyxl
```

## Input Files

Supported input formats:

- `.xlsx`
- `.xls`
- `.csv`

The tool expects a time column and a thrust/force column and will try to detect them automatically.

