# Data Processing and Statistical Analysis Pipeline for Cyclic Test Data

## Overview

This repository contains a two-stage pipeline designed to preprocess, extract features, and perform statistical analysis on experimental data obtained from cyclic mechanical tests. The code is structured to support batch processing of tabular datasets (e.g., Excel spreadsheets) and to generate intermediate and final results suitable for further interpretation or visualization.

---

## Structure

### Stage 1: Preprocessing and Feature Extraction

The first script in the pipeline performs the following tasks:

- Traverses through a predefined directory structure to locate data files (typically Excel format).
- Parses multiple sheets within each file, filtering out irrelevant metadata sheets.
- For each data sheet, extracts numeric values related to elongation and stress over cycles.
- Segments the loading and unloading parts of the stress-strain loops.
- Computes metrics such as:
  - Area under the stress-strain loop (hysteresis energy),
  - Residual elongation,
  - Hysteresis width,
  - Stress degradation,
  - Various modulus values (loading, unloading, secant).
- Saves the results as individual `.csv` files per cycle and accumulates summary metrics in an `.xlsx` file for subsequent analysis.
- Optionally, visualization code is included (but commented out) for drawing curves.

### Stage 2: Statistical Aggregation

The second script aggregates the results produced in Stage 1 and computes average values over selected cycles. Specifically, it:

- Loads the preprocessed summary Excel file.
- Groups data by sample/material type and test step (e.g., elongation percentage).
- Calculates statistics over the final three cycles for each group, including:
  - Average hysteresis area,
  - Average residual deformation,
  - Average loop width,
  - Average stress loss,
  - Maximum and average stress,
  - Average modulus values.
- Writes the resulting aggregated data into a new Excel workbook.

---

## Requirements

The pipeline is written in Python and relies on the following libraries:

- `pandas`
- `numpy`
- `matplotlib` (optional, for visualization)
- `scipy`
- `openpyxl`
- `xlrd`
- `xlsxwriter`
- `glob2`

To install the required dependencies, run:

```bash
pip install -r requirements.txt
```

You may need to manually include all required libraries if no `requirements.txt` is provided.

---

## Usage

1. Place your raw data files (Excel format) into the designated input directory (as defined in the first script).
2. Run the first script:
   ```bash
   python main.py
   ```
   This will generate:
   - Processed data in `.csv` format
   - A compiled Excel workbook with cycle-wise metrics

3. Run the second script to compute summary statistics:
   ```bash
   python process_statistic.py
   ```
4. You can run next script to draw load-unload loops:
   ```bash
   python draw_last_loop_per_step.py
   ```
The final output will be an aggregated Excel workbook containing descriptive statistics for each material and test condition.

---

## Customization

To adapt this project to your dataset:

- Edit the list of target directories and sheets in the first script to match your folder structure and data organization.
- Modify the list of materials and group labels in the second script to reflect your experimental setup.
- Uncomment and adjust the visualization sections to enable plotting of individual loops.

---

## License

This software is officially registered as a computer program in the Russian Federation under the certificate of state registration № RU 2024667626.

All rights reserved by the author(s). Unauthorized copying, distribution, or modification of the code is prohibited without the explicit permission of the copyright holder.

For academic use, citation of the registration number and author(s) is required. For commercial use or licensing inquiries, please contact the author directly.
