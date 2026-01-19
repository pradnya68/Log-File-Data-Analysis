# Log-File-Data-Analysis

## Overview
This project automates the analysis of Hydrocleaner machine log files using Python. The script parses multiple `.TXT` log files, extracts batch and dispensing data (recipe, additives, Hydra pump time, etc.), and generates a structured Excel report containing summaries, raw data, and visualizations of recipe usage. This helps in monitoring machine performance, identifying inconsistencies, and optimizing operations.

## Key Features
- Parse multiple `.TXT` log files and extract relevant dispensing information
- Map additive IDs to chemical names for clarity
- Calculate Hydra activation times, additive usage, progress, and actual dispensed liters
- Combine all parsed data into a single structured Excel report
- Generate individual logs per file and recipe usage charts
- Visualize recipe execution frequency using bar charts
- Automated workflow reduces manual log inspection and errors

## Technologies Used
- Python 3.x
- Pandas
- XlsxWriter
- Regular expressions (`re`)
- Glob & OS modules for file handling
