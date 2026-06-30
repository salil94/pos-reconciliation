# POS Reconciliation Engine

A Python tool that reconciles restaurant POS transaction data against delivery platform records, automatically flagging discrepancies, duplicates, and missing orders.

## What it does

Restaurants selling through delivery platforms need to verify that what the platform reports matches what their POS system recorded. This script automates that process loading data from both sources, matching orders, and categorising every record as matched, a price discrepancy, missing in POS, or missing in the platform system.

## How it works

- Loads POS data from Excel and platform order data from CSV
- Filters by restaurant brand using configurable brand filters
- Flags duplicate order IDs without removing them, using a stable sort and occurrence-based matching so duplicates remain individually matchable
- Matches records on order ID and occurrence, comparing prices to a tolerance threshold
- Generates a structured Excel report with separate sheets for matches, discrepancies, duplicates, and missing records

## Tech stack

Python, pandas, NumPy, openpyxl

## Usage
Run with `all` as the argument to process every configured restaurant brand in sequence.
Sample anonymised data is included in `sample_data/` so you can test the script immediately after cloning.

## Note

Brand names and locations in this version are anonymised for public sharing. The original was built and used in production for a Dubai F&B operator across six restaurant brands and three delivery platforms.
