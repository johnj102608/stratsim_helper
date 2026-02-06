https://github.com/johnj102608/stratsim_helper

# StratSim Helper

This program reads through StratSim "Financial Details" Excel files and populates a pre-built Excel dashboard template with the correct values.

---

## What This Script Does

For each StratSim round (year):

1. Reads input Excel files like  
   `Competition - Financial Summary - Year 1.xlsx`

2. Extracts metric → value pairs from each firm's  
   `Financial Details for A`, `B`, `C`, etc. sheets

3. Matches metrics by **name** to the dashboard template

4. Writes values into the correct cells for each firm

5. Outputs a fully populated dashboard Excel file

The script does **not** depend on:
- block titles (Income Statement / Balance Sheet)
- fixed column letters
- fixed row numbers

Everything is detected automatically.

---

## Folder Structure
Everything in the same folder.

project_folder/
│
├── StratSimHelper.exe
├── config.json # configuration file
├── metric_aliases.json # optional name-mapping file
│
├── Pre-built StratSim dashboard template
├── StratSim "Financial Details" file
│
└── StratSim_Dashboard_UPDATED.xlsx # output (created by script)

---

## How to run this
After having the template file and "Financial Details" ready, run main.exe, and you should see the output file generated.

---

## config.json
Edit this file to change configs.
Brief explanation for each row:
  "dashboard_template_name": name of the pre-built template,
  "output_dashboard_name": name of the output file,
  "input_prefix": prefix of the StratSim "Financial Details" files,
  "financial_sheet_prefix": prefix of each sheet within the StratSim "Financial Details" files,
  "firms": a list of firm names,
  "year_sheet_prefix": prefix of each year-sheet within the template,
  "firm_prefix": prefix of firms within the template,
  "scan_max_rows": the number of rows this script should scan for data,
  "scan_max_cols": the number of columns this script should scan for data,
  "look_right_max": how far right a value MAY appear after a metric name  ----For example: |"Inventory"|      |      |  100  |

If this file is missing, the script will run with this following default config:
{
  "dashboard_template_name": "StratSim_Dashboard_Template.xlsx",
  "output_dashboard_name": "StratSim_Dashboard_UPDATED.xlsx",
  "input_prefix": "Competition - Financial Summary - Year ",
  "financial_sheet_prefix": "Financial Details for ",
  "firms": ["A","B","C","D","E","F","G"],
  "year_sheet_prefix": "Year ",
  "firm_prefix": "FIRM ",
  "scan_max_rows": 250,
  "scan_max_cols": 30,
  "look_right_max": 4
}

---

## metric_aliases.json
Edit this file if metric names differ between input files and the dashboard.
The script WILL run if this file is missing/empty.
Sample format:
{
  "metric name in StratSim Financial Details": "metric name in template",
  "1 Year CD Investment": "Short-term securities",
  "Starting Inventory": "Beg. Inventory",
  "Less Ending Inventory": "End. Inventory"
}
