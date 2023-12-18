# Project README

## Table of Contents

1. [outageMasterScript.py](#outagemasterscriptpy-readme)
2. [availabilityMaster.py](#availabilitymasterpy-readme)
3. [masterGUI.py](#masterguipy-readme)
4. [AVAPADPARMaster.py](#avapadparmasterpy-readme)
5. [powerMaster.py](#powermasterpy-readme)
6. [siteMasterScript.py](#sitemasterscriptpy-readme)



---

## outageMasterScript.py README {#outagemasterscriptpy-readme}

### Overview

This script automates the processing of outage reports from multiple Excel files, applying various rules and modifications to standardize and enhance the data. The primary focus is on utilizing regular expressions (regex) and predefined rules to automate repetitive tasks, saving significant time in data processing.

### Key Features

- **Regex and Custom Rules**: The script employs regular expressions and custom rules to categorize outage reasons, sub-categories, and owners based on predefined patterns.

- **Automated Cascaded To Handling**: The script identifies and processes information related to cascaded sites, updating relevant fields and cleaning up comments.

- **SLA Status Logic**: It modifies the SLA Status based on specific conditions, enhancing accuracy in reporting.

- **Dependency Handling in Subcategories**: The script dynamically adjusts sub-categories based on dependencies, ensuring consistency in data representation.

- **File Renaming**: For Menoufia and Tanta regions, the script renames files to a standardized format, incorporating region, technology, and date information.

- **Integration of Fuzzy Matching**: Fuzzy string matching is used to map new categories to existing ones, enhancing the accuracy of category assignments.

### Usage

#### Input

- **Excel Files**: The script takes one or more input Excel files containing outage report data. Separate inputs are accepted for different types of reports (e.g., 4G, 3G, non-specific).

#### Output

- **Full Outage Report**: The processed data is consolidated and saved into a new Excel file, providing a comprehensive overview of outages.

### Instructions

1. **Prepare Input Files**: Ensure that the input Excel files follow the required format and naming conventions.

2. **Run the Script**: Execute the script with the appropriate command-line arguments, specifying input files and the desired output file.

    ```bash
    python script_name.py -i input1.xlsx input2.xlsx -m input3.xlsx input4.xlsx -o output.xlsx
    ```

3. **Review Output**: Check the generated output Excel file for the consolidated and processed outage report.

4. **File Renaming (Optional)**: If input files are not named according to standards, the script will attempt to rename them automatically.

5. **Customization (Optional)**: Modify the script's `custom_rules` and other sections to adapt it to specific requirements.

### Dependencies

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) for Excel file handling
- [pandas](https://pandas.pydata.org/) for data manipulation
- [argparse](https://docs.python.org/3/library/argparse.html) for command-line argument parsing
- [fuzzywuzzy](https://github.com/seatgeek/fuzzywuzzy) for fuzzy string matching

### Notes

- The script supports both 4G/3G-specific and general outage reports. Ensure proper input file allocation.

- For Menoufia and Tanta regions, the script performs additional file renaming. Review and adjust as needed.

- Customize the script according to specific business rules and requirements by modifying the relevant sections.

- It is recommended to run the script in a virtual environment to manage dependencies effectively.

Feel free to adapt and extend this script based on evolving needs and additional functionality requirements.
---

## AVAPADPARMaster.py README {#avapadparmasterpy-readme}

This script performs the integration of network availability data and commercial power alarm data. It merges the data from multiple sources, updates master sheets, and generates a consolidated output.

### Table of Contents

1. [Overview](#overview)
2. [Usage](#usage)
   - [Integration Process](#integration-process)
   - [Availability Update](#availability-update)
   - [Power Update](#power-update)
3. [Dependencies](#dependencies)
4. [Notes](#notes)

### Overview {#overview}

The script integrates network availability data and commercial power alarm data. It includes the following functionalities:

- **File Selection:** The script looks for availability and power sheets in the specified paths.

- **Vendor Identification:** The script identifies the vendor based on keywords in the file names.

- **Availability Update:**
  - Excludes sites based on an exclusion list.
  - Updates the master network availability sheet.

- **Power Update:**
  - Merges data from multiple power sheets.
  - Adds vendor information, current date, and converts the 'MTTR' column to a numerical format.
  - Verifies alarms against a predefined list in the master power sheet.

- **Site Data Update:**
  - Retrieves availability, power alarm duration (PAD), and power alarm count (PAR) data.
  - Updates a consolidated site data sheet with the latest values.

### Usage {#usage}

#### Integration Process {#integration-process}

To perform the complete integration process (availability update, power update, and site data update):

```bash
python integration_script.py -a "path/to/availability/sheet.xlsx" -p "path/to/power/sheet1.xlsx" "path/to/power/sheet2.xlsx" -t "YYYY-MM-DD"
```

- `-a` or `--avaSheetName`: Path to the availability sheet for the update.
- `-p` or `--powerSheetNames`: Paths to the power sheets for the update. Multiple paths can be provided.
- `-t` or `--time`: Date for which the site data should be updated. If not provided, the script uses the date of the previous day.

#### Availability Update {#availability-update}

To perform only the availability update:

```bash
python integration_script.py -a "path/to/availability/sheet.xlsx"
```

#### Power Update {#power-update}

To perform only the power update:

```bash
python integration_script.py -p "path/to/power/sheet1.xlsx" "path/to/power/sheet2.xlsx"
```

### Dependencies {#dependencies}

- [pandas](https://pandas.pydata.org/) for data manipulation and analysis
- [openpyxl](https://openpyxl.readthedocs.io/) for working with Excel files
- [argparse](https://docs.python.org/3/library/argparse.html) for parsing command-line arguments
- [datetime](https://docs.python.org/3/library/datetime.html) for working with dates and times

### Notes {#notes}

- Ensure that Python is installed on your system.
- The script assumes that the input files are Excel sheets (.xlsx).
- Vendor identification is based on predefined keywords in the file names.
- The 'MTTR' column is converted to numerical format ('MTTR.1') for better analysis.
- Alarms in the merged power data are verified against a predefined list in the master power sheet.
- The availability update excludes sites based on an exclusion list.
- The script updates master sheets for both network availability and commercial power alarms.
- The site data is updated with the latest availability, PAD, and PAR values.
- The consolidated output is saved to the specified output sheet (`Sep-Sites AVA,PAR & PAD.xlsx`).
---

## availabilityMaster.py README {#availabilitymasterpy-readme}

This script processes network availability data and updates a master sheet with the results. It excludes specified sites and performs analysis to determine the impact on the overall network availability.

### Table of Contents

1. [Overview](#overview)
2. [Usage](#usage)
   - [Input](#input)
   - [Output](#output)
3. [Functionality](#functionality)
   - [Excluding Sites](#excluding-sites)
   - [Analysis](#analysis)
   - [Updating Master Sheet](#updating-master-sheet)
4. [Instructions](#instructions)
5. [Dependencies](#dependencies)
6. [Notes](#notes)

### Overview {#overview}

The script reads availability data from input files, excluding specified sites, and performs an analysis to determine the impact on network availability. The final results are saved to an output Excel file.

### Usage {#usage}

#### Input {#input}

- `-a`, `--avaSheetName`: Name of the availability update Excel file.

#### Output {#output}

The processed availability data is saved to an output Excel file named `excludedAvailibilityOutsideCairo.xlsx`. The file contains multiple sheets, including the original availability data, analysis results, and a summary of the impact on each site.

### Functionality {#functionality}

#### Excluding Sites {#excluding-sites}

The script reads the availability data and excludes specific sites listed in the "Excluded List From Cells Breakdown.xlsx" file. Additionally, it handles special cases for sites with the vendor code "Z."

#### Analysis {#analysis}

The script performs an analysis to determine the impact of excluded sites on network availability. It calculates the average availability, average loss, and assigns weights to the sites based on their contribution to the total loss.

#### Updating Master Sheet {#updating-master-sheet}

The script updates the master availability sheet ("Sep-Network Availability Dashboard-2023-18 (F).xlsx") with the processed availability data.

### Instructions {#instructions}

To run the script, use the following command:

```bash
python availability_master.py -a <avaSheetName>
```

Replace `<avaSheetName>` with the name of the availability update Excel file.

### Dependencies {#dependencies}

- [pandas](https://pandas.pydata.org/) for data manipulation
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) for Excel file handling
- [argparse](https://docs.python.org/3/library/argparse.html) for command-line argument parsing

### Notes {#notes}

- Ensure that the input availability update file is correctly specified.
- The output file will be named `excludedAvailibilityOutsideCairo.xlsx`.
- Special handling is applied to sites with the vendor code "Z."

---

## masterGUI.py README {#masterguipy-readme}

This script provides a graphical user interface (GUI) to easily run multiple automation scripts. It allows you to select a directory and execute various automation scripts related to outage, availability, power, and AVA/PADPAR processing.

### Table of Contents

1. [Overview](#overview)
2. [Usage](#usage)
   - [Select Directory](#select-directory)
   - [Run Scripts](#run-scripts)
3. [Dependencies](#dependencies)
4. [Notes](#notes)

### Overview {#overview}

The script utilizes the Tkinter library to create a simple GUI that enables the user to select a directory and run multiple automation scripts. The available scripts include:

- Outage Master (`outageMaster.py`)
- Availability Master (`availabilityMaster.py`)
- Power Master (`powerMaster.py`)
- AVA/PADPAR Master (`AVAPADPARMaster.py`)

### Usage {#usage}

#### Select Directory {#select-directory}

Click the "Browse" button to select the directory where the automation scripts will be executed. The selected directory will be displayed at the top of the GUI.

#### Run Scripts {#run-scripts}

For each script listed, there is a corresponding "Run Script" button. Click the appropriate button to execute the associated automation script. The script will run in the selected directory.

### Dependencies {#dependencies}

- [tkinter](https://docs.python.org/3/library/tkinter.html) for creating the GUI
- [subprocess](https://docs.python.org/3/library/subprocess.html) for running external commands
- [threading](https://docs.python.org/3/library/threading.html) for running subprocesses concurrently
- [filedialog](https://docs.python.org/3/library/tkinter.filedialog.html) for opening file and directory selection dialogs
- [os](https://docs.python.org/3/library/os.html) for interacting with the operating system
- [subprocess](https://docs.python.org/3/library/subprocess.html) for running external commands

### Notes {#notes}

- Ensure that Python is installed on your system.
- The provided directory will be the working directory for script execution.
- The script provides feedback in the console regarding the execution of each subprocess.
- If a script is not found in the specified directory, an error message will be displayed.
- Make sure that the selected directory contains the required automation scripts.
---

## powerMaster.py README {#powermasterpy-readme}

This script is designed to merge data from multiple Power Sheets into a consolidated sheet. It extracts relevant information from each input file, adds vendor information, and appends the data to a master Power Sheet.

### Table of Contents

1. [Overview](#overview)
2. [Usage](#usage)
   - [Merge Power Sheets](#merge-power-sheets)
3. [Dependencies](#dependencies)
4. [Notes](#notes)

### Overview {#overview}

The script processes Power Sheets from different vendors (HUAWEI, NOKIA, ZTE) and consolidates the data into a master Power Sheet. It includes the following functionalities:

- **File Selection:** The script looks for Power Sheets (.xlsx files) in the current directory, excluding the output file (`combinedPowerSheet.xlsx`).
- **Vendor Identification:** The script identifies the vendor based on keywords in the file name.
- **Data Transformation:** It adds vendor information, current date, and converts the 'MTTR' column to a numerical format.
- **Alarm Verification:** It checks whether the alarms in the merged data are present in the predefined list in the master Power Sheet and marks them accordingly.
- **Output:** The merged data is saved to the output file (`combinedPowerSheet.xlsx`).

### Usage {#usage}

#### Merge Power Sheets {#merge-power-sheets}

- To merge Power Sheets, run the script without any arguments.
  
  ```bash
  python powerMaster.py
  ```

- To update the master Power Sheet with a new Power Sheet, provide the path to the new Power Sheet using the `-p` or `--powerSheetName` option.
  
  ```bash
  python powerMaster.py -p "path/to/new/PowerSheet.xlsx"
  ```

### Dependencies {#dependencies}

- [pandas](https://pandas.pydata.org/) for data manipulation and analysis
- [openpyxl](https://openpyxl.readthedocs.io/) for working with Excel files
- [argparse](https://docs.python.org/3/library/argparse.html) for parsing command-line arguments
- [datetime](https://docs.python.org/3/library/datetime.html) for working with dates and times
- [os](https://docs.python.org/3/library/os.html) for interacting with the operating system

### Notes {#notes}

- Ensure that Python is installed on your system.
- The script assumes that the input files are Power Sheets in Excel format (.xlsx).
- The predefined vendor keywords are used to identify the vendor for each Power Sheet.
- The script appends the data to the master Power Sheet (`Sep_2023 Daily Commercial Power alarm Outside-18 (F).xlsx`).
- The 'MTTR' column is converted to numerical format ('MTTR.1') for better analysis.
- Alarms in the merged data are verified against a predefined list, and the results are marked as "Yes" or "No" in the 'ALARM' column.
---

## siteMasterScript.py README {#sitemasterscriptpy-readme}

This script performs the integration of network availability data and commercial power alarm data. It merges the data from multiple sources, updates master sheets, and generates a consolidated output.

## Table of Contents

1. [Overview](#overview)
2. [Usage](#usage)
   - [Integration Process](#integration-process)
   - [Availability Update](#availability-update)
   - [Power Update](#power-update)
3. [Dependencies](#dependencies)
4. [Notes](#notes)

### Overview {#overview}

The script integrates network availability data and commercial power alarm data. It includes the following functionalities:

- **File Selection:** The script looks for availability and power sheets in the specified paths.

- **Vendor Identification:** The script identifies the vendor based on keywords in the file names.

- **Availability Update:**
  - Excludes sites based on an exclusion list.
  - Updates the master network availability sheet.

- **Power Update:**
  - Merges data from multiple power sheets.
  - Adds vendor information, current date, and converts the 'MTTR' column to a numerical format.
  - Verifies alarms against a predefined list in the master power sheet.

- **Site Data Update:**
  - Retrieves availability, power alarm duration (PAD), and power alarm count (PAR) data.
  - Updates a consolidated site data sheet with the latest values.

### Usage {#usage}

### Integration Process {#integration-process}

To perform the complete integration process (availability update, power update, and site data update):

```bash
python integration_script.py -a "path/to/availability/sheet.xlsx" -p "path/to/power/sheet1.xlsx" "path/to/power/sheet2.xlsx" -t "YYYY-MM-DD"
```

- `-a` or `--avaSheetName`: Path to the availability sheet for the update.
- `-p` or `--powerSheetNames`: Paths to the power sheets for the update. Multiple paths can be provided.
- `-t` or `--time`: Date for which the site data should be updated. If not provided, the script uses the date of the previous day.

#### Availability Update {#availability-update}

To perform only the availability update:

```bash
python integration_script.py -a "path/to/availability/sheet.xlsx"
```

#### Power Update {#power-update}

To perform only the power update:

```bash
python integration_script.py -p "path/to/power/sheet1.xlsx" "path/to/power/sheet2.xlsx"
```

### Dependencies {#dependencies}

- [pandas](https://pandas.pydata.org/) for data manipulation and analysis
- [openpyxl](https://openpyxl.readthedocs.io/) for working with Excel files
- [argparse](https://docs.python.org/3/library/argparse.html) for parsing command-line arguments
- [datetime](https://docs.python.org/3/library/datetime.html) for working with dates and times

### Notes {#notes}

- Ensure that Python is installed on your system.
- The script assumes that the input files are Excel sheets (.xlsx).
- Vendor identification is based on predefined keywords in the file names.
- The 'MTTR' column is converted to numerical format ('MTTR.1') for better analysis.
- Alarms in the merged power data are verified against a predefined list in the master power sheet.
- The availability update excludes sites based on an exclusion list.
- The script updates master sheets for both network availability and commercial power alarms.
- The site data is updated with the latest availability, PAD, and PAR values.
- The consolidated output is saved to the specified output sheet (`Sep-Sites AVA,PAR & PAD.xlsx`).
