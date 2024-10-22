# Marble Business Cost Calculator

This is a Python-based desktop application designed to help marble businesses manage their costs, calculate profits, and keep records of each order. The application allows users to input various costs associated with marble products, calculates total costs and profits, and exports the data into an Excel file with a unique order ID, date, and time for each entry. This ensures that each order is properly classified and easy to track.

## Features

- **Graphical User Interface (GUI)**: The user-friendly interface built using `Tkinter` allows for easy data entry and management.
- **Cost Calculation**: Calculates total cost, profit, and remaining balance based on user inputs for various marble-related expenses.
- **Order Classification**: Assigns a unique Order ID, date, and time to each order, ensuring all data is classified properly.
- **Data Export**: Exports the order data into an Excel file (`marble_business_calculation.xlsx`) with structured rows for easy tracking.
- **No Data Overwrite**: Each order is appended to the existing Excel file without overwriting previous entries.

## Requirements

- **Python**: Version 3.6 or higher
- **Required Libraries**:
  - `tkinter` (for GUI creation, pre-installed with Python)
  - `openpyxl` (for creating and exporting Excel files)
  - `uuid` (for generating unique order IDs, part of Python's standard library)
  - `datetime` (for capturing current date and time, part of Python's standard library)

## Installation Guide

### Step 1: Install Python

#### For Windows:
1. Go to the official Python website: https://www.python.org/downloads/
2. Download the latest version of Python (version 3.6 or higher).
3. Run the installer and **check the box to add Python to the system PATH**.
4. Complete the installation process.

#### For macOS:
1. Open Terminal and use Homebrew (if installed) to install Python:
   ```bash
   brew install python
### Step 2: Install Dependencies
  pip install openpyxl
