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

## Installation

### Step 1: Clone the Repository

First, clone this repository or download the project files.

```bash
git clone https://github.com/your-username/marble-business-cost-calculator.git
cd marble-business-cost-calculator
