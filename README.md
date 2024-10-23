# Marble Business Cost Calculator

This is a Python-based desktop application designed to help marble businesses manage their costs, calculate profits, and keep records of each order. The application allows users to input various costs associated with marble products, calculates total costs and profits, and exports the data into an Excel file with a unique order ID, date, and time for each entry. This ensures that each order is properly classified and easy to track.

## Features

- **Graphical User Interface (GUI)**: The user-friendly interface built using `Tkinter` allows for easy data entry and management.
- **Cost Calculation**: Calculates total cost, profit, and remaining balance based on user inputs for various marble-related expenses.
- **Order Classification**: Assigns a unique Order ID, date, and time to each order, ensuring all data is classified properly.
- **Data Export**: Exports the order data into an Excel file (`marble_business_calculation.xlsx`) with structured rows for easy tracking.
- **No Data Overwrite**: Each order is appended to the existing Excel file without overwriting previous entries.

## Installation

### Step 1: Install Python

To run this application, you need to have Python installed on your system.

1. **Download Python**: Go to the official Python website [python.org](https://www.python.org/downloads/) and download the latest version of Python (3.6 or higher).
2. **Install Python**:
   - **Windows/Mac**: Run the installer and follow the prompts. Ensure that you check the box that says "Add Python to PATH" during installation.
   - **Linux**: Use your package manager to install Python. For example:
     ```bash
     sudo apt update
     sudo apt install python3
     ```
3. **Verify the Installation**:
   After installation, open a terminal or command prompt and type:
   ```bash
   python --version
