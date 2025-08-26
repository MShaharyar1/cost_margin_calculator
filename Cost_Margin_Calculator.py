import tkinter as tk
from tkinter import messagebox
import openpyxl
import os
import uuid
from datetime import datetime


def _to_float(value):
    """Safely convert a string to float.

    Returns 0.0 when conversion fails so that export_to_excel writes
    numeric values instead of text with currency symbols.
    """
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _label_amount(label_widget):
    """Extract the numeric amount from a label widget."""
    text = label_widget.cget("text")
    # Labels are of the form "Description: PKR <amount>"
    # We only need the numeric part for Excel.
    return _to_float(text.split()[-1])


# Function to calculate total cost and profit for the marble business
def calculate_cost():
    try:
        # Retrieve input values from the user
        qabar_marble = float(entry_qabar_marble.get())
        thin_qabar_marble = float(entry_thin_qabar_marble.get())
        style = float(entry_style.get())
        liner = float(entry_liner.get())
        steel = float(entry_steel.get())
        labor_cost = float(entry_labor.get())
        transportation_cost = float(entry_transportation.get())
        writing_cost = float(entry_writing.get())
        advance_payment = float(entry_advance.get())
        profit_margin_percentage = float(entry_profit_margin.get())

        # Calculate total cost
        total_cost = (qabar_marble + thin_qabar_marble + style + liner + steel +
                      labor_cost + transportation_cost + writing_cost)
        
        # Calculate profit and final price
        profit = (profit_margin_percentage / 100) * total_cost
        final_price = total_cost + profit
        remaining_balance = final_price - advance_payment
        
        # Display the calculated values in the labels
        label_total_cost.config(text=f"Total Cost: PKR {total_cost:.2f}")
        label_profit.config(text=f"Profit: PKR {profit:.2f}")
        label_final_price.config(text=f"Total Price (with profit): PKR {final_price:.2f}")
        label_remaining_balance.config(text=f"Remaining Balance: PKR {remaining_balance:.2f}")
    
    except ValueError:
        messagebox.showerror("Invalid input", "Please enter valid numeric values.")

# Function to export data to Excel in the required format
def export_to_excel():
    try:
        # Generate unique order ID and capture the current date and time
        order_id = str(uuid.uuid4())  # Generate a unique ID for the order
        current_datetime = datetime.now()
        current_date = current_datetime.strftime("%Y-%m-%d")
        current_time = current_datetime.strftime("%H:%M:%S")
        
        # Check if the Excel file already exists
        file_name = "Rizwan_Marble_Cost.xlsx"
        if os.path.exists(file_name):
            # Open the existing workbook
            wb = openpyxl.load_workbook(file_name)
            ws = wb.active
        else:
            # Create a new workbook and add a sheet if the file doesn't exist
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Rizwan_Marble_Cost"
            
            # Write headers in the new Excel sheet
            ws.append(["Order ID", "Date", "Time", "Description", "Cost (PKR)"])
        
        # Prepare the cost breakdown data
        data = [
            ("Qabar Marble", _to_float(entry_qabar_marble.get())),
            ("Thin Qabar Marble", _to_float(entry_thin_qabar_marble.get())),
            ("Style", _to_float(entry_style.get())),
            ("Liner", _to_float(entry_liner.get())),
            ("Steel", _to_float(entry_steel.get())),
            ("Labor (Mazdoori)", _to_float(entry_labor.get())),
            ("Transportation", _to_float(entry_transportation.get())),
            ("Writing (Likhai)", _to_float(entry_writing.get())),
            ("Total Cost", _label_amount(label_total_cost)),
            ("Profit", _label_amount(label_profit)),
            ("Total Price (with profit)", _label_amount(label_final_price)),
            ("Advance Payment", _to_float(entry_advance.get())),
            ("Remaining Balance", _label_amount(label_remaining_balance)),
        ]
        
        # Append Order ID, Date, and Time in the first row
        ws.append([order_id, current_date, current_time, "", ""])
        
        # Append cost breakdown in subsequent rows
        for item in data:
            ws.append(["", "", "", item[0], item[1]])

        # Save the workbook
        wb.save(file_name)
        messagebox.showinfo("Success", f"Data exported successfully to {file_name}")
    
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to set up the user interface (UI)
def setup_ui(root):
    # Labels and entry fields for user input
    tk.Label(root, text="Cost of Qabar Marble:").grid(row=0, column=0, padx=10, pady=5)
    tk.Label(root, text="Cost of Thin Qabar Marble:").grid(row=1, column=0, padx=10, pady=5)
    tk.Label(root, text="Cost of Style:").grid(row=2, column=0, padx=10, pady=5)
    tk.Label(root, text="Cost of Liner:").grid(row=3, column=0, padx=10, pady=5)
    tk.Label(root, text="Cost of Steel:").grid(row=4, column=0, padx=10, pady=5)
    tk.Label(root, text="Labor Cost (Mazdoori):").grid(row=5, column=0, padx=10, pady=5)
    tk.Label(root, text="Transportation Cost:").grid(row=6, column=0, padx=10, pady=5)
    tk.Label(root, text="Writing (Likhai) Cost:").grid(row=7, column=0, padx=10, pady=5)
    tk.Label(root, text="Profit Margin (%):").grid(row=8, column=0, padx=10, pady=5)
    tk.Label(root, text="Advance Payment:").grid(row=9, column=0, padx=10, pady=5)

    # Entry fields to take inputs
    global entry_qabar_marble, entry_thin_qabar_marble, entry_style, entry_liner
    global entry_steel, entry_labor, entry_transportation, entry_writing
    global entry_profit_margin, entry_advance

    entry_qabar_marble = tk.Entry(root)
    entry_qabar_marble.grid(row=0, column=1)

    entry_thin_qabar_marble = tk.Entry(root)
    entry_thin_qabar_marble.grid(row=1, column=1)

    entry_style = tk.Entry(root)
    entry_style.grid(row=2, column=1)

    entry_liner = tk.Entry(root)
    entry_liner.grid(row=3, column=1)

    entry_steel = tk.Entry(root)
    entry_steel.grid(row=4, column=1)

    entry_labor = tk.Entry(root)
    entry_labor.grid(row=5, column=1)

    entry_transportation = tk.Entry(root)
    entry_transportation.grid(row=6, column=1)

    entry_writing = tk.Entry(root)
    entry_writing.grid(row=7, column=1)

    entry_profit_margin = tk.Entry(root)
    entry_profit_margin.grid(row=8, column=1)

    entry_advance = tk.Entry(root)
    entry_advance.grid(row=9, column=1)

    # Buttons for calculations and export
    calculate_button = tk.Button(root, text="Calculate", command=calculate_cost)
    calculate_button.grid(row=10, column=0, columnspan=2, pady=10)

    export_button = tk.Button(root, text="Export to Excel", command=export_to_excel)
    export_button.grid(row=11, column=0, columnspan=2, pady=10)

    # Labels to display output (total cost, profit, etc.)
    global label_total_cost, label_profit, label_final_price, label_remaining_balance

    label_total_cost = tk.Label(root, text="Total Cost: PKR 0.00")
    label_total_cost.grid(row=12, column=0, columnspan=2)

    label_profit = tk.Label(root, text="Profit: PKR 0.00")
    label_profit.grid(row=13, column=0, columnspan=2)

    label_final_price = tk.Label(root, text="Total Price (with profit): PKR 0.00")
    label_final_price.grid(row=14, column=0, columnspan=2)

    label_remaining_balance = tk.Label(root, text="Remaining Balance: PKR 0.00")
    label_remaining_balance.grid(row=15, column=0, columnspan=2)

# Main function to run the application
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Marble Business Cost Calculator")
    setup_ui(root)
    root.mainloop()
