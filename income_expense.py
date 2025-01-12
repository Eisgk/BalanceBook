import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from openpyxl import Workbook
from datetime import datetime

# Initialize main window
root = tk.Tk()
root.title("Income and Expense Tracker")
root.geometry("600x400")  # Set window size

# Define a dictionary to store data for each month
monthly_data = {}

# Function to add income and expenses
def add_entry():
    try:
        # Get user inputs
        date_str = date_entry.get()
        category = category_entry.get()
        income = float(income_entry.get()) if income_entry.get() else 0.0
        expense = float(expense_entry.get()) if expense_entry.get() else 0.0
        
        # Validate the date format
        try:
            date = datetime.strptime(date_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter the date in dd/mm/yyyy format.")
            return
        
        # Format the date for display
        formatted_date = date.strftime("%d/%m/%Y")
        month_str = date.strftime("%Y-%m")  # Key for monthly data

        # Calculate the balance for this entry
        if month_str not in monthly_data:
            monthly_data[month_str] = []
        
        if monthly_data[month_str]:
            last_balance = monthly_data[month_str][-1]['Balance']
        else:
            last_balance = 0.0
        
        balance = last_balance + income - expense

        # Create the entry and add it to the list for the month
        entry = {
            'Date': formatted_date,
            'Category': category,
            'Income': income,
            'Expense': expense,
            'Balance': balance
        }
        monthly_data[month_str].append(entry)
        
        # Update the table with the new entry
        update_table(month_str)
        
        messagebox.showinfo("Success", f"Entry added for {formatted_date}!")
        clear_fields()
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter valid numbers for income and expenses.")

# Function to update the table with data for a specific month
def update_table(month_str):
    # Clear the current table
    for row in tree.get_children():
        tree.delete(row)
    
    # Insert new rows into the table for the selected month
    for entry in monthly_data[month_str]:
        tree.insert("", "end", values=(entry['Date'], entry['Category'], entry['Income'], entry['Expense'], entry['Balance']))

# Function to export the total yearly data to Excel
def export_yearly_to_excel():
    wb = Workbook()  # Create a new Excel workbook
    
    # Sort months in ascending order
    months_sorted = sorted(monthly_data.keys())

    # Loop through each month to create a separate sheet for each
    for idx, month_str in enumerate(months_sorted):
        ws = wb.create_sheet(title=month_str)  # Create a new sheet for each month
        
        # Write custom header
        ws.append([f"รายรับรายจ่ายเดือน{get_month_name(month_str)}"])
        ws.append(["Month", "Category", "Income", "Expense", "Balance"])

        # Add balance forward (first row for the month)
        if idx == 0:
            # For the first month, start balance is 0.0
            balance_forward = 0.0
        else:
            # For subsequent months, carry forward the balance from the previous month
            prev_month_str = months_sorted[idx - 1]
            balance_forward = monthly_data[prev_month_str][-1]['Balance']
        
        # Insert the 'Bring forward' row with the carried over balance
        ws.append(["Bring forward", "", "", "", balance_forward])

        # Add the data for each entry in the month
        for entry in monthly_data[month_str]:
            ws.append([entry['Date'], entry['Category'], entry['Income'], entry['Expense'], entry['Balance']])

    # Remove the default empty sheet created by Workbook
    del wb['Sheet']
    
    # Save the file
    wb.save("yearly_income_expense_report.xlsx")
    messagebox.showinfo("Exported", "Yearly data exported to yearly_income_expense_report.xlsx")

# Helper function to get Thai month name
def get_month_name(month_str):
    month_names = {
        "01": "มกราคม",
        "02": "กุมภาพันธ์",
        "03": "มีนาคม",
        "04": "เมษายน",
        "05": "พฤษภาคม",
        "06": "มิถุนายน",
        "07": "กรกฎาคม",
        "08": "สิงหาคม",
        "09": "กันยายน",
        "10": "ตุลาคม",
        "11": "พฤศจิกายน",
        "12": "ธันวาคม"
    }
    return month_names[month_str[5:7]]

# Function to clear input fields
def clear_fields():
    date_entry.delete(0, tk.END)
    category_entry.delete(0, tk.END)
    income_entry.delete(0, tk.END)
    expense_entry.delete(0, tk.END)

# GUI layout
# Entry for Date
date_label = tk.Label(root, text="Enter Date (dd/mm/yyyy):")
date_label.pack()
date_entry = tk.Entry(root)
date_entry.pack()

# Entry for Category/Description
category_label = tk.Label(root, text="Enter Category (e.g., Salary, Buy Cake):")
category_label.pack()
category_entry = tk.Entry(root)
category_entry.pack()

# Entry for Income
income_label = tk.Label(root, text="Enter Income:")
income_label.pack()
income_entry = tk.Entry(root)
income_entry.pack()

# Entry for Expenses
expense_label = tk.Label(root, text="Enter Expense:")
expense_label.pack()
expense_entry = tk.Entry(root)
expense_entry.pack()

# Buttons for adding entries and exporting
add_button = tk.Button(root, text="Add Entry", command=add_entry)
add_button.pack()

export_excel_button = tk.Button(root, text="Export Yearly Data to Excel", command=export_yearly_to_excel)
export_excel_button.pack()

# Table for displaying entries
columns = ("Date", "Category", "Income", "Expense", "Balance")
tree = ttk.Treeview(root, columns=columns, show="headings")
tree.heading("Date", text="Date")
tree.heading("Category", text="Category")
tree.heading("Income", text="Income")
tree.heading("Expense", text="Expense")
tree.heading("Balance", text="Balance")
tree.pack(fill=tk.BOTH, expand=True)

# Start the Tkinter event loop
root.mainloop()
