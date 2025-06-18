import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import PhotoImage
from docxtpl import DocxTemplate
import datetime
import os
import sys

# Initialize the main window
window = tk.Tk()
window.title("MUSAFIR STUDIOS")
window.geometry("1000x900")


# # Set the window icon
# company_logo = PhotoImage(file='logo.png')
# window.iconphoto(True, company_logo)

# Initialize global variables
productList = []
global_t = 0  # Cumulative subtotal

def get_resource_path(relative_path):
    """ Get the absolute path to a resource file. """
    if getattr(sys, 'frozen', False):  # Check if the app is frozen
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def generateInvoice():
    f = for_var.get()
    to = to_var.get()
    d = date_var.get()
    ta = float(tax_var.get())
    tax = global_t * (ta / 100)
    grand_total = tax + global_t
    
    template_path = get_resource_path("template.docx")
    doc = DocxTemplate(template_path)
    doc.render({
        "for": f,
        "to": to,
        "date": d,
        "tax": ta,
        "itemList": productList,
        "gtotal": grand_total
    })
    
    file_name = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        initialfile=f + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    )

    if file_name:  # If the user didn't cancel the dialog
        doc.save(file_name)
        print("Invoice generated and saved at:", file_name)
    print("Invoice generated and saved.")
    reset()

def addService():
    global global_t
    s = service_var.get()
    q = quantity_var.get()
    p = float(price_var.get())
    item_total = q * p
    
    # Update global_t to reflect the cumulative subtotal
    global_t += item_total
    
    # Insert into table and update product list
    table.insert("", 0, values=(s, q, p, item_total))
    productList.append([s, q, p, item_total])
    
    # Clear input fields
    clearServiceField()
    
    # Update totals
    updateTotals()

def clearServiceField():
    service_var.set("")
    quantity_var.set(1)
    price_var.set('0')
    tax_var.set('0')

def reset():
    clearServiceField()
    for_var.set("")
    to_var.set("")
    date_var.set("")
    table.delete(*table.get_children())
    productList.clear()

def updateTotals():
    """Calculate tax and grand total based on current subtotal."""
    try:
        ta = float(tax_var.get())
        tax = global_t * (ta / 100)
        grand_total = tax + global_t
        
        # Update display values (print to console here)
        print(f"Subtotal: {global_t:.2f}")
        print(f"Tax: {tax:.2f}")
        print(f"Grand Total: {grand_total:.2f}")
        
    except ValueError:
        print("Invalid tax value")

# Focus event handlers for price, quantity, and tax
def Pon_focus_in(event):
    """Clear the entry if it contains the default integer value '0'."""
    entry = event.widget
    if entry.get() == '0':
        entry.delete(0, tk.END)

def Pon_focus_out(event):
    """Restore the default integer value '0' if the entry is empty when it loses focus."""
    entry = event.widget
    if entry.get() == '':
        entry.insert(0, '0')

def Qon_focus_in(event):
    """Clear the entry if it contains the default integer value '1'."""
    entry = event.widget
    if entry.get() == '1':
        entry.delete(0, tk.END)

def Qon_focus_out(event):
    """Restore the default integer value '1' if the entry is empty when it loses focus."""
    entry = event.widget
    if entry.get() == '':
        entry.insert(0, '1')

def Ton_focus_in(event):
    """Clear the entry if it contains the default integer value '0'."""
    entry = event.widget
    if entry.get() == '0':
        entry.delete(0, tk.END)

def Ton_focus_out(event):
    """Restore the default integer value '0' if the entry is empty when it loses focus."""
    entry = event.widget
    if entry.get() == '':
        entry.insert(0, '0')

# UI Components
for_label = ttk.Label(window, text="Quotation For")
for_label.grid(row=0, column=0)

for_var = tk.StringVar()
for_entry = ttk.Entry(window, textvariable=for_var)
for_entry.grid(row=1, column=0)

to_label = ttk.Label(window, text="Invoice to")
to_label.grid(row=0, column=1)

to_var = tk.StringVar()
to_entry = ttk.Entry(window, textvariable=to_var)
to_entry.grid(row=1, column=1)

date_label = ttk.Label(window, text="Date")
date_label.grid(row=0, column=2)

date_var = tk.StringVar()
date_entry = ttk.Entry(window, textvariable=date_var)
date_entry.grid(row=1, column=2)

service_label = ttk.Label(window, text="Service Description")
service_label.grid(row=3, column=0)

service_var = tk.StringVar()
service_entry = ttk.Entry(window, textvariable=service_var)
service_entry.grid(row=4, column=0)

quantity_label = ttk.Label(window, text="Quantity")
quantity_label.grid(row=3, column=1)

quantity_var = tk.IntVar(value=1)
quantity_entry = ttk.Entry(window, textvariable=quantity_var)
quantity_entry.grid(row=4, column=1)
quantity_entry.bind("<FocusIn>", Qon_focus_in)
quantity_entry.bind("<FocusOut>", Qon_focus_out)

price_label = ttk.Label(window, text="Price")
price_label.grid(row=3, column=2)

price_var = tk.StringVar(value='0')
price_entry = ttk.Entry(window, textvariable=price_var)
price_entry.grid(row=4, column=2)
price_entry.bind("<FocusIn>", Pon_focus_in)
price_entry.bind("<FocusOut>", Pon_focus_out)

tax_label = ttk.Label(window, text="Tax (%)")
tax_label.grid(row=5, column=1)

tax_var = tk.StringVar(value='0')
tax_entry = ttk.Entry(window, textvariable=tax_var)
tax_entry.grid(row=6, column=1)
tax_entry.bind("<FocusIn>", Ton_focus_in)
tax_entry.bind("<FocusOut>", Ton_focus_out)

add_button = ttk.Button(window, text="Add Service", command=addService)
add_button.grid(row=7, column=0, columnspan=3, sticky="ewns")

# Table for displaying services
table = ttk.Treeview(window, columns=("service description", "quantity", "price", "total"), show="headings")
table.grid(row=8, columnspan=3, sticky="ewns", padx=20)
table.heading("service description", text="Service Description", anchor="w")
table.heading("quantity", text="Quantity", anchor="w")
table.heading("price", text="Price", anchor="w")
table.heading("total", text="Total", anchor="w")

# New and Generate buttons
new_button = ttk.Button(window, text="New Invoice", command=reset)
new_button.grid(row=9, column=1)

generate_button = ttk.Button(window, text="Generate Invoice", command=generateInvoice)
generate_button.grid(row=10, column=1)

# Set up row and column configurations
for i in range(11):
    window.rowconfigure(i, weight=1 if i < 8 else 6)

for i in range(3):
    window.columnconfigure(i, weight=1)

window.mainloop()
