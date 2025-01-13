import tkinter
from tkinter import ttk, messagebox
from docxtpl import DocxTemplate
from datetime import datetime
from num2words import num2words  

si_number_counter = 0

def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    qty_spinbox.focus_set()

def add_item():
    global si_number_counter  
    si_number_counter += 1  
    qty = int(qty_spinbox.get())
    desc = desc_entry.get().title()
    price = float(price_spinbox.get()) 
    taxable_value = qty * price
    tax = int(0.12 * taxable_value)
    total_amount = round(taxable_value + tax , 2)  
    invoice_item = [si_number_counter, qty, desc, price, taxable_value, tax, total_amount]  
    invoice_item = [str(item) for item in invoice_item]
    tree.insert('', tkinter.END, values=invoice_item)
    clear_item()

def delete_last_item():
    global si_number_counter
    # Get all items in the Treeview
    items = tree.get_children()
    if items:
        # Remove the last item
        tree.delete(items[-1])
        # Decrement the SI number counter
        si_number_counter -= 1

def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    invoice_number = invoice_number_entry.get().upper()
    date = date_entry.get()
    vehicle_number = vehicle_number_entry.get().upper()
    place = place_entry.get().title()
    gstin = gstin_entry.get().upper()
    name = first_name_entry.get().title()
    address1 = address1_entry.get().title() 
    address_line_2 = address_line_2_entry.get().title()
    invoice_list = []
    for child in tree.get_children():
        item = tree.item(child)['values']
        si_number = int(item[0])
        qty = int(item[1])
        price = float(item[3])
        taxable_value = float(item[4])
        tax = int(0.12 * taxable_value)  
        total_amount = round(taxable_value + tax, 2)
        invoice_list.append([si_number, qty, item[2], price, taxable_value, tax, total_amount])
    subtotal = round(sum(item[4] for item in invoice_list), 2)
    total_taxable_value = subtotal
    salestax = round(0.06 * total_taxable_value, 2)
    total = round(subtotal + 2 * salestax) 
    total_words = num2words(total, lang='en').title() + ' Only'
    doc.render({
        "invoice_number": invoice_number,
        "date": date,
        "vehicle_number": vehicle_number,
        "place": place,
        "gstin": gstin,
        "name": name, 
        "address1": address1,
        "address_line_2": address_line_2,
        "invoice_list": invoice_list,
        "subtotal": int(subtotal),
        "salestax": int(salestax),
        "total": int(total),
        "total_words": total_words
    })  
    doc_name = invoice_number + ".docx"
    doc.save(doc_name)
    messagebox.showinfo("Invoice Complete", "Invoice Complete")
    new_invoice()

def new_invoice():
    global si_number_counter  
    si_number_counter = 0
    invoice_number_entry.delete(0, tkinter.END)
    date_entry.delete(0, tkinter.END)
    vehicle_number_entry.delete(0, tkinter.END)
    place_entry.delete(0, tkinter.END)
    gstin_entry.delete(0, tkinter.END)
    first_name_entry.delete(0, tkinter.END)
    address1_entry.delete(0, tkinter.END)
    address_line_2_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())

# GUI Setup
window = tkinter.Tk()
window.title("Bill Generator")

frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)

# Invoice Details
invoice_number_label = tkinter.Label(frame, text="Invoice Number")
invoice_number_label.grid(row=0, column=0)
invoice_number_entry = tkinter.Entry(frame)
invoice_number_entry.grid(row=1, column=0)

date_label = tkinter.Label(frame, text="Date")
date_label.grid(row=0, column=1)
date_entry = tkinter.Entry(frame)
date_entry.grid(row=1, column=1)

vehicle_number_label = tkinter.Label(frame, text="Vehicle Number")
vehicle_number_label.grid(row=0, column=2)
vehicle_number_entry = tkinter.Entry(frame)
vehicle_number_entry.grid(row=1, column=2)

place_label = tkinter.Label(frame, text="Place")
place_label.grid(row=2, column=0)
place_entry = tkinter.Entry(frame)
place_entry.grid(row=3, column=0)

gstin_label = tkinter.Label(frame, text="GSTIN")
gstin_label.grid(row=2, column=1)
gstin_entry = tkinter.Entry(frame)
gstin_entry.grid(row=3, column=1)

first_name_label = tkinter.Label(frame, text="Name of Business")
first_name_label.grid(row=4, column=0)
first_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=5, column=0)

address1_label = tkinter.Label(frame, text="Address_Line_1")
address1_label.grid(row=6, column=0)
address1_entry = tkinter.Entry(frame)
address1_entry.grid(row=7, column=0)

address_line_2_label = tkinter.Label(frame, text="Address_Line_2")
address_line_2_label.grid(row=6, column=1)
address_line_2_entry = tkinter.Entry(frame)
address_line_2_entry.grid(row=7, column=1)

# Item Details
qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=8, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=9, column=0)

desc_label = tkinter.Label(frame, text="Description")
desc_label.grid(row=8, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=9, column=1)

price_label = tkinter.Label(frame, text="Rate")
price_label.grid(row=8, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=9, column=2)

add_item_button = tkinter.Button(frame, text="Add Item", command=add_item)
add_item_button.grid(row=9, column=3, padx=20, pady=5)

# Treeview for Items
columns = ('si_number', 'qty', 'desc', 'price', 'taxable value', 'tax', 'total_amount')  
tree = ttk.Treeview(frame, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col.replace('_', ' ').title())
tree.grid(row=10, column=0, columnspan=6, padx=20, pady=10)

# Buttons
save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=12, column=0, columnspan=6, sticky="news", padx=20, pady=5)

new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=13, column=0, columnspan=6, sticky="news", padx=20, pady=5)

delete_item_button = tkinter.Button(frame, text="Delete Last Item", command=delete_last_item)
delete_item_button.grid(row=14, column=0, columnspan=6, sticky="news", padx=20, pady=5)

window.mainloop()
