import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime

def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0,"1")
    description_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0,"0.0")

invoice_list = []  
def add_item():
    qty = int(qty_spinbox.get())
    desc = description_entry.get()
    price = float(price_spinbox.get())
    line_total = qty*price
    invoice_item = [qty, desc, price, line_total]
    tree.insert('',0, values=invoice_item)
    clear_item()
    
    invoice_list.append(invoice_item)
    
def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

def generate_invoice():
    doc = DocxTemplate("invoicetemp.docx")
    name = first_name_entry.get()+" "+last_name_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal*(1-salestax)
    email = email_entry.get()
    address = address_entry.get()
    date = datetime.datetime.now().strftime("%Y-%m-%d")
    invoice_no = datetime.datetime.now().strftime("%Y%m%d%H%S")
    
    
    doc.render({
        "name": name,
        "phone": phone,
        "invoice_list": invoice_list,
        "subtotal": subtotal,
        "salestax": str(salestax*100)+"%",
        "total": total,
        "email": email,
        "address": address,
        "date": date,
        "invoice_no": invoice_no 
        
    })
    
    doc_name= invoice_no + name + ".docx"
    doc.save(doc_name)
    new_invoice()
    


window = tkinter.Tk()
window.title("Invoice Generator")

frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)

first_name_label = tkinter.Label(frame, text="First Name")
first_name_label.grid(row=0, column=0)
first_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1,column=0)

last_name_label = tkinter.Label(frame, text="Last Name")
last_name_label.grid(row=0, column=1)
last_name_entry = tkinter.Entry(frame)
last_name_entry.grid(row=1,column=1)

phone_label = tkinter.Label(frame, text="Phone")
phone_label.grid(row=0, column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1,column=2)

email_label = tkinter.Label(frame, text="Email")
email_label.grid(row=2, column=0)
email_entry = tkinter.Entry(frame)
email_entry.grid(row=3,column=0)

address_label = tkinter.Label(frame, text="Address")
address_label.grid(row=2, column=1, columnspan=2, sticky="news")
address_entry = tkinter.Entry(frame)
address_entry.grid(row=3,column=1, columnspan=2, sticky="news", padx=75)

qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=4, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=5, column=0)

description_label = tkinter.Label(frame, text="Description")
description_label.grid(row=4, column=1)
description_entry = tkinter.Entry(frame)
description_entry.grid(row=5,column=1)

price_label = tkinter.Label(frame, text="Qty")
price_label.grid(row=4, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=5, column=2)

add_item_button = tkinter.Button(frame, text="Add item",command = add_item)
add_item_button.grid(row=6, column=1, pady=10)

columns = ('Qty', 'Desc', 'Price', 'Total')
tree = ttk.Treeview(frame,columns=columns, show="headings")
tree.heading('Qty', text='QTY')
tree.heading('Desc', text='Description')
tree.heading('Price', text='Unit Price')
tree.heading('Total', text='Total')

tree.grid(row=7, column=0, columnspan=3, padx=20, pady=10)

save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=8, column=0,columnspan=3,sticky="news", pady=5, padx=20)

new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=9, column=0,columnspan=3,sticky="news", pady=5, padx=20)



window.mainloop()
