from docxtpl import DocxTemplate

doc = DocxTemplate("invoice_template.docx")

invoice_list = [
    [2, "pen", 0.5, 1],
    [1, "paper pack", 5, 5],
    [2, "notebook", 2, 4]
]

subtotal = sum(item[2] * item[0] for item in invoice_list)
sales_tax_percentage = 0.12  # Assuming sales tax is 10%
sales_tax = subtotal * sales_tax_percentage
total = subtotal + sales_tax

doc.render({
    "name": "John",
    "phone": "555-55555",
    "invoice_list": invoice_list,
    "subtotal": subtotal,
    "salestax": f"{sales_tax_percentage * 100}%",  # Convert sales tax to percentage format
    "total": total
})

doc.save("new_invoice.docx")
