from docxtpl import DocxTemplate

doc = DocxTemplate("invoicetemp.docx")

invoice_list = [[2,"pen",0.5,1],
                [1,"peper",5,5],
                [2,"pencile",5,10]]


doc.render({"name":"Ehashanul", "invoice_list": invoice_list})
doc.save("new_invoice.docx")


