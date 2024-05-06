from docxtpl import DocxTemplate
import datetime

doc = DocxTemplate("invoice_template.docx")

# Provide values for placeholders in the template
data = {
    "First name": "John",
    "Last name": "Doe",
    "Phone": "123-456-7890"
}

# Render the template with provided data
doc.render(data)

# Save the rendered document with a new filename
current_datetime = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
new_filename = f"New_invoice_{current_datetime}.docx"
doc.save(new_filename)
