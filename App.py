import customtkinter as ctk
from tkinter import messagebox
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def test_google_sheet():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("Path to Credentials", scope)
        client = gspread.authorize(creds)
        
        # List all available spreadsheets

        sheets = client.openall()
        print("Available Sheets:")
        for sheet in sheets:
            print(sheet.title)
        
        # Try opening the "Invoices" sheet
        sheet = client.open("Invoices").sheet1
        print("Successfully connected to 'Invoices'!")
    except gspread.exceptions.SpreadsheetNotFound:
        print("❌ ERROR: Google Sheet 'Invoices' NOT FOUND. Check the name and permissions.")
    except Exception as e:
        print(f"❌ ERROR: {e}")

# Run the test function
test_google_sheet()

# Function to test Google Sheets connection
def test_google_sheet_connection():
    try:
        sheet = connect_to_google_sheet("Invoices")  # Change this to match your Google Sheet name
        sheet_title = sheet.title
        messagebox.showinfo("Success", f"✅ Connected to Google Sheets!\nSheet Name: {sheet_title}")
    except gspread.exceptions.SpreadsheetNotFound:
        messagebox.showerror("Error", "❌ Google Sheet 'Invoices' NOT FOUND. Check the name and permissions.")
    except Exception as e:
        messagebox.showerror("Error", f"❌ ERROR: {e}")




# Google Sheets Authentication
def connect_to_google_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("Path to Credentials", scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1  # Open the first sheet
    return sheet

# Function to generate invoice PDF and send data to Google Sheets
def generate_invoice():
    order_form_number = entry_order_form.get()
    invoice_number = entry_invoice.get()
    customer_name = entry_name.get()
    company = entry_company.get()
    phone = entry_phone.get()
    
    items = []
    for row in item_rows:
        description = row[0].get()
        if not description:
            continue
        qty = int(row[1].get() or 0)
        price = int(row[2].get() or 0)
        discount = int(row[3].get() or 0)
        item_total = (qty * price) - discount
        items.append((description, qty, price, discount, item_total))
    
    total = sum(item[4] for item in items)
    bg_image_path = "/Your Path/inv.png"  # Background image path (Update this)
    output_file = f"invoice_{invoice_number}.pdf"
    
    # Generate PDF Invoice
    c = canvas.Canvas(output_file, pagesize=A4)
    width, height = A4
    
    try:
        bg_image = ImageReader(bg_image_path)
        c.drawImage(bg_image, 0, 0, width=width, height=height)
    except:
        print("Background image not found, proceeding without it.")
    
    # Invoice Header
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 200, f"ORDER FORM NUMBER: {order_form_number}")
    c.drawString(400, height - 200, f"INVOICE NUMBER: {invoice_number}")
    
    # Customer Details
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 230, "NAME:")
    c.setFont("Helvetica", 12)
    c.drawString(100, height - 230, customer_name)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 250, "COMPANY:")
    c.setFont("Helvetica", 12)
    c.drawString(120, height - 250, company)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 270, "PHONE NO:")
    c.setFont("Helvetica", 12)
    c.drawString(120, height - 270, phone)
    
    # Table Header
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.black)
    c.rect(50, height - 310, 500, 20, fill=1, stroke=1)
    c.setFillColor(colors.white)
    c.drawString(55, height - 305, "NO.")
    c.drawString(100, height - 305, "ITEM DESCRIPTION")
    c.drawString(300, height - 305, "QTY")
    c.drawString(350, height - 305, "PRICE")
    c.drawString(420, height - 305, "DISCOUNT")
    c.drawString(500, height - 305, "TOTAL")
    
    # Table Content
    y_position = height - 330
    c.setFont("Helvetica", 12)
    c.setFillColor(colors.black)
    for index, item in enumerate(items, start=1):
        description, qty, price, discount, item_total = item
        c.drawString(55, y_position, str(index))
        c.drawString(100, y_position, description)
        c.drawString(300, y_position, str(qty))
        c.drawString(350, y_position, f"{price:.2f}")
        c.drawString(420, y_position, f"{discount:.2f}")
        c.drawString(500, y_position, f"{item_total:.2f}")
        y_position -= 20
    
    # Total Amount
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.darkred)
    c.drawString(400, y_position - 20, "TOTAL: ")
    c.drawString(500, y_position - 20, f"{total:.2f}/=")
    
    # Footer Message
    c.setFont("Helvetica-Oblique", 10)
    c.setFillColor(colors.black)
    c.drawString(20,140, "This is a computer-generated invoice and does not require a signature.")

    c.save()


    # Send data to Google Sheet
    sheet = connect_to_google_sheet("Invoices")  # Update with your Google Sheet name

    
    for item in items:
        description, qty, price, discount, item_total = item
        sheet.append_row([order_form_number, invoice_number, customer_name, company, phone, description, qty, price, discount, item_total])
    
    sheet.append_row(["", "", "", "", "", "", "", "", "",])

    messagebox.showinfo("Success", f"Invoice {invoice_number} generated and sent to Google Sheets!")


# Initialize the app
ctk.set_appearance_mode("dark")  # Dark mode
ctk.set_default_color_theme("blue")  # Theme color
root = ctk.CTk()
root.title("Invoice Generator")
root.geometry("500x600")
root.resizable(False, False)

def add_item_row():
    row_frame = ctk.CTkFrame(item_frame)
    row_frame.pack(fill='x', pady=5)
    
    desc_entry = ctk.CTkEntry(row_frame, width=180, placeholder_text="Description")
    desc_entry.pack(side='left', padx=5)
    qty_entry = ctk.CTkEntry(row_frame, width=80, placeholder_text="Qty")
    qty_entry.pack(side='left', padx=5)
    price_entry = ctk.CTkEntry(row_frame, width=80, placeholder_text="Price")
    price_entry.pack(side='left', padx=5)
    discount_entry = ctk.CTkEntry(row_frame, width=80, placeholder_text="Discount")
    discount_entry.pack(side='left', padx=5)
    
    item_rows.append((desc_entry, qty_entry, price_entry, discount_entry))
def new_invoice():
    for entry in [entry_order_form, entry_invoice, entry_name, entry_company, entry_phone]:
        entry.delete(0, tk.END)
    for row in item_rows:
        for field in row:
            field.destroy()
    item_rows.clear()
    add_item_row()


# Main Frame
main_frame = ctk.CTkFrame(root)
main_frame.pack(pady=10, padx=20, fill='both', expand=True)

ctk.CTkLabel(main_frame, text="Order Form Number:").pack(anchor='w')
entry_order_form = ctk.CTkEntry(main_frame, width=300)
entry_order_form.pack()

ctk.CTkLabel(main_frame, text="Invoice Number:").pack(anchor='w')
entry_invoice = ctk.CTkEntry(main_frame, width=300)
entry_invoice.pack()

ctk.CTkLabel(main_frame, text="Customer Name:").pack(anchor='w')
entry_name = ctk.CTkEntry(main_frame, width=300)
entry_name.pack()

ctk.CTkLabel(main_frame, text="Company:").pack(anchor='w')
entry_company = ctk.CTkEntry(main_frame, width=300)
entry_company.pack()

ctk.CTkLabel(main_frame, text="Phone Number:").pack(anchor='w')
entry_phone = ctk.CTkEntry(main_frame, width=300)
entry_phone.pack()

# Item Section
item_frame = ctk.CTkFrame(main_frame)
item_frame.pack(pady=10, fill='x')

item_rows = []
add_item_row()

# Buttons
btn_frame = ctk.CTkFrame(main_frame)
btn_frame.pack(pady=10)

ctk.CTkButton(btn_frame, text="Add Item", command=add_item_row).pack(side='left', padx=5)
ctk.CTkButton(btn_frame, text="Generate Invoice", command=generate_invoice).pack(side='left', padx=5)
ctk.CTkButton(btn_frame, text="New Invoice", command=new_invoice).pack(side='left', padx=5)

btn_frame = ctk.CTkFrame(main_frame)
btn_frame.pack(pady=10)
ctk.CTkButton(btn_frame, text="Test Google Sheets", command=test_google_sheet_connection).pack(side='left', padx=5)


root.mainloop()