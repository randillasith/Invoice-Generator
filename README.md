# Invoice Generator

This project is an **Invoice Generator** application built with **Python, CustomTkinter, Google Sheets API, and ReportLab**. It allows users to create invoices, generate PDF files, and store invoice data in Google Sheets.

## Features
- User-friendly GUI with **CustomTkinter**
- Generate invoices with multiple items
- Export invoices as **PDF** using **ReportLab**
- Save invoice details to **Google Sheets** via the **Google Sheets API**
- Dark mode interface

## Prerequisites
Ensure you have **Python 3.7+** installed on your system. 

### Required Python Packages
Install the required dependencies using pip:

```sh
pip install -r requirements.txt
```

### Setup Google Sheets API
1. Create a new project on **Google Cloud Console**.
2. Enable the **Google Sheets API** and **Google Drive API**.
3. Generate a **Service Account Key** in JSON format.
4. Share the Google Sheet with the service account email.
5. Download the **JSON key file** and update its path in `connect_to_google_sheet()`.

For a detailed guide on creating Google API credentials, visit: [Google API Credentials Guide](https://developers.google.com/workspace/guides/create-credentials)

Your project is now hosted on GitHub! 🎉

## License
This project is licensed under the MIT License.

