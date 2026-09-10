# https://learn.microsoft.com/en-us/azure/azure-functions/add-bindings-existing-function?tabs=python-v2%2Cisolated-process%2Cnode-v4&pivots=programming-language-python

# import sys
#
# sys.path.append("./custom_packages")


import azure.functions as func
import logging
import os

# from azure.keyvault.secrets import SecretClient
# from azure.storage.blob import BlobServiceClient
from azure.identity import DefaultAzureCredential
# import base64
# import datetime
# from io import StringIO, BytesIO
# from openpyxl import load_workbook
# from openpyxl.styles import Border, Side, Font
# import pandas as pd
# import pyodbc
# from sendgrid import SendGridAPIClient
# from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

# import time

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.route(route="test")
def test(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse("OK")

# def set_secret(secret_value: str, secret_name: str):
#     """ """
#     # Read Key Vault
#     key_vault_url = os.environ["KEY_VAULT"]
#     credential = DefaultAzureCredential(
#         managed_identity_client_id="c168f556-9145-40a8-8f27-90f39a026f28"
#     )
#     client = SecretClient(vault_url=key_vault_url, credential=credential)

#     # Return secret
#     set_secret = client.set_secret(secret_name, secret_value)
#     return set_secret


# def get_secret(secret_name: str):
#     """
#     """
#     # Read Key Vault
#     key_vault_url = os.environ["KEY_VAULT"]
#     credential = DefaultAzureCredential(
#         managed_identity_client_id="c168f556-9145-40a8-8f27-90f39a026f28"
#     )
#     client = SecretClient(vault_url=key_vault_url, credential=credential)

#     # Return secret
#     retrieved_secret = client.get_secret(secret_name)
#     return retrieved_secret.value


# def read_data():
#     """

#     :return:
#     """
#     # Define your Azure SQL Database connection_db details
#     server_name = "ban-powbi-sql-01.database.windows.net"
#     database_name = "banner-platform"
#     driver = (
#         "{ODBC Driver 18 for SQL Server}"  # Use the appropriate driver version
#     )
#     sqlserver_username = get_secret("ban-powbi-sql-01-username")
#     sqlserver_password = get_secret("ban-powbi-sql-01-password")
#     connection_string = f"DRIVER={driver};SERVER={server_name};DATABASE={database_name};UID={sqlserver_username};PWD={sqlserver_password};TrustServerCertificate=yes"

#     # Generate connection
#     connection = pyodbc.connect(connection_string)

#     # Queries for price, stock, and sales
#     query_price = "SELECT * FROM clean_swi_txbannerwebplatform.PriceBySchool"
#     query_stock = "SELECT * FROM clean_swi_txbannerwebplatform.StockBySchool"
#     query_sales = "SELECT * FROM clean_swi_txbannerwebplatform.SalesBySchool"
#     query_recipients = "SELECT * FROM clean_swi_txbannerwebplatform.PurchaseOrderFormRecipients"

#     # Extract dataframes
#     df_price = pd.read_sql(query_price, connection)
#     df_stock = pd.read_sql(query_stock, connection)
#     df_sales = pd.read_sql(query_sales, connection)
#     df_recipients = pd.read_sql(query_recipients, connection)

#     # Return Dataframes
#     return df_price, df_stock, df_sales, df_recipients


# def build_excel(
#     df_price: pd.DataFrame,
#     df_stock: pd.DataFrame,
#     df_sales: pd.DataFrame,
#     recipient: pd.Series
# ):
#     """

#     :param df_price:
#     :param df_stock:
#     :param df_sales:
#     :param school_name:
#     :return:
#     """
#     school_name = recipient.SchoolName
#     account_number = recipient.AccountNumber

#     # Filter data
#     df_price_school = df_price[df_price["Customer"] == school_name].copy()
#     df_stock_school = df_stock[df_stock["Customer"] == school_name].copy()
#     df_sales_school = df_sales[df_sales["Customer"] == school_name].copy()

#     # Log shapes
#     logging.info(f"DataFrame sizes for school {school_name}: {df_price_school.shape[0]}, {df_stock_school.shape[0]}, {df_sales_school.shape[0]}")

#     # Create new sheet
#     df_po_school = df_price_school.copy()
#     df_po_school = df_po_school[["ItemId", "Name", "ColourName", "SizeName", "Price"]].copy()
#     df_po_school.columns = ["c1", "c2", "c3", "c4", "c5"]
#     df_po_school["c6"] = ""
#     df_po_school = df_po_school.sort_values(by=["c2", "c3", "c4"])
#     len_po = df_po_school.shape[0]

#     # Padding
#     df_padding = pd.DataFrame([
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#         {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
#     ])
#     df_po_school = pd.concat([df_padding, df_po_school], axis=0, ignore_index=True)
#     df_po_school["c0"] = ""
#     df_po_school = df_po_school[["c0", "c1", "c2", "c3", "c4", "c5", "c6"]]

#     # Update values
#     df_po_school.iat[10, 1] = "Item Number"
#     df_po_school.iat[10, 2] = "Name"
#     df_po_school.iat[10, 3] = "Colour"
#     df_po_school.iat[10, 4] = "Size"
#     df_po_school.iat[10, 5] = "Price"
#     df_po_school.iat[10, 6] = "Quantity"
#     df_po_school.iat[10, 1] = "Item Number"
#     df_po_school.iat[4, 1] = "Account Number"
#     df_po_school.iat[5, 1] = "Purchase Order"
#     df_po_school.iat[6, 1] = "FAO"
#     df_po_school.iat[7, 1] = "Delivery Address"

#     # Create writer
#     buffer = BytesIO()
#     # Write DataFrames to BytesIO without calling save()
#     with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
#         df_po_school.to_excel(writer, sheet_name='Form', index=False, header=False)
#         # df_price_school.to_excel(writer, sheet_name='Price', index=False)
#         df_stock_school.to_excel(writer, sheet_name='Stock', index=False)
#         df_sales_school.to_excel(writer, sheet_name='Sales', index=False)
#     buffer.seek(0)

#     # Improve formatting
#     # Load the workbook from the buffer
#     wb = load_workbook(buffer)
#     ws = wb['Form']  # Access the 'Items' sheet
#     ws.sheet_view.showGridLines = False

#     # Apply currency format to the 'Price' column (column B, starting from row 2)
#     for row in ws.iter_rows(min_row=12, min_col=6, max_col=6):  # Assuming 'Price' is in column B
#         for cell in row:
#             cell.number_format = u'£#,##0.00'  # Currency format (GBP)

#     # Define a black border style
#     black_border = Border(
#         left=Side(border_style="thin", color="000000"),
#         right=Side(border_style="thin", color="000000"),
#         top=Side(border_style="thin", color="000000"),
#         bottom=Side(border_style="thin", color="000000")
#     )

#     # Auto-adjust column width based on the maximum length of the content
#     for column in ws.columns:
#         max_length = 0
#         column_letter = column[0].column_letter  # Get the column letter
#         for cell in column:
#             try:
#                 max_length = max(max_length, len(str(cell.value)))  # Update max_length if current cell's value is longer
#             except:
#                 pass
#         # Set the column width, adding some padding
#         adjusted_width = max_length + 2  # Adding extra space for padding
#         ws.column_dimensions[column_letter].width = adjusted_width

#     # Set font bold
#     bold_font = Font(bold=True)
#     ws['B5'].font = bold_font
#     ws['B6'].font = bold_font
#     ws['B7'].font = bold_font
#     ws['B8'].font = bold_font
#     ws['B11'].font = bold_font
#     ws['C11'].font = bold_font
#     ws['D11'].font = bold_font
#     ws['E11'].font = bold_font
#     ws['F11'].font = bold_font
#     ws['G11'].font = bold_font

#     # Merge cells
#     ws.merge_cells('C5:G5')
#     ws.merge_cells('C6:G6')
#     ws.merge_cells('C7:G7')
#     ws.merge_cells('C8:G8')

#     # Apply black border to specific cells
#     for row in range(5, 12 + len_po):  # Rows 1 to 3
#         if row not in [9, 10]:
#             for col in range(2, 8):  # Columns A to C
#                 cell = ws.cell(row=row, column=col)
#                 cell.border = black_border  # Apply the border

#     # Add school name
#     ws['B2'] = f"{school_name} - SWI {datetime.datetime.now().year} Order Form"
#     ws['B2'].font = Font(underline='single', bold=True)
#     ws['C5'] = account_number

#     # Save the workbook back to the BytesIO buffer
#     buffer = BytesIO()
#     wb.save(buffer)

#     # Reset the buffer position to the start if you need to further process or write it somewhere
#     buffer.seek(0)
#     return buffer


# def send_email_with_attachment(buffer: BytesIO, recipient: pd.Series):
#     """

#     :param buffer:
#     :param recipient:
#     :return:
#     """
#     # Set your API key
#     # API Keys:
#     # sendgrid-api-key -> kiko.rullan@banner.co.uk
#     # sendgrid-api-key-Nov24 -> nasiruddin.patel@banner.co.uk
#     sendgrid_api_key = get_secret("sendgrid-api-key-Nov24")

#     # Create the email object
#     email_recipient = recipient.ContactEmail # recipient.ContactEmail
#     message = Mail(
#         from_email="nasiruddin.patel@banner.co.uk",
#         to_emails=email_recipient,
#         subject=f"Banner Order Form - {recipient.SchoolName}",
#         html_content=f'<p>Hi {recipient.ContactFirstName},</p><p>Please find attached the Stock And Sales Report. For any queries, please contact swicam@monkhouse.com. </p><p>Thank you.</p>'
#     )

#     # Encode the buffer content to base64
#     buffer_content = buffer.getvalue()
#     encoded_file = base64.b64encode(buffer_content).decode('utf-8')

#     # Create an attachment object for the Excel file
#     attachment = Attachment(
#         FileContent(encoded_file),  # Base64 encoded file content
#         FileName(f'{recipient.SchoolName}.xlsx'),  # Name of the attachment file
#         FileType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),  # MIME type for Excel
#         Disposition('attachment')  # Disposition (attached file)
#     )

#     # Add attachment to the email
#     message.attachment = attachment

#     try:
#         # Send the email
#         sg = SendGridAPIClient(sendgrid_api_key)
#         response = sg.send(message)
#         logging.info(f"Email sent! Status Code: {response.status_code}")
#     except Exception as e:
#         logging.info(f"Error sending email: {str(e)}")


# @app.function_name(name="build_order_forms")
# @app.route(route="build_order_forms")
# def funct_build_reports(req: func.HttpRequest) -> func.HttpResponse:
#     """
#     TEST FUNCTION
#     """

#     logging.info(f"UPDATED01 test function.")

#     # Read data
#     df_price, df_stock, df_sales, df_recipients = read_data()

#     # Loop over recipients
#     # recipient.SchoolName == "The Co-Operative Academy of Manchester"
#     for index, recipient in df_recipients.iterrows():
#         if index < 1000:
#             logging.info(recipient)
#             # Build excel
#             buffer = build_excel(df_price=df_price, df_stock=df_stock, df_sales=df_sales, recipient=recipient)
#             # Send email
#             send_email_with_attachment(buffer=buffer, recipient=recipient)

#     # Connect
#     # connection_string = # get_secret("bannerdwstorage_connection_string")
#     # blob_service_client = BlobServiceClient.from_connection_string(connection_string)
#     # container = "school-report"
#     # blob_client = blob_service_client.get_blob_client(container=container, blob="SchoolReportTemplate.xlsx")
#     # blob_client.upload_blob(buffer, overwrite=True)

#     return func.HttpResponse("BUILT ORDER FORM ran successfully!", status_code=200)
