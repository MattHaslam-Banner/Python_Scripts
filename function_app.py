# https://learn.microsoft.com/en-us/azure/azure-functions/add-bindings-existing-function?tabs=python-v2%2Cisolated-process%2Cnode-v4&pivots=programming-language-python

# import sys
#
# sys.path.append("./custom_packages")

from azure.identity import DefaultAzureCredential
import azure.functions as func
from azure.keyvault.secrets import SecretClient
from azure.storage.blob import BlobServiceClient
import base64
import datetime
from io import StringIO, BytesIO
import logging
import os
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Font
import pandas as pd
import pyodbc
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions
from sqlalchemy import create_engine, text
import urllib
import json
import csv
import io
import logging
import tempfile

from openai import AzureOpenAI

# import time
app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

def set_secret(secret_value: str, secret_name: str):
    """ """
    # Read Key Vault
    key_vault_url = os.environ["KEY_VAULT"]
    credential = DefaultAzureCredential(
        managed_identity_client_id="c168f556-9145-40a8-8f27-90f39a026f28"
    )
    client = SecretClient(vault_url=key_vault_url, credential=credential)

    # Return secret
    set_secret = client.set_secret(secret_name, secret_value)
    return set_secret


def get_secret(secret_name: str):
    """
    """
    # Read Key Vault
    key_vault_url = os.environ["KEY_VAULT"]
    credential = DefaultAzureCredential(
        managed_identity_client_id="c168f556-9145-40a8-8f27-90f39a026f28"
    )
    client = SecretClient(vault_url=key_vault_url, credential=credential)

    # Return secret
    retrieved_secret = client.get_secret(secret_name)
    return retrieved_secret.value


def read_data():
    """

    :return:
    """
    # Define your Azure SQL Database connection_db details
    server_name = "ban-powbi-sql-01.database.windows.net"
    database_name = "banner-platform"
    driver = (
        "{ODBC Driver 18 for SQL Server}"  # Use the appropriate driver version
    )
    sqlserver_username = get_secret("ban-powbi-sql-01-username")
    sqlserver_password = get_secret("ban-powbi-sql-01-password")
    connection_string = f"DRIVER={driver};SERVER={server_name};DATABASE={database_name};UID={sqlserver_username};PWD={sqlserver_password};TrustServerCertificate=yes"

    # Generate connection
    connection = pyodbc.connect(connection_string)

    # Queries for price, stock, and sales
    query_price = "SELECT * FROM clean_swi_txbannerwebplatform.PriceBySchool"
    query_stock = "SELECT * FROM clean_swi_txbannerwebplatform.StockBySchool"
    query_sales = "SELECT * FROM clean_swi_txbannerwebplatform.SalesBySchool"
    query_recipients = "SELECT * FROM clean_swi_txbannerwebplatform.PurchaseOrderFormRecipients"

    # Extract dataframes
    df_price = pd.read_sql(query_price, connection)
    df_stock = pd.read_sql(query_stock, connection)
    df_sales = pd.read_sql(query_sales, connection)
    df_recipients = pd.read_sql(query_recipients, connection)

    # Return Dataframes
    return df_price, df_stock, df_sales, df_recipients


def build_excel(
    df_price: pd.DataFrame,
    df_stock: pd.DataFrame,
    df_sales: pd.DataFrame,
    recipient: pd.Series
):
    """

    :param df_price:
    :param df_stock:
    :param df_sales:
    :param school_name:
    :return:
    """
    school_name = recipient.SchoolName
    account_number = recipient.AccountNumber

    # Filter data
    df_price_school = df_price[df_price["Customer"] == school_name].copy()
    df_stock_school = df_stock[df_stock["Customer"] == school_name].copy()
    df_sales_school = df_sales[df_sales["Customer"] == school_name].copy()

    # Create new sheet
    df_po_school = df_price_school.copy()
    df_po_school = df_po_school[["ItemId", "Name", "ColourName", "SizeName", "Price"]].copy()
    df_po_school.columns = ["c1", "c2", "c3", "c4", "c5"]
    df_po_school["c6"] = ""
    df_po_school = df_po_school.sort_values(by=["c2", "c3", "c4"])
    len_po = df_po_school.shape[0]

    # Padding
    df_padding = pd.DataFrame([
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
        {'c1': '', 'c2': '', 'c3': '', 'c4': '', 'c5': '', 'c6': ''},
    ])
    df_po_school = pd.concat([df_padding, df_po_school], axis=0, ignore_index=True)
    df_po_school["c0"] = ""
    df_po_school = df_po_school[["c0", "c1", "c2", "c3", "c4", "c5", "c6"]]

    # Update values
    df_po_school.iat[10, 1] = "Item Number"
    df_po_school.iat[10, 2] = "Name"
    df_po_school.iat[10, 3] = "Colour"
    df_po_school.iat[10, 4] = "Size"
    df_po_school.iat[10, 5] = "Price"
    df_po_school.iat[10, 6] = "Quantity"
    df_po_school.iat[10, 1] = "Item Number"
    df_po_school.iat[4, 1] = "Account Number"
    df_po_school.iat[5, 1] = "Purchase Order"
    df_po_school.iat[6, 1] = "FAO"
    df_po_school.iat[7, 1] = "Delivery Address"

    # Create writer
    buffer = BytesIO()
    # Write DataFrames to BytesIO without calling save()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_po_school.to_excel(writer, sheet_name='Form', index=False, header=False)
        # df_price_school.to_excel(writer, sheet_name='Price', index=False)
        df_stock_school.to_excel(writer, sheet_name='Stock', index=False)
        df_sales_school.to_excel(writer, sheet_name='Sales', index=False)
    buffer.seek(0)

    # Improve formatting
    # Load the workbook from the buffer
    wb = load_workbook(buffer)
    ws = wb['Form']  # Access the 'Items' sheet
    ws.sheet_view.showGridLines = False

    # Apply currency format to the 'Price' column (column B, starting from row 2)
    for row in ws.iter_rows(min_row=12, min_col=6, max_col=6):  # Assuming 'Price' is in column B
        for cell in row:
            cell.number_format = u'£#,##0.00'  # Currency format (GBP)

    # Define a black border style
    black_border = Border(
        left=Side(border_style="thin", color="000000"),
        right=Side(border_style="thin", color="000000"),
        top=Side(border_style="thin", color="000000"),
        bottom=Side(border_style="thin", color="000000")
    )

    # Auto-adjust column width based on the maximum length of the content
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get the column letter
        for cell in column:
            try:
                max_length = max(max_length, len(str(cell.value)))  # Update max_length if current cell's value is longer
            except:
                pass
        # Set the column width, adding some padding
        adjusted_width = max_length + 2  # Adding extra space for padding
        ws.column_dimensions[column_letter].width = adjusted_width

    # Set font bold
    bold_font = Font(bold=True)
    ws['B5'].font = bold_font
    ws['B6'].font = bold_font
    ws['B7'].font = bold_font
    ws['B8'].font = bold_font
    ws['B11'].font = bold_font
    ws['C11'].font = bold_font
    ws['D11'].font = bold_font
    ws['E11'].font = bold_font
    ws['F11'].font = bold_font
    ws['G11'].font = bold_font

    # Merge cells
    ws.merge_cells('C5:G5')
    ws.merge_cells('C6:G6')
    ws.merge_cells('C7:G7')
    ws.merge_cells('C8:G8')

    # Apply black border to specific cells
    for row in range(5, 12 + len_po):  # Rows 1 to 3
        if row not in [9, 10]:
            for col in range(2, 8):  # Columns A to C
                cell = ws.cell(row=row, column=col)
                cell.border = black_border  # Apply the border

    # Add school name
    ws['B2'] = f"{school_name} - SWI {datetime.datetime.now().year} Order Form"
    ws['B2'].font = Font(underline='single', bold=True)
    ws['C5'] = account_number

    # Save the workbook back to the BytesIO buffer
    buffer = BytesIO()
    wb.save(buffer)

    # Reset the buffer position to the start if you need to further process or write it somewhere
    buffer.seek(0)
    return buffer


def send_email_with_attachment(buffer: BytesIO, recipient: pd.Series):
    """

    :param buffer:
    :param recipient:
    :return:
    """
    # Set your API key
    # API Keys:
    # sendgrid-api-key -> kiko.rullan@banner.co.uk
    # sendgrid-api-key-Nov24 -> nasiruddin.patel@banner.co.uk
    sendgrid_api_key = get_secret("sendgrid-api-key-Nov24")

    # Create the email object
    email_recipient = recipient.ContactEmail # recipient.ContactEmail
    to_emails = [email.strip() for email in recipient.ContactEmail.split(";")]

    message = Mail(
        from_email="nasiruddin.patel@banner.co.uk",
        to_emails=to_emails,
        subject=f"Banner Order Form - {recipient.SchoolName}",
        html_content=f'<p>Hi {recipient.ContactFirstName},</p><p>Please find attached the Stock And Sales Report. For any queries, please contact swicam@monkhouse.com. </p><p>Thank you.</p>'
    )

    # Encode the buffer content to base64
    buffer_content = buffer.getvalue()
    encoded_file = base64.b64encode(buffer_content).decode('utf-8')

    # Create an attachment object for the Excel file
    attachment = Attachment(
        FileContent(encoded_file),  # Base64 encoded file content
        FileName(f'{recipient.SchoolName}.xlsx'),  # Name of the attachment file
        FileType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),  # MIME type for Excel
        Disposition('attachment')  # Disposition (attached file)
    )

    # Add attachment to the email
    message.attachment = attachment

    try:
        # Send the email
        sg = SendGridAPIClient(sendgrid_api_key)
        response = sg.send(message)
        logging.info(f"Email sent! Status Code: {response.status_code}")
    except Exception as e:
        logging.info(f"Error sending email: {str(e)}")


@app.function_name(name="build_order_forms")
@app.route(route="build_order_forms")
def funct_build_reports(req: func.HttpRequest) -> func.HttpResponse:
    """
    TEST FUNCTION
    """

    # Read data
    df_price, df_stock, df_sales, df_recipients = read_data()

    # Loop over recipients
    # recipient.SchoolName == "The Co-Operative Academy of Manchester"
    for index, recipient in df_recipients.iterrows():
        if index < 1000:
            logging.info(recipient)
            # Build excel
            buffer = build_excel(df_price=df_price, df_stock=df_stock, df_sales=df_sales, recipient=recipient)
            # Send email
            send_email_with_attachment(buffer=buffer, recipient=recipient)

    # Connect
    # connection_string = # get_secret("bannerdwstorage_connection_string")
    # blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    # container = "school-report"
    # blob_client = blob_service_client.get_blob_client(container=container, blob="SchoolReportTemplate.xlsx")
    # blob_client.upload_blob(buffer, overwrite=True)

    return func.HttpResponse("BUILT ORDER FORM ran successfully!", status_code=200)


@app.route(route="Line_Reports")
def Line_Reports(req: func.HttpRequest) -> func.HttpResponse:

    ReportRecipients = fetch_datalake_query("SELECT * FROM [dbo].[azure_autoReport_recipients] WHERE reportBranch = 'lineReports'")
    ReportDetails = fetch_datalake_query("SELECT * FROM [dbo].[azure_autoReport_config] WHERE report_Branch = 'lineReports'")
    emailLinks = []
    logging.warning("data collected")   

    recipient_emails = ReportRecipients.iloc[0]["EmailRecipients"]
    logging.warning(f"Recipient emails: {recipient_emails}")   
    logging.warning("starting reports")   


    for row in ReportDetails.itertuples():
        try:
            dataset = fetch_datalake_query(row.SQL_Script)
            buffer = build_raw_Excel(dataset)
            fileprefix = row.Report_Name.replace(" ", "_")
            filename = (
                f"{fileprefix}_"
                f"{datetime.datetime.now():%Y%m%d_%H%M%S}.csv"
            )
            sas_url = upload_report_to_blob(
                buffer=buffer,
                filename=filename
            )

            tempTuple = (row.Report_Name, sas_url)

            emailLinks.append(tempTuple)

            del dataset
            buffer.close()
            buffer = None

        except Exception as e:
            logging.exception(
                f"Failed to process report '{row.Report_Name}': {str(e)}"
       )

    if emailLinks:
        logging.warning("starting emails")   
        send_URL_email_report(emailLinks, recipient_emails)

    return func.HttpResponse("Test Function")

def fetch_datalake_query(querystr: str):
    """

    :return:
    """
    # Define your Azure SQL Database connection_db details
    server_name = "ban-powbi-sql-01.database.windows.net"
    database_name = "banner-platform"
    driver = (
        "{ODBC Driver 18 for SQL Server}"  # Use the appropriate driver version
    )

    sqlserver_username = get_secret("ban-powbi-sql-01-username")
    sqlserver_password = get_secret("ban-powbi-sql-01-password")
    connection_string = f"DRIVER={driver};SERVER={server_name};DATABASE={database_name};UID={sqlserver_username};PWD={sqlserver_password};TrustServerCertificate=yes"

    with pyodbc.connect(connection_string) as connection:
        df = pd.read_sql(querystr, connection)

    logging.info(
        f"Query returned {len(df):,} rows x {len(df.columns)} columns"
    )

    memory_mb = df.memory_usage(deep=True).sum() / (1024 * 1024)
    logging.info(f"DataFrame memory usage: {memory_mb:.2f} MB")

    # Return Dataframes
    return df

def build_raw_Excel(
    df: pd.DataFrame
):
    """
    Build an Excel file in memory and return it as a BytesIO object.
    """

    buffer = BytesIO()

    df.to_csv(
        buffer,
        index=False
    )

    buffer.seek(0)

    excel_size_mb = len(buffer.getvalue()) / (1024 * 1024)
    logging.info(f"Excel size: {excel_size_mb:.2f} MB")

    return buffer


def upload_report_to_blob(buffer: BytesIO, filename: str):
    """
    Upload report to blob storage and return SAS URL
    """
    ##
    connection_string = get_secret("bannerreportblobsecret")
    storage_account_name = get_secret("banner-report-blob-name")
    storage_account_key = get_secret("bannerreportblob")
    container_name = "line-reports"
    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    blob_client = blob_service_client.get_blob_client(
        container=container_name,
        blob=filename
    )
    buffer.seek(0)
    logging.warning(f"uploading file")
    blob_client.upload_blob(
        buffer,
        overwrite=True
    )

    try:
        sas_token = generate_blob_sas(
            account_name=storage_account_name,
            container_name=container_name,
            blob_name=filename,
            account_key=storage_account_key,
            permission=BlobSasPermissions(read=True),
            expiry=datetime.datetime.utcnow() + datetime.timedelta(days=7)
        )

    except Exception as e:
        logging.exception("SAS generation failed")
        raise

    sas_url = (
        f"https://{storage_account_name}.blob.core.windows.net/"
        f"{container_name}/{filename}?{sas_token}"
    )

    return sas_url

def send_URL_email_report(reportLinks: [], emails: str):
    """
    Send report download link via email.

    :param report_name:
    :param sas_url:
    :return:
    """
    sendgrid_api_key = get_secret("sendgrid-api-key-Nov24")
    logging.warning(f"emails: {emails}")
    to_emails = [email.strip() for email in emails.split(";")]
    ##to_emails = "george.petch@monkhouse.com"


    links_html = ""

    for report_name, sas_url in reportLinks:
        links_html += f"""
        <p>
            {sas_url}
                Download {report_name}
            </a>
        </p>
        """

    html_content = f"""
        <p>Hi,</p>

        <p>Your reports are ready.</p>

        {links_html}

        <p>
            These links will expire in 7 days.
        </p>

        <p>
            For any queries please contact Banner IT.
        </p>

        <p>
            Thank you
        </p>
        """

    message = Mail(
        from_email="nasiruddin.patel@banner.co.uk",
        to_emails=to_emails,
        subject="Monkhouse Line Reports",
        html_content=html_content
    )

    try:
        sg = SendGridAPIClient(sendgrid_api_key)
        response = sg.send(message)

        logging.info(
            f"Email sent successfully. Status Code: {response.status_code}"
        )

    except Exception as e:
        logging.error(
            f"Error sending email: {str(e)}"
        )
        raise


@app.route(route="upload_Stock_Take_File_to_blob", auth_level=func.AuthLevel.FUNCTION)
def upload_Stock_Take_File_To_blob(req: func.HttpRequest) -> func.HttpResponse:

    file_data = req.get_body()
    logging.info(f"Received {len(file_data)} bytes")

    filename = req.headers.get(
        "x-filename",
        "upload.xlsx"
    )

    connection_string = get_secret("bannerreportblobsecret")
    storage_account_name = get_secret("banner-report-blob-name")
    storage_account_key = get_secret("bannerreportblob")

    container_name = "monkhouse-stock-take"
    logging.warning(f"got secrets for blob storage: {container_name}")

    blob_service_client = BlobServiceClient.from_connection_string(connection_string)

    logging.warning(f"connected to blob service client")

    logging.warning(f"get blob client")
    blob_client = blob_service_client.get_blob_client(
        container=container_name,
        blob=filename
    )

    blob_client.upload_blob(
        file_data,
        overwrite=True
    )

    return func.HttpResponse(
        f"Received {len(file_data)} bytes",
        status_code=200
    )


def upload_file_to_lake(buffer: BytesIO, filename: str):
    """
    Upload report to blob storage and return SAS URL
    """


@app.route(route="Retail_product_attributes", auth_level=func.AuthLevel.FUNCTION)
def Retail_product_attributes(req: func.HttpRequest) -> func.HttpResponse:

    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLake(file_data, False, "Retail_ProductAttributes", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

def get_datalake_connection():
    conn_str = get_datalake_conn_string()
    conn = pyodbc.connect(conn_str)
    return conn

def get_datalake_conn_string(): 
    # SQL Connection
    server = "ban-powbi-sql-01.database.windows.net"
    database = "banner-platform"

    sqlserver_username = get_secret("ban-powbi-sql-01-username")
    sqlserver_password = get_secret("ban-powbi-sql-01-password")

    conn_str = (
        f"Driver={{ODBC Driver 18 for SQL Server}};"
        f"Server=tcp:{server};"
        f"Database={database};"
        f"UID={sqlserver_username};"
        f"PWD={sqlserver_password};"
        "Encrypt=yes;"
        "TrustServerCertificate=yes;"
    )

    return conn_str

def get_datalake_conn_string_BO_Live(): 
    # SQL Connection
    server = "ban-powbi-sql-01.database.windows.net"
    database = "bo_live"

    sqlserver_username = get_secret("ban-powbi-sql-01-username")
    sqlserver_password = get_secret("ban-powbi-sql-01-password")

    conn_str = (
        f"Driver={{ODBC Driver 18 for SQL Server}};"
        f"Server=tcp:{server};"
        f"Database={database};"
        f"UID={sqlserver_username};"
        f"PWD={sqlserver_password};"
        "Encrypt=yes;"
        "TrustServerCertificate=yes;"
    )

    return conn_str


def trigger_sproc(sproc: str):
    conn = get_datalake_connection()
    cursor = conn.cursor()
    cursor.execute(f"EXEC {sproc}")
    conn.commit()
    conn.close()

def fetch_sproc_data_json_for_VBA(sproc: str):

    conn = get_datalake_connection()
    cursor = conn.cursor()
    cursor.execute(f"EXEC {sproc}")
    # Get column names
    columns = [col[0] for col in cursor.description]
    # Get data rows
    data = []
    try:
        for row in cursor.fetchall():
            data.append(list(row))
        result = {
            "columns": columns,
            "data": data
        }
    except Exception as e:
        logging.exception("Error fetching data from sproc")
        result = {
            "columns": columns,
            "data": [],
            "error": str(e)
        }
    conn.close()
    logging.warning(f"complete {sproc}")
    return result



def UploadCSVtoLake(filedata: bytes, delete_existing: bool, table_name: str, schema: str = "dbo") -> func.HttpResponse:

    try:
        try:
            df = pd.read_csv(
            BytesIO(filedata),
            dtype=str,
            keep_default_na=False,
            encoding="utf-8"
            )
        except UnicodeDecodeError:
            df = pd.read_csv(
            BytesIO(filedata),
            dtype=str,
            keep_default_na=False,
            encoding="cp1252"
            )

        params = urllib.parse.quote_plus(
            get_datalake_conn_string()
        )

        engine = create_engine(
            f"mssql+pyodbc:///?odbc_connect={params}",
            fast_executemany=True
        )

        # Optional truncate
        if delete_existing:
            with engine.begin() as conn:
                conn.execute(
                    text(f"TRUNCATE TABLE {schema}.[{table_name}]")
                )

        ##Used for error converting blank strings to int (if applicable)
        df = df.replace('', None)

        # Insert data
        df.to_sql(
            table_name,
            engine,
            schema=schema,
            if_exists="append",
            index=False,
            chunksize=5000
        )

        return func.HttpResponse(
            f"Uploaded {len(df):,} rows",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

def UploadCSVtoLakeAutoTypeSet(filedata: bytes, delete_existing: bool, table_name: str, schema: str = "dbo") -> func.HttpResponse:
    logging.warning(f"Uploading CSV to Lake: {table_name} in schema {schema}")
    try:
        try:
            df = pd.read_csv(
            BytesIO(filedata),

            keep_default_na=False,
            encoding="utf-8"
            )
        except UnicodeDecodeError:
            df = pd.read_csv(
            BytesIO(filedata),
            keep_default_na=False,
            encoding="cp1252"
            )

        params = urllib.parse.quote_plus(
            get_datalake_conn_string()
        )

        engine = create_engine(
            f"mssql+pyodbc:///?odbc_connect={params}",
            fast_executemany=True
        )

        # Optional truncate
        if delete_existing:
            with engine.begin() as conn:
                conn.execute(
                    text(f"TRUNCATE TABLE {schema}.[{table_name}]")
                )

        ##Used for error converting blank strings to int (if applicable)
        df = df.replace('', None)

        for col in df.columns:
            logging.warning(f"Column: {col}, Data Type: {df[col].dtype}")

        # Insert data
        df.to_sql(
            table_name,
            engine,
            schema=schema,
            if_exists="append",
            index=False,
            chunksize=5000
        )

        return func.HttpResponse(
            f"Uploaded {len(df):,} rows",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )


def UploadCSVtoLakeCustomTypeSet(file_data: bytes, delete_existing: bool, table_name: str, schema: str = "dbo") -> func.HttpResponse:


    params = urllib.parse.quote_plus(
        get_datalake_conn_string()
    )

    engine = create_engine(
        f"mssql+pyodbc:///?odbc_connect={params}",
        fast_executemany=True
    )

    # Optional truncate
    if delete_existing:
        with engine.begin() as conn:
            conn.execute(
                text(f"TRUNCATE TABLE {schema}.[{table_name}]")
            )

    logging.warning(f"Uploading CSV to Lake: {table_name} in schema {schema}")
    # Read raw CSV with no header
    try: 
        raw_df = pd.read_csv(
            io.BytesIO(file_data),
            header=None,
            dtype=str,
            keep_default_na=False,
            encoding="utf-8"
        )

    except UnicodeDecodeError:
        raw_df = pd.read_csv(
        BytesIO(file_data),
        header=None,
        dtype=str,
        keep_default_na=False,
        encoding="cp1252"
        )

    # Row 1 = column names
    column_names = raw_df.iloc[0].tolist()

    # Row 2 = datatypes
    column_types = raw_df.iloc[1].tolist()

    # Row 3 onwards = data
    df = raw_df.iloc[2:].copy()

    # Apply headers
    df.columns = column_names

    # Reset row index
    df.reset_index(drop=True, inplace=True)

    for col in df.columns:
        max_len = df[col].astype(str).str.len().max()
        print(f"{col}: {max_len}")

    # Apply datatypes
    for col, datatype in zip(column_names, column_types):

        datatype = str(datatype).strip().upper()

        if datatype == "INT":
            df[col] = pd.to_numeric(df[col], errors="coerce")

        elif datatype == "FLOAT":
            df[col] = pd.to_numeric(df[col], errors="coerce")

        elif datatype == "DECIMAL":
            df[col] = pd.to_numeric(df[col], errors="coerce")

        elif datatype == "DATE":
            df[col] = pd.to_datetime(
                df[col],
                dayfirst=True,
                errors="coerce"
            ).dt.date

        elif datatype == "DATETIME":
            df[col] = pd.to_datetime(
                df[col],
                dayfirst=True,
                errors="coerce"
            )

        else:
            df[col] = df[col].astype(str)

    print(df.dtypes)

    df.to_sql(
        table_name,
        engine,
        schema=schema,
        if_exists="append",
        index=False,
        method=None
    )


@app.route(route="PO_Upload_1a_upload_staging_test", auth_level=func.AuthLevel.FUNCTION)
def PO_Upload_1a_upload_staging_test(req: func.HttpRequest) -> func.HttpResponse:

    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(file_data, True, "Retail_PurchaseOrder_1_Staging", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="PO_Upload_1b_Fetch_Staging_Test", auth_level=func.AuthLevel.FUNCTION)
def PO_Upload_1b_Fetch_Staging_Test(req: func.HttpRequest) -> func.HttpResponse:

    try: 
        logging.warning(f"Fetching data from sproc: dbo.[Retail_PurchaseOrders_1_staging_test]")

        result = fetch_sproc_data_json_for_VBA("dbo.[Retail_PurchaseOrders_1_staging_test]")

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )
    except Exception as e:
        logging.exception(str(e))
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="PO_Upload_2a_upload_approved", auth_level=func.AuthLevel.FUNCTION)
def PO_Upload_2a_upload_approved(req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("Uploading approved purchase orders")
    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(file_data, False, "Retail_PurchaseOrders_2_Approved", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )



def fetch_sproc_data_json_for_VBA(sproc: str):

    conn = get_datalake_connection()
    cursor = conn.cursor()
    cursor.execute(f"EXEC {sproc}")
    columns = [col[0] for col in cursor.description]
    # Get data rows
    data = []
    logging.warning(f"Fetching data rows")
    try:
        for row in cursor.fetchall():
            data.append(list(row))
        result = {
            "columns": columns,
            "data": data
        }
    except Exception as e:
        logging.exception("Error fetching data from sproc")
        result = {
            "columns": columns,
            "data": [],
            "error": str(e)
        }
    conn.close()
    logging.warning(f"complete {sproc}")
    return result

def fetch_view_data_json_for_VBA(view_name: str):

    conn = get_datalake_connection()
    cursor = conn.cursor()

    cursor.execute(f"SELECT * FROM {view_name}")

    columns = [col[0] for col in cursor.description]

    data = []
    logging.warning(f"Fetching data from view {view_name}")

    try:
        for row in cursor.fetchall():
            data.append(list(row))

        result = {
            "columns": columns,
            "data": data
        }

    except Exception as e:
        logging.exception("Error fetching data from view")
        result = {
            "columns": columns,
            "data": [],
            "error": str(e)
        }

    conn.close()
    logging.warning(f"Complete {view_name}")

    return result

@app.route(route="Replen_Upload", auth_level=func.AuthLevel.FUNCTION)
def Replen_Upload(req: func.HttpRequest) -> func.HttpResponse:

    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeBOLive(file_data, True, "store_replen_master", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

def UploadCSVtoLakeBOLive(filedata: bytes, delete_existing: bool, table_name: str, schema: str = "dbo") -> func.HttpResponse:

    try:
        try:
            df = pd.read_csv(
            BytesIO(filedata),
            dtype=str,
            keep_default_na=False,
            encoding="utf-8"
            )
        except UnicodeDecodeError:
            df = pd.read_csv(
            BytesIO(filedata),
            dtype=str,
            keep_default_na=False,
            encoding="cp1252"
            )

        params = urllib.parse.quote_plus(
            get_datalake_conn_string_BO_Live()
        )

        engine = create_engine(
            f"mssql+pyodbc:///?odbc_connect={params}",
            fast_executemany=True
        )

        # Optional truncate
        if delete_existing:
            with engine.begin() as conn:
                conn.execute(
                    text(f"TRUNCATE TABLE {schema}.[{table_name}]")
                )

        ##Used for error converting blank strings to int (if applicable)
        df = df.replace('', None)

        # Insert data
        df.to_sql(
            table_name,
            engine,
            schema=schema,
            if_exists="append",
            index=False,
            chunksize=5000
        )

        return func.HttpResponse(
            f"Uploaded {len(df):,} rows",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="fetchStoreSales", auth_level=func.AuthLevel.FUNCTION)
def fetchStoreByWeekByLocation(req: func.HttpRequest) -> func.HttpResponse:

    try:
        body = req.get_json()
    except ValueError:
        return func.HttpResponse(
            "Request body must be valid JSON",
            status_code=400
        )

    store_code = body.get("storeCode")

    if not store_code:
        return func.HttpResponse(
            "storeCode is required",
            status_code=400
        )

    logging.warning(f"Fetching store sales for store code: {store_code}")

    try:
        result = fetch_sproc_data_json_for_VBA(
            f"dbo.retail_fetchStoreWeeklySales @location_name='{store_code}'"
        )

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )


def fetch_sproc_data_csv(sproc: str):

    logging.warning(f"Fetching data from sproc: {sproc}")

    conn = get_datalake_connection()
    cursor = conn.cursor()

    cursor.execute(f"EXEC {sproc}")

    output = io.StringIO()
    writer = csv.writer(output)

    # headers
    columns = [col[0] for col in cursor.description]
    writer.writerow(columns)

    # data
    for row in cursor:
        writer.writerow(row)

    conn.close()

    return output.getvalue()

@app.route(route="fetchStoreByWeekByLocationCSV", auth_level=func.AuthLevel.FUNCTION)
def fetchStoreByWeekByLocationCSV(req: func.HttpRequest) -> func.HttpResponse:

    try:
        body = req.get_json()
    except ValueError:
        return func.HttpResponse(
            "Request body must be valid JSON",
            status_code=400
        )

    store_code = body.get("storeCode")

    if not store_code:
        return func.HttpResponse(
            "storeCode is required",
            status_code=400
        )

    try:

        csv_data = fetch_sproc_data_csv(
            f"dbo.retail_fetchStoreWeeklySales @location_name='{store_code}'"
        )

        return func.HttpResponse(
            csv_data,
            mimetype="text/csv",
            headers={
                "Content-Disposition":
                f'attachment; filename="{store_code}_sales.csv"'
            },
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

@app.route(route="data_Upload_SC_LeadTimeMatrix", auth_level=func.AuthLevel.FUNCTION)
def data_Upload_SC_LeadTimeMatrix(req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("Uploading supply chain lead time matrix")
    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(file_data, True, "Retail_SupplyChain_Product_Leadtimes", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="data_Upload_SC_MOQs", auth_level=func.AuthLevel.FUNCTION)
def data_Upload_SC_MOQs(req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("Uploading supply chain MOQs")
    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(file_data, True, "Retail_SupplyChain_Product_MOQs", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="data_Upload_Location_Capacity", auth_level=func.AuthLevel.FUNCTION)
def data_Upload_Location_Capacity(req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("Uploading location unit capacity data")
    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(file_data, True, "Retail_SupplyChain_Location_Capacity", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="data_Upload_Location_Budget_Capacity", auth_level=func.AuthLevel.FUNCTION)
def data_Upload_Location_Budget_Capacity(req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("Uploading location budget capacity data")
    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(file_data, True, "Retail_SupplyChain_Budget_Capacity", "dbo")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="fetchStockAgeReport", auth_level=func.AuthLevel.FUNCTION)
def fetchStockAgeReport(req: func.HttpRequest) -> func.HttpResponse:

    try:
        result = fetch_view_data_csv(
            f"dbo.retail_age_of_stock_report"
        )

        return func.HttpResponse(
            result,
            mimetype="text/csv",
            headers={
                "Content-Disposition":
                f'attachment; filename="age_of_stock.csv"'
            },
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

def fetch_view_data_csv(View: str):

    logging.warning(f"Fetching data from view: {View}")

    conn = get_datalake_connection()
    cursor = conn.cursor()

    cursor.execute(f"SELECT * FROM {View}")

    output = io.StringIO()
    writer = csv.writer(output)

    # headers
    columns = [col[0] for col in cursor.description]
    writer.writerow(columns)

    # data
    for row in cursor:
        writer.writerow(row)

    conn.close()

    return output.getvalue()

@app.route(route="data_upload_custom_attribute", auth_level=func.AuthLevel.FUNCTION)
def data_upload_custom_attribute(req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("Uploading custom product attributes")
    try:
        # Get uploaded file
        file_data = req.get_body()

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(file_data, True, "Retail_Custom_Product_Attribute_Staging", "dbo")

        trigger_sproc("dbo.retail_update_custom_product_attribute")

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")
        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="fetch_custom_attributes", auth_level=func.AuthLevel.FUNCTION)
def fetch_custom_attributes(req: func.HttpRequest) -> func.HttpResponse:

    trigger_sproc("dbo.Build_Retail_Custom_Product_Attributes_Pivot")

    try:
        result = fetch_view_data_csv(
            f"dbo.Retail_Custom_Product_Attributes_Pivot"
        )

        return func.HttpResponse(
            result,
            mimetype="text/csv",
            headers={
                "Content-Disposition":
                f'attachment; filename="age_of_stock    .csv"'
            },
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )



@app.route(
    route="data_upload_to_lake_generic", auth_level=func.AuthLevel.FUNCTION)
def data_upload_to_lake_generic(    req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("generic uploader")

    try:
        file_data = req.get_body()

        table_name = req.headers.get("x-table-name")
        schema_name = req.headers.get("x-schema-name")
        truncate_table = req.headers.get("x-truncate_table")

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        if not table_name:
            return func.HttpResponse(
                "No table name supplied",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSet(
            file_data,
            truncate_table,
            table_name,
            schema_name
        )

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes to {schema_name}.{table_name}",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")

        return func.HttpResponse(
            str(e),
            status_code=500
        )

@app.route(route="data_upload_to_lake_generic", auth_level=func.AuthLevel.FUNCTION)
def data_upload_to_lake_generic_with_upsert(req: func.HttpRequest) -> func.HttpResponse:
    logging.warning("generic uploader")

    try:
        file_data = req.get_body()

        table_name = req.headers.get("x-table-name")
        schema_name = req.headers.get("x-schema-name")
        truncate_table = req.headers.get("x-truncate_table")
        key_columns = req.headers.get("x-key-columns")

        if not file_data:
            return func.HttpResponse(
                "No file received",
                status_code=400
            )

        if not table_name:
            return func.HttpResponse(
                "No table name supplied",
                status_code=400
            )

        UploadCSVtoLakeCustomTypeSetWithUpsert(
            file_data,
            truncate_table,
            table_name,
            key_columns,
            schema_name
        )

        return func.HttpResponse(
            f"Uploaded {len(file_data)} bytes to {schema_name}.{table_name}",
            status_code=200
        )

    except Exception as e:
        logging.exception("Upload failed")

        return func.HttpResponse(
            str(e),
            status_code=500
        )

def UploadCSVtoLakeCustomTypeSetWithUpsert(
    file_data: bytes,
    delete_existing: bool,
    table_name: str,
    key_columns: list[str],
    schema: str = "dbo"
):

    params = urllib.parse.quote_plus(
        get_datalake_conn_string()
    )

    engine = create_engine(
        f"mssql+pyodbc:///?odbc_connect={params}",
        fast_executemany=True
    )

    # Optional truncate
    if delete_existing:
        with engine.begin() as conn:
            conn.execute(
                text(f"TRUNCATE TABLE {schema}.[{table_name}]")
            )

    logging.warning(
        f"Uploading CSV to Lake: {table_name} in schema {schema}"
    )

    # -------------------------------
    # Read CSV
    # -------------------------------

    try:
        raw_df = pd.read_csv(
            io.BytesIO(file_data),
            header=None,
            dtype=str,
            keep_default_na=False,
            encoding="utf-8"
        )

    except UnicodeDecodeError:

        raw_df = pd.read_csv(
            io.BytesIO(file_data),
            header=None,
            dtype=str,
            keep_default_na=False,
            encoding="cp1252"
        )

    # Row 1 = Column Names
    column_names = raw_df.iloc[0].tolist()

    # Row 2 = Datatypes
    column_types = raw_df.iloc[1].tolist()

    # Row 3+ = Data
    df = raw_df.iloc[2:].copy()

    df.columns = column_names

    df.reset_index(
        drop=True,
        inplace=True
    )

    # -------------------------------
    # Apply Datatypes
    # -------------------------------

    for col, datatype in zip(column_names, column_types):

        datatype = str(datatype).strip().upper()

        if datatype == "INT":

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            )

        elif datatype in ("FLOAT", "DECIMAL"):

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            )

        elif datatype == "DATE":

            df[col] = pd.to_datetime(
                df[col],
                dayfirst=True,
                errors="coerce"
            ).dt.date

        elif datatype == "DATETIME":

            df[col] = pd.to_datetime(
                df[col],
                dayfirst=True,
                errors="coerce"
            )

        else:

            df[col] = df[col].astype(str)

    logging.warning(
        f"Processed {len(df)} rows"
    )

    # -------------------------------
    # Build Existing Key Set
    # -------------------------------

    key_select = ", ".join(
        [f"[{col}]" for col in key_columns]
    )

    existing_keys_df = pd.read_sql(
        f"""
        SELECT {key_select}
        FROM {schema}.[{table_name}]
        """,
        engine
    )

    existing_keys = set(
        existing_keys_df[key_columns]
        .fillna("")
        .astype(str)
        .agg("|".join, axis=1)
    )

    # -------------------------------
    # Generate Composite Key
    # -------------------------------

    df["_upsert_key"] = (
        df[key_columns]
        .fillna("")
        .astype(str)
        .agg("|".join, axis=1)
    )

    # -------------------------------
    # Split Updates / Inserts
    # -------------------------------

    inserts = df[
        ~df["_upsert_key"].isin(existing_keys)
    ].copy()

    updates = df[
        df["_upsert_key"].isin(existing_keys)
    ].copy()

    logging.warning(
        f"Found {len(inserts)} inserts and {len(updates)} updates"
    )

    # -------------------------------
    # INSERT NEW RECORDS
    # -------------------------------

    if not inserts.empty:

        inserts = inserts.drop(
            columns=["_upsert_key"],
            errors="ignore"
        )

        inserts.to_sql(
            table_name,
            engine,
            schema=schema,
            if_exists="append",
            index=False,
            method=None
        )

        logging.warning(
            f"Inserted {len(inserts)} rows"
        )

    # -------------------------------
    # UPDATE EXISTING RECORDS
    # -------------------------------

    if not updates.empty:

        update_columns = [
            col
            for col in updates.columns
            if col not in key_columns
            and col != "_upsert_key"
        ]

        set_clause = ", ".join(
            [
                f"[{col}] = :{col}"
                for col in update_columns
            ]
        )

        where_clause = " AND ".join(
            [
                f"[{col}] = :{col}"
                for col in key_columns
            ]
        )

        update_sql = text(
            f"""
            UPDATE {schema}.[{table_name}]
            SET {set_clause}
            WHERE {where_clause}
            """
        )

        with engine.begin() as conn:

            for _, row in updates.iterrows():

                params = {}

                for col in updates.columns:

                    if col == "_upsert_key":
                        continue

                    value = row[col]

                    if pd.isna(value):
                        value = None

                    params[col] = value

                conn.execute(
                    update_sql,
                    params
                )

        logging.warning(
            f"Updated {len(updates)} rows"
        )

    return {
        "table": table_name,
        "rows_processed": len(df),
        "rows_inserted": len(inserts),
        "rows_updated": len(updates)
    }

@app.route(
    route="fetch_generic_view_data_CSV",
    auth_level=func.AuthLevel.FUNCTION
)
def fetch_generic_view_data_CSV(req: func.HttpRequest) -> func.HttpResponse:

    view_name = req.headers.get("x-view-name")
    schema_name = req.headers.get("x-schema-name")

    if not view_name or not schema_name:
        return func.HttpResponse(
            "Missing x-view-name or x-schema-name header",
            status_code=400
        )

    try:
        view = f"{schema_name}.{view_name}"

        result = fetch_view_data_csv(view)

        return func.HttpResponse(
            result,
            mimetype="text/csv",
            headers={
                "Content-Disposition":
                f'attachment; filename="{view_name}.csv"'
            },
            status_code=200
        )

    except Exception:
        logging.exception(f"Fetch failed for view {schema_name}.{view_name}")

        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

@app.route(route="fetch_TC_orders_csv", auth_level=func.AuthLevel.FUNCTION)
def fetch_TC_orders_csv(req: func.HttpRequest) -> func.HttpResponse:

    try:
        conn = get_datalake_connection()
        cursor = conn.cursor()

        sql = """
        SELECT 
              [order_id]
            , [order_id_num]
            , [order_number]
            , CAST([created_at] AS VARCHAR(50))     AS [created_at]
            , CAST([cancelled_at] AS VARCHAR(50))   AS [cancelled_at]
            , [currency_code]
            , [financial_status]
            , [fulfillment_status]
            , [source_name]
            , [is_test]
            , [customer_first_name]
            , [customer_last_name]
            , [shipping_name]
            , [billing_name]
            , [order_line]
            , [line_item_id]
            , [title]
            , [sku]
            , [variant_title]
            , [vendor]
            , [product_id]
            , [product_id_num]
            , [variant_id]
            , [variant_id_num]
            , [option1_name]
            , [option1_value]
            , [option2_name]
            , [option2_value]
            , [option3_name]
            , [option3_value]
            , [quantity]
            , [current_quantity]
            , [fulfillable_quantity]
            , [taxable]
            , [requires_shipping]
            , [unit_price]
            , [discounted_unit_price]
            , [total_discount]
            , [line_value_before_discount]
            , [line_value_after_discount]
            , [price_currency]
        FROM [TC_View].[Orders]
        """

        cursor.execute(sql)

        import csv
        import io

        output = io.StringIO()
        writer = csv.writer(output)

        headers = [column[0] for column in cursor.description]
        writer.writerow(headers)

        for row in cursor.fetchall():
            writer.writerow(row)

        csv_data = output.getvalue()

        cursor.close()
        conn.close()

        return func.HttpResponse(
            csv_data,
            mimetype="text/csv",
            headers={
                "Content-Disposition":
                    'attachment; filename="Orders.csv"'
            },
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed for TC_View.Orders")

        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

@app.route(route="fetchGTMRProducts", auth_level=func.AuthLevel.FUNCTION)
def fetchGTMRProducts(req: func.HttpRequest) -> func.HttpResponse:

    try:
        result = fetch_view_data_json(f"dbo.Retail_GTMR_ProductSelectionData")

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

@app.route(route="fetchGTMRProductGroups", auth_level=func.AuthLevel.FUNCTION)
def fetchGTMRProductGroups(req: func.HttpRequest) -> func.HttpResponse:

    try:
        result = fetch_view_data_json(f"dbo.Retail_GTMR_Product_Groups")

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )


@app.route(route="fetchGTMRProductSizeDetail", auth_level=func.AuthLevel.FUNCTION)
def fetchGTMRProductSizeDetail(req: func.HttpRequest) -> func.HttpResponse:

    try:
        result = fetch_view_data_json(f"dbo.Retail_GTMR_Product_Size_Detail")

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

@app.route(route="fetchGTMRProductColour", auth_level=func.AuthLevel.FUNCTION)
def fetchGTMRProductColour(req: func.HttpRequest) -> func.HttpResponse:

    try:
        result = fetch_view_data_json(f"dbo.Retail_GTMR_Product_Colours")

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

@app.route(route="fetchGTMRPlainItems", auth_level=func.AuthLevel.FUNCTION)
def fetchGTMRPlainItems(req: func.HttpRequest) -> func.HttpResponse:

    try:
        result = fetch_view_data_json(f"dbo.Retail_GTMR_Plain_Items")

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )


@app.route(
    route="NSP_Create_New_NSP",
    auth_level=func.AuthLevel.FUNCTION
)
def NSP_Create_New_NSP(req: func.HttpRequest) -> func.HttpResponse:

    try:

        result = execute_sproc(
            "dbo.Retail_NSP_Create_New_NSP"
        )

        return func.HttpResponse(
            json.dumps(
                {
                    "success": True,
                    "data": result[0] if result else None
                },
                default=str
            ),
            mimetype="application/json",
            status_code=200
        )

    except Exception as ex:

        logging.exception("NSP creation failed")

        return func.HttpResponse(
            json.dumps(
                {
                    "success": False,
                    "error": str(ex)
                }
            ),
            mimetype="application/json",
            status_code=500
        )

def execute_sproc(
    sproc_name: str,
    params: tuple = None
):
    """
    Execute a stored procedure and return any result set.

    Example:
        execute_sproc("dbo.Retail_NSP_Create_New_NSP")
        execute_sproc("dbo.MySproc", ("ABC", 123))
    """

    conn = get_datalake_connection()

    try:
        cursor = conn.cursor()

        if params:
            placeholders = ",".join(["?"] * len(params))
            sql = f"EXEC {sproc_name} {placeholders}"
            cursor.execute(sql, params)
        else:
            cursor.execute(f"EXEC {sproc_name}")

        results = []

        if cursor.description:
            columns = [col[0] for col in cursor.description]

            for row in cursor.fetchall():
                results.append(
                    dict(zip(columns, row))
                )

        conn.commit()

        return results

    except Exception:
        logging.exception(f"Failed executing {sproc_name}")
        raise

    finally:
        conn.close()

@app.route(route="NSP_fetch_options", auth_level=func.AuthLevel.FUNCTION)
def NSP_fetch_options(req: func.HttpRequest) -> func.HttpResponse:

    try:
        result = fetch_view_data_json(f"dbo.Retail_NSP_options")

        return func.HttpResponse(
            json.dumps(result, default=str),
            mimetype="application/json",
            status_code=200
        )

    except Exception:
        logging.exception("Fetch failed")
        return func.HttpResponse(
            "Internal server error",
            status_code=500
        )

@app.route(route="NSP_Upload_School_Info", auth_level=func.AuthLevel.FUNCTION)
def NSP_Upload_School_Info(req: func.HttpRequest) -> func.HttpResponse:

    try:
        # Raw request body
        body = req.get_body().decode("utf-8")

        logging.info("=== NSP Upload Received ===")
        logging.info(f"Method: {req.method}")
        logging.info(f"Headers: {dict(req.headers)}")
        logging.info(f"Params: {dict(req.params)}")
        logging.info(f"Raw Body: {body}")

        # Attempt to parse JSON
        try:
            payload = req.get_json()

            logging.info("=== JSON Payload ===")
            logging.info(json.dumps(payload, indent=2))

        except ValueError:
            payload = {"raw_body": body}
            logging.warning("Request body is not valid JSON")

        return func.HttpResponse(
            json.dumps({
                "success": True,
                "message": "Payload received",
                "received": payload
            }),
            mimetype="application/json",
            status_code=200
        )

    except Exception as ex:
        logging.exception("NSP upload failed")

        return func.HttpResponse(
            json.dumps({
                "success": False,
                "error": str(ex)
            }),
            mimetype="application/json",
            status_code=500
        )
def fetch_view_data_json(view_name: str):

    conn = get_datalake_connection()
    cursor = conn.cursor()

    cursor.execute(f"SELECT * FROM {view_name}")

    columns = [col[0] for col in cursor.description]

    try:

        rows = cursor.fetchall()

        data = [
            dict(zip(columns, row))
            for row in rows
        ]

        result = {
            "data": data
        }

    except Exception as e:

        logging.exception("Error fetching data from view")

        result = {
            "data": [],
            "error": str(e)
        }

    conn.close()

    return result


@app.route(route="email_ingest", auth_level=func.AuthLevel.FUNCTION)
def email_ingest(req: func.HttpRequest):

    data = req.get_json()

    logging.info(f"Subject: {data['subject']}")
    logging.info(f"From: {data['sender']}")

    for attachment in data.get("attachments", []):
        filename = attachment["fileName"]

        content = base64.b64decode(
            attachment["contentBase64"]
        )

        with open(f"/tmp/{filename}", "wb") as f:
            f.write(content)

    return func.HttpResponse(
        json.dumps({"success": True}),
        mimetype="application/json"
    )



@app.route(route="email_ingest_ai", auth_level=func.AuthLevel.FUNCTION)
def email_ingest_ai(req: func.HttpRequest) -> func.HttpResponse:

    try:

        # Parse payload
        data = req.get_json()

        subject = data.get("subject", "")
        sender = data.get("sender", "")
        body = data.get("body", "")
        attachments = data.get("attachments", [])

        logging.info(f"Subject: {subject}")
        logging.info(f"Sender : {sender}")
        logging.info(f"Attachments: {len(attachments)}")

        # Save attachments
        saved_files = []

        for attachment in attachments:

            filename = attachment.get("fileName")
            content_base64 = attachment.get("contentBase64")

            if not filename or not content_base64:
                continue

            file_bytes = base64.b64decode(content_base64)

            file_path = os.path.join(
                tempfile.gettempdir(),
                filename
            )

            with open(file_path, "wb") as f:
                f.write(file_bytes)

            saved_files.append(
                {
                    "filename": filename,
                    "path": file_path,
                    "size": len(file_bytes)
                }
            )

        # Azure OpenAI / Foundry Client
        client = AzureOpenAI(
            api_key=os.environ["AZURE_OPENAI_KEY"],
            api_version="2024-10-21",
            azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"]
        )

        ai_response = client.chat.completions.create(
            model=os.environ["AZURE_OPENAI_DEPLOYMENT"],
            messages=[
                {
                    "role": "system",
                    "content": """
                    You are an email processing assistant.

                    Return VALID JSON ONLY.

                    Extract:

                    {
                        "intent": "",
                        "customer_name": "",
                        "business_name": "",
                        "SQL_search_business_name": "",
                        "delivery_address": "",
                        invoice_address": "",
                        "school_name": "",
                        "order_numbers": [],
                        "action_required": "",
                        "priority": "",
                        "summary": ""
                    }

                    for the sql_search_business_name, create a minimal string of just one word to maximise search likelihood and remove any special characters or punctuation.

                    Priority must be:
                    Low, Medium, High or Critical.
                    """
                                    },
                                    {
                                        "role": "user",
                                        "content": f"""
                    EMAIL DETAILS

                    From:
                    {sender}

                    Subject:
                    {subject}

                    Body:
                    {body}
                    """
                }
            ]
        )

        ai_result = ai_response.choices[0].message.content

        try:
            parsed_result = json.loads(ai_result)
        except Exception:
            parsed_result = {
                "raw_response": ai_result
            }

        customer_name = parsed_result.get("business_name")
        customer_sql_name = parsed_result.get("SQL_search_business_name")
        delivery_address = parsed_result.get("delivery_address")
        invoice_address = parsed_result.get("invoice_address")

        likeness_score = 95

        logging.info(f"Extracted delivery address: {delivery_address}")
        logging.info(f"Extracted invoice address: {invoice_address}")
        logging.info(f"Extracted customer name: {customer_name}")

        customer_search_data = fetch_sproc_data_json(
           f"dbo.breakneck_customers @customer='{customer_sql_name}', @Likeness={likeness_score}"
        )
        logging.info(f"Extracted customer name: {customer_sql_name}")

        customer_search_data = fetch_sproc_data_json(
            f"dbo.breakneck_customers @customer='{customer_sql_name}', @Likeness={likeness_score}"
        )

        logging.info(f"Customer search result: {customer_search_data}")

        ai_response = client.chat.completions.create(
            model=os.environ["AZURE_OPENAI_DEPLOYMENT"],
            messages=[
                {
                    "role": "system",
                    "content": f"""
                    You are an email processing assistant.

                    Return VALID JSON ONLY.

                    Identify the most likely customer from the previous order data.

                    Return:

                    {{
                        "Selected_Business_ID": "",
                        "selected_Business_Name": "",
                        "selected_business_invoice_address": ""
                    }}

                    Email data:
                    {json.dumps(parsed_result)}

                    Previous order data:
                    {json.dumps(customer_search_data)}
                    """
                }
            ]
        )

        ai_result_2 = ai_response.choices[0].message.content

        try:
            parsed_result_2 = json.loads(ai_result_2)
        except Exception:
            parsed_result_2 = {
                "raw_response": ai_result_2
            }

        logging.info(f"AI result 2: {parsed_result_2}")

        selected_business_id = parsed_result_2.get("Selected_Business_ID")
        selected_business_name = parsed_result_2.get("selected_Business_Name")

        product_search_data = fetch_sproc_data_json(
            f"dbo.breakneck_products_2 @customer_code='{selected_business_id}'"
        )

        logging.info(f"Product search result: {product_search_data}")

        ai_response = client.chat.completions.create(
            model=os.environ["AZURE_OPENAI_DEPLOYMENT"],
            messages=[
                {
                    "role": "system",
                    "content": f"""
                    You are an order matching assistant.

                    Return VALID JSON ONLY.

                    Do not return markdown.
                    Do not return explanations.
                    Do not return code fences.
                    Do not return text before or after the JSON.

                    Use ONLY this structure:

                    {{
                        "products": [
                            {{
                                "product": "PRODUCT_CODE",
                                "colour_code": "COLOUR_CODE",
                                "quantity": 1
                            }}
                        ]
                    }}

                    Rules:

                    - product must be the matched product code from the supplied data.
                    - colour_code must be the matched colour code from the supplied data.
                    - quantity must be an integer.
                    - If quantity is not stated in the email use 1.
                    - Return only products that are confidently matched.
                    - Do not invent product codes.
                    - Do not invent colour codes.
                    - All products must come from the supplied product data.

                    Email data:
                    {json.dumps(parsed_result)}

                    Previous product data:
                    {json.dumps(product_search_data)}
                    """
                }
            ]
        )

        ai_response_3 = ai_response.choices[0].message.content

        try:
            parsed_result_3 = json.loads(ai_response_3)
        except Exception:
            parsed_result_3 = {
                "raw_response": ai_response_3
            }

        logging.info(f"AI result 3: {parsed_result_3}")

        product_size_data = fetch_sproc_data_json_explicit(
            "EXEC dbo.[breakneck_product_sizes_3] @items=?",
            [json.dumps(parsed_result_3)]
        )

        ai_response = client.chat.completions.create(
            model=os.environ["AZURE_OPENAI_DEPLOYMENT"],
            messages=[
                {
                    "role": "system",
                    "content": f"""
                    You are an email processing assistant.

                    Return VALID JSON ONLY.

                    
                    Return a json list with the parsed results, the parsed_results_2 for the selected customer, the selected products, and finally select the sizes from the product_size_data for the selected products in a final json list.

                    :
                    'products': [{'product': '', 'description': '', 'colour_code': '', 'colour': '', 'quantity': 21}]


                    Email data:
                    {json.dumps(parsed_result)}

                    Selected Products:  
                    {json.dumps(parsed_result_3)}

                    product size data:
                    {json.dumps(product_size_data)}
                    """
                }
            ]
        )

        ai_response_4 = ai_response.choices[0].message.content

        try:
            parsed_result_4 = json.loads(ai_response_4)
        except Exception:
            parsed_result_4 = {
                "raw_response": ai_response_4
            }    

        logging.info(f"AI result 4: {parsed_result_4}")

        return func.HttpResponse(
            json.dumps(
                {
                    "success": True,
                    "email": {
                        "sender": sender,
                        "subject": subject,
                        "attachment_count": len(saved_files)
                    },
                    "attachments": saved_files,
                    "ai_result": parsed_result
                },
                default=str
            ),
            mimetype="application/json",
            status_code=200
        )

    except Exception as ex:

        logging.exception("Email processing failed")

        return func.HttpResponse(
            json.dumps(
                {
                    "success": False,
                    "error": str(ex),
                    "error_type": type(ex).__name__
                }
            ),
            mimetype="application/json",
            status_code=500
        )

def fetch_sproc_data_json(sproc_name: str):

    conn = get_datalake_connection()
    cursor = conn.cursor()

    cursor.execute(f"EXEC {sproc_name}")

    columns = [col[0] for col in cursor.description]

    try:

        rows = cursor.fetchall()

        data = [
            dict(zip(columns, row))
            for row in rows
        ]

        result = {
            "data": data
        }

    except Exception as e:

        logging.exception("Error fetching data from sproc")

        result = {
            "data": [],
            "error": str(e)
        }

    conn.close()

    return result

def fetch_sproc_data_json_explicit(
    sql: str,
    params: list = None
):

    conn = get_datalake_connection()
    cursor = conn.cursor()

    if params:
        cursor.execute(sql, params)
    else:
        cursor.execute(sql)

    columns = [col[0] for col in cursor.description]

    try:

        rows = cursor.fetchall()

        data = [
            dict(zip(columns, row))
            for row in rows
        ]

        result = {
            "data": data
        }

    except Exception as e:

        logging.exception("Error fetching data from sproc")

        result = {
            "data": [],
            "error": str(e)
        }

    conn.close()

    return result
