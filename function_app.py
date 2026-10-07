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
            buffer.close
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


def trigger_sproc(sproc: str):
    logging.warning(f"Triggering sproc: {sproc}")
    conn = get_datalake_connection()
    logging.warning(f"Connected to database")
    cursor = conn.cursor()
    logging.warning(f"Executing sproc: {sproc}")
    cursor.execute(f"EXEC {sproc}")
    logging.warning(f"Executed sproc: {sproc}")
    conn.commit()
    conn.close()

def fetch_sproc_data_json_for_VBA(sproc: str):
    logging.warning(f"Fetching data from sproc: {sproc}")
    conn = get_datalake_connection()
    logging.warning(f"Connected to database")
    cursor = conn.cursor()
    logging.warning(f"Executing sproc: {sproc}")
    cursor.execute(f"EXEC {sproc}")
    logging.warning(f"Executed sproc: {sproc}")
    # Get column names
    logging.warning(f"Fetching column names")
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