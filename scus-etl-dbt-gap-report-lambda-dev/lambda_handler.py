import os
import psycopg2
import json
import boto3
import uuid
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import datetime
import io
import xlsxwriter


SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
EMAIL_QUEUE_URL = os.environ.get("EMAIL_QUEUE_URL")
DB_HOST = os.environ.get("DB_HOST")
DB_PORT = os.environ.get("DB_PORT")
DB_NAME = os.environ.get("DB_NAME")
DB_USER = os.environ.get("DB_USER")
DB_PASSWORD = os.environ.get("DB_PASSWORD")
GAP_REDSHIFT_TABLE = os.environ.get("GAP_REDSHIFT_TABLE")
CLIENT_CONFIG_REDSHIFT_TABLE = os.environ.get("CLIENT_CONFIG_REDSHIFT_TABLE")

def Connection(host,port,database,user,password):
    conn = psycopg2.connect(host=host, database=database, port=int(port), user=user, password=password)
    print("Connection Established")
    return conn

def generateReportandUploadToS3(conn,redshift_tables,s3_bucket,s3_key):
    s3_urls = []
    for redshift_table in redshift_tables:
        # Get the current date and time
        current_datetime = datetime.datetime.now()
        # Format the date and time as per your requirement
        formatted_datetime = current_datetime.strftime("%Y-%m-%d %H:%M")
        temp_s3_key = f'{s3_key}/{redshift_table}_{formatted_datetime}.xlsx'
        # Retrieve data from Redshift table
        query = f"SELECT * FROM {redshift_table}"
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        # Create an in-memory binary stream for writing the Excel file
        output = io.BytesIO()

        # Create a new Excel workbook and add a worksheet
        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        # Write the column headers
        for col_idx, col_name in enumerate(columns):
            # Calculate the maximum width needed for the column
            max_width = max(len(str(col_name)), max(len(str(row[col_idx])) for row in rows))
            # Set the column width to the maximum width plus a little extra space
            worksheet.set_column(col_idx, col_idx, max_width + 2)
            worksheet.write(0, col_idx, col_name)

        # Write the data rows
        for row_idx, row in enumerate(rows):
            for col_idx, cell_value in enumerate(row):
                worksheet.write(row_idx + 1, col_idx, cell_value)

        # Close the workbook
        workbook.close()

        # Move the stream cursor to the beginning
        output.seek(0)
        # Upload xlsx file to S3
        s3 = boto3.client('s3')
        s3.put_object(Bucket=s3_bucket, Key=temp_s3_key, Body=output, ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        s3_urls.append(f"https://{s3_bucket}.s3.amazonaws.com/{temp_s3_key}")
    print(s3_urls)
    return s3_urls

def sendReportEmail(to,subject,s3_urls):
    # Create a MIME multipart message
    message = MIMEMultipart('alternative')
    message['Subject'] = subject
    message['From'] = SENDER_EMAIL
    message['To'] = to
    # Create an HTML table from the s3_urls array
    table_rows = "".join([f"<tr><td>{url}</td></tr>" for url in s3_urls])

    # HTML code with the table
    html_body = f"""
    <html>
        <head>
            <style>
                table {{
                    border-collapse: collapse;
                    border: 1px solid black;
                }}
                th, td {{
                    border: 1px solid black;
                    padding: 8px;
                }}
            </style>
        </head>
        <body>
            <p>Here is/are the S3 URL(s) of Gap Report(s):</p>
            <table>
                {table_rows}
            </table>
        </body>
    </html>
    """
    # Attach the HTML part to the message
    message.attach(MIMEText(html_body, 'html'))
    recipient_email_list = to.split(',')
    body = {
        "receipient": recipient_email_list,
        "data": message.as_string()
    }
    send_sqs_message(EMAIL_QUEUE_URL,body)
    print('Email Sent!',to,subject)

def fetchRecord(job_status,job_name):
    # Connect to Redshift
    conn = Connection(DB_HOST,DB_PORT,DB_NAME,DB_USER,DB_PASSWORD) 
    query = f"SELECT * FROM {GAP_REDSHIFT_TABLE} where job_status={job_status} and job_name='{job_name}'"
    # Open a cursor to perform database operations
    cur = conn.cursor()
    cur.execute(query)
    rows = cur.fetchall()
    # Convert rows to JSON
    result = {}
    for row in rows:
        result=dict(zip([column[0] for column in cur.description], row))
    # Fetch all results
    # results = cur.fetchall()

    # Commit the results
    conn.commit()

    # Close the cursor and connection
    cur.close()
    conn.close()
    return result

def fetchClientConfig(tenant_name):
    # Connect to Redshift
    conn = Connection(DB_HOST,DB_PORT,DB_NAME,DB_USER,DB_PASSWORD) 
    query = f"SELECT * FROM {CLIENT_CONFIG_REDSHIFT_TABLE} where tenant_name='{tenant_name}'"
    # Open a cursor to perform database operations
    cur = conn.cursor()
    cur.execute(query)
    rows = cur.fetchall()
    # Convert rows to JSON
    result = {}
    for row in rows:
        result=dict(zip([column[0] for column in cur.description], row))
    # Fetch all results
    # results = cur.fetchall()

    # Commit the results
    conn.commit()

    # Close the cursor and connection
    cur.close()
    conn.close()
    return result

def fetchExportTableNames(clientConfig,export_table,job_name):
    # Connect to Redshift
    conn = Connection(clientConfig['host_name'],clientConfig['port'],clientConfig['db_name'],clientConfig['user_name'],clientConfig['password']) 
    query = f"SELECT * FROM {export_table}"
    # Open a cursor to perform database operations
    cur = conn.cursor()
    cur.execute(query)
    rows = cur.fetchall()
    # Convert rows to JSON
    result = {}
    for row in rows:
        result=dict(zip([column[0] for column in cur.description], row))
    # Fetch all results
    # results = cur.fetchall()

    # Commit the results
    conn.commit()

    # Close the cursor and connection
    cur.close()
    conn.close()
    final_result = [] if result['table_names'] is None else result['table_names'].split(',')
    return final_result

def send_sqs_message(queue_url,body):
    print(queue_url,body)
    sqs = boto3.client('sqs')
    messageGroupId = str(uuid.uuid4())
    response = sqs.send_message(
        QueueUrl=queue_url,
        MessageBody=json.dumps(body),
        MessageGroupId=messageGroupId, 
        MessageDeduplicationId = messageGroupId
    )

    return response
    
def lambda_handler(event, context):
    try:
        body = event["Records"][0]['body']
        if isinstance(body, str):
            body = json.loads(body)
        if((body['eventType']=='job.run.completed' and body['data']['runStatus']=='Errored') or (body['eventType']=='job.run.errored')):
            status = False
            job_name=body['data']['jobName']
            res = fetchRecord(status,job_name)
            # if there is no record for job name
            if(not res):
                return
            clientConfig = fetchClientConfig(res['tenant_name'])
            tables = fetchExportTableNames(clientConfig,res['export_table'],job_name)
            clientConn = Connection(clientConfig['host_name'],clientConfig['port'],clientConfig['db_name'],clientConfig['user_name'],clientConfig['password']) 
            urls = generateReportandUploadToS3(clientConn,tables,res['s3_bucket_name'],res['s3_key_prefix'])
            clientConn.close()
            subject = f'{job_name} Gap Report Generated Successfully!'
            sendReportEmail(res['email_receipients'],subject,urls)
            return urls
        elif(body['eventType']=='job.run.completed' and body['data']['runStatus']=='Success'):
            status = True
            job_name=body['data']['jobName']
            res = fetchRecord(status,job_name)
            # if there is no record for job name
            if(not res):
                return
            clientConfig = fetchClientConfig(res['tenant_name'])
            tables = fetchExportTableNames(clientConfig,res['export_table'],job_name)
            clientConn = Connection(clientConfig['host_name'],clientConfig['port'],clientConfig['db_name'],clientConfig['user_name'],clientConfig['password']) 
            urls = generateReportandUploadToS3(clientConn,tables,res['s3_bucket_name'],res['s3_key_prefix'])
            clientConn.close()
            subject = f'{job_name} Gap Report Generated Successfully!'
            sendReportEmail(res['email_receipients'],subject,urls)
            return urls
        else:
            return
    except Exception as e: 
        print(e)
    return
