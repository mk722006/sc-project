import json
import boto3
import pandas as pd
import csv

# Mapping strings to separators for txt files
TXT_SEPARATOR_MAP = {
    'comma': ',',
    'semicolon': ';'
}

# List of headers to be added
headers_to_add = [
    "Item ID", "Item Name", "Item Description", "Spend Category ID",
    "Base UOM", "GTIN UOM Base", "UPN UOM Base", "UOM1", "UOM1 Conversion Factor",
    "GTIN UOM 1", "UPN UOM 1", "UOM2", "UOM2 Conversion Factor", "GTIN UOM 2",
    "UPN UOM 2", "UOM3", "UOM3 Conversion Factor", "GTIN UOM 3", "UPN UOM 3",
    "UOM4", "UOM4 Conversion Factor", "GTIN UOM 4", "UPN UOM 4", "UOM5",
    "UOM5 Conversion Factor", "GTIN UOM 5", "UPN UOM 5", "UOM6", "UOM6 Conversion Factor",
    "GTIN UOM 6", "UPN UOM 6", "Manufacturer ID", "Manufacture Ref ID", "UNSPSC",
    "UNSPSC Description", "PMM Item Number", "Lawson Item Number", "CDM PMM",
    "CDM Lawson", "Epic Interfaced", "Latex", "LatexFree", "LatexUnknown",
    "Reusable", "Tissue Tracking", "HCPCS", "Contract Number", "Contract Description",
    "Contract Start Date", "Contract End Date", "Contract UOM", "Contract QOE",
    "Contract Price", "Vendor Name", "Vendor Code", "Vendor Item ID", "Reference ID",
    "Item Active", "Brand Name", "NDC"
]



def lambda_handler(event, context):
    #console.log(“event: “, util.inspect(event, { showHidden: false, depth: null }));
    # TODO implement
    print('__________', event)
    event = event["Records"][0]

    bucketName = event["s3"]["bucket"]["name"]
    objectKey = event["s3"]["object"]["key"]
    
    fileName = objectKey.split("/")[-1]

        
        
    localFilePath = "/tmp/"+fileName

    print("BucketName", bucketName, "objectKey",
          objectKey, "fileName", fileName)

    downloadS3File(bucketName, objectKey, localFilePath)
    
    convertedFilePath = convertToCsv(localFilePath, fileName)

    uploadKey = objectKey[:len(objectKey) - len(fileName)].replace("/raw_files/", "/up_files/") + \
        convertedFilePath.split("/")[-1]

    uploadFile(convertedFilePath, bucketName, uploadKey)

    return {
        'statusCode': 200,
        'body': json.dumps('Successful')
    }


def uploadFile(convertedFilePath, bucketName, key):
    try:
        # Create an S3 client
        s3 = boto3.client('s3')

        # Upload the file to S3
        # Specify the key/name of the uploaded file in S3
        s3.upload_file(convertedFilePath, bucketName, key)


        # Print a success message
        print(f"File uploaded successfully.")

    except Exception as e:
        # Print an error message
        print(f"Error occurred while uploading file: {str(e)}")


def downloadS3File(bucketName, objectKey, localFilePath):
    try:
        print("________________________________",bucketName, objectKey, localFilePath)
        # Create an S3 client
        s3 = boto3.client('s3')

        # Download the file from S3 to /tmp/

        s3.download_file(bucketName, objectKey, localFilePath)

        # Print a success message
        print(f"File '{localFilePath}' downloaded successfully.")
        
    except Exception as e:
        # Print an error message
        print(f"Error occurred while downloading file: {str(e)}")

def check_headers(df, headers_to_add):
    if headers_to_add[0] in    df.columns.values.tolist():
        return True
    else:
        return False


def convertToCsv(localFilePath, fileName):
    # Extract the file extension
    file_extension = fileName.split('.')[-1]
    print("passing filename", fileName)
    outputFilePath = "/tmp/"+fileName.split(".")[0]+".csv"
    sheetName = fileName.split('Items')[0]
    # Read Excel file based on file extension
    if file_extension == 'xlsx':
        

        df = pd.read_excel(localFilePath, sheet_name=sheetName)
        if fileName.find("FINT"  ) == 0:
            if check_headers(df, headers_to_add):
                pass
            else:
                df = pd.read_excel(localFilePath, sheet_name = sheetName, header = None )
                df.columns = headers_to_add
                
        # df.to_csv(outputFilePath, mode='w', header=True, index=True,quotechar='"',quoting=csv.QUOTE_NONE,escapechar='\\')
        
        values=df.values.tolist()

        with open(outputFilePath, 'w') as file:
          writer = csv.writer(file,quotechar='"', delimiter=',', quoting=csv.QUOTE_ALL, skipinitialspace=False)
          writer.writerow(df.columns)
          writer.writerows(values)

    if file_extension == 'csv':
        df = pd.read_csv(localFilePath)
        if fileName.find("FINT"  ) == 0:
            if check_headers(df, headers_to_add):
                pass
            else:
                df = pd.read_csv(localFilePath, header = None)
                df.columns = headers_to_add
        values=df.values.tolist()
        print("=======================================")
        with open(outputFilePath, 'w') as file:
          writer = csv.writer(file,quotechar='"', delimiter=',', quoting=csv.QUOTE_ALL, skipinitialspace=False)
          writer.writerow(df.columns)
          writer.writerows(values)
        
    elif file_extension == 'txt':
        separatorString = fileName.split('_')[-1].split('.')[0]
        separator = TXT_SEPARATOR_MAP[separatorString]
        df = pd.read_csv(localFilePath, sep=separator)
        df.to_csv(outputFilePath, mode='w', header=True, index=False)

    print("CSV writing completed successfully.")

    return outputFilePath
