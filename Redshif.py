
import sys
from awsglue.transforms import *
from awsglue.utils import getResolvedOptions
from pyspark.context import SparkContext
from awsglue.context import GlueContext
from awsglue.job import Job
from awsglue.dynamicframe import DynamicFrame
import sys
from pyspark.context import SparkContext
from awsglue.context import GlueContext
from awsglue.job import Job
from awsglue.utils import getResolvedOptions
import json
import time
from re import sub
import re
import functools
from functools import reduce
from pyspark.sql.functions import col, trim, lit, length, regexp_extract, to_timestamp
import datetime
import boto3
import uuid
import json


sc = SparkContext.getOrCreate()
glueContext = GlueContext(sc)
spark = glueContext.spark_session

job = Job(glueContext)
QUEUE_URL = 'https://sqs.us-east-2.amazonaws.com/339834084992/scus-glue-etl-prod-retention'

LOG_QUEUE_URL = 'https://sqs.us-east-2.amazonaws.com/339834084992/scus-etl-glue-logging.fifo'
LOG_TABLE = "import.glue_etl_logs"

JDBC_URL = "jdbc:redshift://reds-learning.cbrszyzljcmo.us-east-2.redshift.amazonaws.com:"
DRIVER = "com.amazon.redshift.jdbc.Driver"
CONFIG_TABLE = "import.config_ahp"
USER = "Your_userName"
PASSWORD = "Your_password"

EMAIL_CONFIG_TABLE = "import.glue_etl_email"

REDSHIFT_IAM_ASSOCIATED_ROLE = "arn:aws:iam::339834084992:role/myRedshiftRole"

GLUE_CONNECTION='scus-glue-connection1'

SDC_CRON_BATCHED_AT = datetime.datetime.now()
SDC_CRON_HASH = str(uuid.uuid4())
print("Starting job at",SDC_CRON_BATCHED_AT, "with hash", SDC_CRON_HASH)
### 
class Logger:
    def __init__(self, queueUrl, logTable, cronHash):
        self._sqs = boto3.client('sqs')
        self._queueUrl = queueUrl
        self._logTable = logTable
        self._cronHash = cronHash
        
    def _sendMessage(self, body):
        messageGroupId = str(uuid.uuid4())
        response = self._sqs.send_message(QueueUrl= self._queueUrl, MessageBody=body, MessageGroupId=messageGroupId, MessageDeduplicationId = messageGroupId)
        print(f"Message sent. Message ID: {response['MessageId']}")
    
    def logStart(self):
        query = f"Insert into {self._logTable} values (GETDATE(), default, default, default, default, default, \'RUNNING\',\'{self._cronHash}\', default, default, \'STARTED\')"
        print("Logging Query", query)
        self._sendMessage(query)
        
    def update(self,params):
        setFields = 'set'

        for key in params:
          if(isinstance(params[key],str)):
            setFields += f' {key}=\'{params[key]}\','
          else:
            setFields += f' {key}={params[key]},'

        setFields=setFields[:-1]
        
        query = f"Update {self._logTable} {setFields} where unique_job_hash=\'{self._cronHash}\'"
        
        print(query)
        self._sendMessage(query)
    
    def complete(self):
        query = f"Update {self._logTable} set processing_finished_at=GETDATE() where unique_job_hash=\'{self._cronHash}\'"
        print(query)
        self._sendMessage(query)
        
        
        
class SqsHandler:
    def __init__(self, queueUrl):
        self._sqs = boto3.client('sqs')
        self._receiptHandle = ''
        self._messageBody = {}
        self._queueUrl = queueUrl
        self._maxMessages = 1
        self._visibilityTimeout = 300
        
    def receiveMessage(self):
        response = self._sqs.receive_message(
            QueueUrl=self._queueUrl,
            MaxNumberOfMessages=1,
            VisibilityTimeout=self._visibilityTimeout,
            WaitTimeSeconds=20
        )
        
        
        if(not 'Messages' in response):
            print("Sqs queue is empty")
            raise Exception("No messages in Sqs queue", self._queueUrl)

        message = response['Messages'][0]
        
        print("Received message",message, "from sqs queue", self._queueUrl)
        
        self._messageBody = json.loads(message['Body'])
        self._receiptHandle = message['ReceiptHandle']
    
    def deleteMessage(self):        
        self._sqs.delete_message(
            QueueUrl=self._queueUrl,
            ReceiptHandle=self._receiptHandle
        )
        print("Delete message from sqs queue", self._queueUrl)

    def getS3PathFromMessage(self):
        return 's3://' + self.getBucketName() + '/' + self.getObjectKey()
    
    def getS3OutputPathFromMessage(self):
        s3UnprocessedFolder = self.getS3Folder()
        partitionPath=self.getPartitionOutputPath()
        return s3UnprocessedFolder.replace('up_files','p_files')+partitionPath+"/"+str(SDC_CRON_HASH), s3UnprocessedFolder.replace('up_files','e_files')+partitionPath+ "/"+str(SDC_CRON_HASH)
    
    def getPartitionOutputPath(self):
        filename=self.getS3Filename()
        year=filename.split('_')[-1].split('.')[0]
        month=filename.split('_')[-2]
        return year+"/"+month
    

    
    def ValidateFileFormate(self,filename_format):
        filename=self.getS3Filename()
        if (filename_format==None):
            return True
        filename_format = json.loads(filename_format)[0]
        checkYear=filename.split('_')[-1].split(".")[0]
        print("year=",checkYear)
        if (checkYear!='2023'):
            print("invalid filname year")
            return False
        fileNameLength = len(filename.split('_')[-1])+len(filename.split('_')[-2])+2
        filename=filename[0: len(filename) - fileNameLength] 
        print("filename,filename_format=",filename,filename_format)
        if (filename==filename_format): 
            return True
        else: 
            return False
        
    
    def getS3Filename(self):
        s3Path = self.getS3PathFromMessage()
        fileNameLength = len(s3Path.split('/')[-1])
        return s3Path[len(s3Path) - fileNameLength:len(s3Path)]
    
    
    def getS3Folder(self):
        s3Path = self.getS3PathFromMessage()
        fileNameLength = len(s3Path.split('/')[-1])
        return s3Path[0: len(s3Path) - fileNameLength]       

    def getBucketName(self):
        return self._messageBody['Records'][0]['s3']['bucket']['name']

    def getObjectKey(self):
        return self._messageBody['Records'][0]['s3']['object']['key']
    
class S3Handler:
    def __init__(self, glueContext):
        self._glueContext = glueContext
        
    def readDynamicFrameFromS3(self, glue_context, path, fileFormat, fileSeparator):
        dynamicframe = glue_context.create_dynamic_frame.from_options(
            connection_type='s3',
            connection_options={
                'paths': [path],
                'recurse': True
            },
            format=fileFormat,
            format_options={"separator":fileSeparator , "withHeader": True}
        )
        return dynamicframe 
    
    def writeDataFrameToS3(self, df, outputPath):
        print("Writing dataframe to", outputPath)
        df.write.option("header","true").csv(outputPath)
        
    def deleteObject(bucketName, fileKey):
        s3 = boto3.client('s3')
        s3.delete_object(Bucket=bucketName, Key=fileKey)
        
        

class DataframeFormatter:    
    
    def renameAndDropColumns(self, df, columnConfig):
        dfColumns = df.columns
        
        for column in dfColumns:
            
            if column in columnConfig:
                df = df.withColumnRenamed(column, columnConfig[column]['table_column'])
            else:
                df = df.drop(column)
        return df
            
            
    
    def addNewColumns(self, df, objectKey):
        df = df.withColumn('sdc_cron_batched_at', lit(SDC_CRON_BATCHED_AT))
        df = df.withColumn('sdc_cron_hash_key', lit(SDC_CRON_HASH))
        df = df.withColumn('sdc_cron_filename', lit(objectKey))
        df = df.withColumn('sdc_updated_by', lit("glue"))
        return df
    
    def getNonBlankColumns(self, columnConfig):
        nonBlankColumns = []
        for column in columnConfig:
            if(column == 'is_test'):
                continue
            if(columnConfig[column]['can_be_blank'] == 'false'):
                nonBlankColumns.append(columnConfig[column]['table_column'])
                
        return nonBlankColumns
    
    def filterInvalidBlankRows(self, df, columnConfig):
        nonBlankColumns = dfFormatter.getNonBlankColumns(columnConfig)

        print("Non blank columns",nonBlankColumns)

        errorConditions = reduce(lambda a, b: a | b, [
            (col(c).isNull() | (trim(col(c)) == ''))
            for c in nonBlankColumns
        ])

        validCondition = reduce(lambda a, b: a & b, [
            (~col(c).isNull() & (trim(col(c)) != ''))
            for c in nonBlankColumns
        ])

        errorDf = df.filter(errorConditions)
        df = df.filter(validCondition)
        
        return df, errorDf
    def filterInvalidDataTypeRows(self, df, columnConfig):
        schema = df.schema

        # Create an empty dataframe with the schema of the original dataframe
        invalidDataTypeRows = spark.createDataFrame(spark.sparkContext.emptyRDD(), schema)

        for column in columnConfig:
            # Skipping is test column 
            if(column == 'is_test'):
                continue
            
            # print(column)
            dataType = columnConfig[column]['type']
            
            print(columnConfig[column]['table_column'],dataType)

            if(dataType == 'string'):
                maxLength = columnConfig[column]['maxLength']
            
                errorResults = df.filter(length(columnConfig[column]['table_column']) > maxLength)
                df = df.filter(length(columnConfig[column]['table_column']) <= maxLength)               
                invalidDataTypeRows = invalidDataTypeRows.union(errorResults)
            elif(dataType == 'float' or dataType == 'int'):
                errorResults = df.filter(regexp_extract(col(columnConfig[column]['table_column']), r'^\d+(\.\d+)?$', 0) == '')
                df = df.filter(regexp_extract(col(columnConfig[column]['table_column']), r'^\d+(\.\d+)?$', 0) != '')
                invalidDataTypeRows = invalidDataTypeRows.union(errorResults)
            else:
                # Filtering datetime format
                print('timestamp format',columnConfig[column]['format'])
                errorResults = df.filter(to_timestamp(col(columnConfig[column]['table_column']), columnConfig[column]['format']).isNull()) 
                df = df.filter(to_timestamp(col(columnConfig[column]['table_column']), columnConfig[column]['format']).isNotNull())
                invalidDataTypeRows = invalidDataTypeRows.union(errorResults)
                
                
        return df, invalidDataTypeRows
        
    
    
    def validateMandatoryColumns(self, df, columnConfig):
        dfColumns = set(df.columns)
        
        for column in columnConfig.keys():
            if( columnConfig[column]['can_be_absent'] == 'false' ) and columnConfig[column]['table_column'] not in dfColumns:
                return False
        
        return True        
    
    
class SparkRedshiftHandler:
    def __init__(self, spark, jdbcUrl, driver, user, password):
        self.jdbcUrl = jdbcUrl
        self.driver = driver
        self.user = user
        self.password = password
        
        
    def readFromRedshift(self, table):
        return spark.read.format("jdbc").option("url", self.jdbcUrl) \
            .option("driver", self.driver) \
            .option("dbtable", table) \
            .option("user", self.user) \
            .option("password", self.password) \
            .load()  
logger = Logger(LOG_QUEUE_URL, LOG_TABLE, SDC_CRON_HASH)
logger.logStart()
from pyspark.sql.functions import regexp_replace, col

try:
    sqsHandler = SqsHandler(QUEUE_URL)
    sqsHandler.receiveMessage()
    s3Path = sqsHandler.getS3PathFromMessage()
    print("s3 path",s3Path)

    logger.update({"last_successful_job_stage":"FETCHED_SQS_MESSAGE", "file_path":sqsHandler.getS3PathFromMessage()})

    sparkRedshiftHandler = SparkRedshiftHandler(spark, JDBC_URL, DRIVER, USER, PASSWORD)

    print("fetching config from redshift")
    config = sparkRedshiftHandler.readFromRedshift(CONFIG_TABLE)
    config = config.toPandas()
  

    s3Folder = sqsHandler.getS3Folder()
    print("s3 folder for file", s3Folder)

    datasetConfig = config.loc[config['s3_folder_name'] == s3Folder].iloc[0]

    
    fileValidation=sqsHandler.ValidateFileFormate(datasetConfig["filename_format"])
    if (fileValidation==False):
      print("Invalid filename")
      logger.update({"processing_status":"Invalid filename" , "upload_table_name":datasetConfig['data_table'] , "upload_schema_name":datasetConfig['data_schema']})
      sqsHandler.deleteMessage()
      raise Exception('Invalid filename')


    columnConfig = json.loads(datasetConfig['column_map'])
    print("column configs", columnConfig)

    logger.update({"last_successful_job_stage":"FETCHED_FILE_CONFIG"})

    delimiter = config['delimiter'][0]
    fileFormat = config['expected_format'][0][1:]

    print("delimiter", delimiter, "file format", fileFormat)

    s3Handler = S3Handler(glueContext)
    s3Path = sqsHandler.getS3PathFromMessage()
    ddf = s3Handler.readDynamicFrameFromS3(glueContext, s3Path , fileFormat ,delimiter)

    dfFormatter = DataframeFormatter()
    df = ddf.toDF()

    print("Previous column list",df.columns)
    if ("is_test" not in df.columns):
            print("====================adding is_test field in file")
            df = df.withColumn('is_test', lit(0))
            print(df.show())
   


    print("Updated column list",columnConfig)


    df = dfFormatter.renameAndDropColumns(df, columnConfig)

    df, invalidBlankColumns = dfFormatter.filterInvalidBlankRows(df, columnConfig)

    df, invalidDataTypeRows = dfFormatter.filterInvalidDataTypeRows(df, columnConfig)

    df = dfFormatter.addNewColumns(df, sqsHandler.getObjectKey())


    valid = df.count()
    invalidDataType = invalidDataTypeRows.count()
    invalidBlank = invalidBlankColumns.count()
    errorRows = invalidDataType + invalidBlank
    total = valid + errorRows

    print("Valid entries",valid, "Invalid data type entries", invalidDataTypeRows,  "Invalid blank type entries", invalidBlankColumns)

    logger.update({"last_successful_job_stage":"PROCESSED_DATAFRAME", "total_processed_rows":valid, "total_rows":total ,"error_rows":errorRows})

    processedS3Path, errorS3Path = sqsHandler.getS3OutputPathFromMessage()

    print("Writing processed files at", processedS3Path,"Writing errorS3Path at", errorS3Path)


    s3Handler.writeDataFrameToS3(df, processedS3Path)

    s3Handler.writeDataFrameToS3(invalidDataTypeRows, errorS3Path+'_invalidDataTypes')

    s3Handler.writeDataFrameToS3(invalidBlankColumns, errorS3Path+'_invalidBlank')

    logger.update({"last_successful_job_stage":"WRITTEN_DF_TO_S3"})

    print(datasetConfig)
    
    print("Writing to db")

    df = df.withColumn('sdc_uploaded_at',lit(datetime.datetime.now()))
    all_column_names = df.columns
    columns_for_replacement = [i for i in all_column_names]
    for i in columns_for_replacement:
       df = df.withColumn(i,regexp_replace(i, 'nan', ''))

    
    df.write.format("com.databricks.spark.redshift") \
        .mode(datasetConfig["upload_mechanism"]) \
        .option("url", JDBC_URL) \
        .option("database", datasetConfig['data_db']) \
        .option("dbtable", datasetConfig['data_schema']+'.'+datasetConfig['data_table']) \
        .option("user", USER) \
        .option("password", PASSWORD) \
        .option("aws_iam_role", REDSHIFT_IAM_ASSOCIATED_ROLE) \
        .option("tempdir","s3://ohio-glue-etl-bucket/glue_redshift_temp") \
        .save()


    logger.update({"last_successful_job_stage":"UPLOADED_TO_DB" , "upload_table_name":datasetConfig['data_table'] , "upload_schema_name":datasetConfig['data_schema']})
    
    sqsHandler.deleteMessage()
    
    
    processingStatus = "SUCCESS"
    
    if(errorRows != 0):
        processingStatus = "PARTIAL_SUCCESS"

    
    logger.update({"last_successful_job_stage":"COMPLETED", "processing_status":processingStatus})
    logger.complete()

except  Exception as e:
    print("Caught exception", e)
    logger.update({"processing_status":"FAILED"})
    logger.complete()
    
job.commit()
