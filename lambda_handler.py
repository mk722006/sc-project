# Imports
from global_report import global_report
from local_report import local_report

import psycopg2
from psycopg2.extras import RealDictCursor

# Establishing Connection

def Connection():
    conn = psycopg2.connect(host="uhc-prod.cbrszyzljcmo.us-east-2.redshift.amazonaws.com", database="scus_poc", port='5439', user="smahin", password="UHBwed@1726")  
    curr = conn.cursor()
    print("Connection Established")
    return conn, curr
def lambda_handler(event, context):

	conn, curr = Connection()
	# Input Fields
	# conn=''
	# curr=''
	
	# month = 4
	# year = 2023
	# selected_country = 'Chile'


	# ## Table Name = 'l_exec_report_data'
	# table_name = 'adhoc.l_exec_report_data'


	# markets_loop = ['Global','Brazil','Chile','Colombia', 'Peru']
	# if selected_country not in markets_loop:
	#     print(f'{selected_country} not found')
	# else:
	#     print(f'Running the code for - "{selected_country}"')
	    
	# if year != 2023:
	#     print(f'Year should be 2023 | Change the Input')
	# else:
	#     if selected_country == 'Global':
	#     	function_running = global_report(month, year, selected_country, table_name, conn, curr)
	#     else:
	#     	function_running = local_report(month, year, selected_country, table_name, conn, curr)
		
	return {
		'statusCode': 200,
		'body': json.dumps('Successful')
	    }        

