import zipfile
import os
import pandas as pd
import xlrd
import re
import requests
import time
import shutil

final_output_path = 'Final Output'
os.makedirs(final_output_path, exist_ok=True)


zip_download_directory = 'download_directory_zip_files'
zip_file_path = f'{zip_download_directory}/ndctext.zip'
extract_to_directory = f'{zip_download_directory}/extracted_files'

# removing - zip_download_directory if available | Removing any Previous Files
if zip_download_directory in os.listdir():
    shutil.rmtree(zip_download_directory)

os.makedirs(zip_download_directory, exist_ok=True)
## Function to download Zip File: https://www.accessdata.fda.gov/cder/ndctext.zip
def download_file(url, directory):
    filename = url.split('/')[-1]
    file_path = os.path.join(directory, filename)

    # Send an HTTP GET request to the URL
    response = requests.get(url, stream=True)

    # Check if the request was successful
    if response.status_code == 200:
        with open(file_path, 'wb') as file:
            # Write the content of the response to the file
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
        print(f"File downloaded successfully to: {file_path}")
    else:
        print(f"Failed to download the file. Status code: {response.status_code}")

## Downloading Txt File ## 
downloading_zip_file = download_file('https://www.accessdata.fda.gov/cder/ndctext.zip', zip_download_directory)

## Extracting the Zip File  ## 
os.makedirs(extract_to_directory, exist_ok=True)
with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
    zip_ref.extractall(extract_to_directory)
    
list_of_txt_files = os.listdir(extract_to_directory)
print(f'Total Txt Found: {len(list_of_txt_files)}')

## Reading TXT files: Two files | Two DataFrames
product_file_path = f'{extract_to_directory}\product.txt'
package_file_path = f'{extract_to_directory}\package.txt'

# Assuming your text file is tab-separated (use delimiter='\t' for tab-separated files)
df1 = pd.read_csv(product_file_path, delimiter='\t', encoding= 'unicode_escape')
df1 = df1.fillna('')
df1.to_excel(f'{final_output_path}/product.xlsx', index=False)

df2 = pd.read_csv(package_file_path, delimiter='\t')
df2 = df2.fillna('')
## Working on previous file
def extract_and_process_codes(df):
    '''Description: The code extracts and processes codes from the 'PACKAGEDESCRIPTION' column, creates a new DataFrame by exploding the codes, 
    and returns a cleaned and modified DataFrame.'''
    df = df.fillna('')
    df['matched_pattern'] = ''
    def getting_codes(search_string):
        codes_ = (', ').join(i.group() for i in re.finditer(f'\(([0-9-])+\)', search_string))
        return codes_

    df['matched_pattern'] = df.PACKAGEDESCRIPTION.apply(getting_codes)

    df['new_code'] =''

    for i in range(len(df)):
        new_code = df['matched_pattern'].iloc[i].replace(df['NDCPACKAGECODE'].iloc[i], '').replace('(','').replace(')','').replace(',','').strip()
        df.loc[i, 'new_code'] = new_code

    ### Creating a new dataframe (wp) and doing changes accordingly | Final Dataframe: nf
    wp = df[df['new_code']!='']

    def converting_to_list(x):
        x = x.split()
        return x 

    ndc_package_codes = wp['new_code'].apply(converting_to_list)
    wp.loc[:, 'NDCPACKAGECODE'] = ndc_package_codes

    wp = wp.explode('NDCPACKAGECODE').reset_index(drop=True)

    ## nf dataframe ##
    nf = pd.concat([df, wp], axis=0)
    nf.drop(['matched_pattern', 'new_code'], axis=1, inplace=True)
    nf.reset_index(drop=True, inplace=True)
    nf['NDCPACKAGECODE'] = nf['NDCPACKAGECODE'].str.strip()
    
    return nf
print(f'Processing File: package.txt')
df2 = extract_and_process_codes(df2)
df2.to_excel(f'{final_output_path}/package.xlsx', index=False)

# removing - zip_download_directory if available | Removing any Previous Files
if zip_download_directory in os.listdir():
    shutil.rmtree(zip_download_directory)

