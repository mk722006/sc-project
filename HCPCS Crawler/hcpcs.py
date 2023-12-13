import zipfile
import os
import pandas as pd
import xlrd
import re
import requests
import time
import shutil
from bs4 import BeautifulSoup

final_output_path = 'Final Output'
zip_download_directory = 'download_directory_zip_files'
# Functions: getting_file_info | download_file
def download_file(url, directory, filename):
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
        
        
def getting_file_info(url, zip_download_directory):
    
    # removing - zip_download_directory if available | Removing any Previous Files
    if zip_download_directory in os.listdir():
        shutil.rmtree(zip_download_directory)
    
    os.makedirs(zip_download_directory, exist_ok=True)
    
    response = requests.get(url, stream=True)
    soup = BeautifulSoup(response.text, 'lxml')

    file_info = []
    file_links = soup.select('a[data-entity-type="media"]')[0:3]
    for file in file_links:
        title = file.text.strip()
        link = 'https://www.cms.gov'+file['href']
        file_name = re.sub(r'[^\w\s-]', '', title).strip().replace(' ', '_')
        file_name = file_name +'.zip'
        file_info.append([title, link, file_name])
    
    ## Downloading Zil Files ## 
    for i in range(len(file_info)):
        downloading_zip_file = download_file(file_info[i][1], zip_download_directory, file_info[i][2])
        
    return file_info
# Running - getting_file_info | Downloading all Zip files (first 3)
file_info = getting_file_info('https://www.cms.gov/medicare/coding/hcpcsreleasecodesets/hcpcs-quarterly-update', zip_download_directory)
## Extracting Zip files into respective Folders | Final Output
if final_output_path in os.listdir():
    shutil.rmtree(final_output_path)
    time.sleep(2)
os.makedirs(final_output_path, exist_ok=True)


df = pd.DataFrame(file_info, columns = ['title', 'url','file_name'])
for i, file_path in enumerate(os.listdir(zip_download_directory)):
    if '.zip' in file_path:
        
        folder_path = f"{final_output_path}/{df[df['file_name'] == file_path]['file_name'].values[0]}".replace('.zip','')
        ## Creating a Folder | Extracting Zip Files in this Folder
        os.makedirs(folder_path, exist_ok=True)
        
        ## Extracting Zip Files
        with zipfile.ZipFile(f'{zip_download_directory}\{file_path}', 'r') as zip_ref:
            zip_ref.extractall(folder_path)
            
        ## Removing all the files except Excels (.xlsx)
        for file_available in os.listdir(folder_path):
            if '.xlsx' not in file_available:
                os.remove(f'{folder_path}/{file_available}')

## removing - zip_download_directory if available | Removing any Previous Files
if zip_download_directory in os.listdir():
    shutil.rmtree(zip_download_directory)

print(f'Completed Running')