# # https://www.whocc.no/atc_ddd_index/

from time import sleep
import time
from bs4 import BeautifulSoup
import requests
import pandas as pd
import string 
import random 
import os
os.makedirs('Final Output', exist_ok=True)

Ultlimate_Records=[]
def Printing(record):
    
    return [record[0],
    record[1],
    record[2],
    record[4],
    record[5],
    record[7],
    record[8],
    record[10],
    record[11],
    record[12],
    record[13],
    record[14],
    record[15],
    record[16]]
r=requests.Session()

for i in string.ascii_uppercase:
    print(f"Running -> {i}")
    base_url=f'https://www.whocc.no/atc_ddd_index/?code={i}&showdescription=no'
    response=r.get(base_url)
    soup=BeautifulSoup(response.text,"lxml")
    
    ## First Tier 
    
    if soup.find('h2')==None: 
        Records1=[]
        atc_code_first=soup.select('#content b a')[0]['href'].split('&')[0].split('=')[-1]
        name_atc=soup.select('#content b a')[0].text

        for tags in soup.select('#content b a'):
            if "=no" in tags['href']:
                atc_code=tags['href'].split('&')[0].split('=')[-1]
                Records1.append([atc_code_first, name_atc, atc_code, 'https://www.whocc.no/atc_ddd_index'+tags['href'].split('.')[1],tags.text])
    else:
        continue
    
    time.sleep(2)
    
    ## Second Tier 
    
    new_records=[]
    for record in Records1:
        atc_check=record[2]
        url=record[-2]
        name=record[-1]

        response=r.get(url)
        soup=BeautifulSoup(response.text,"lxml")
        for tags in soup.find('div',{"id":"content"})('a'):
            if "=no" in tags['href']:
                atc_code=tags['href'].split('&')[0].split('=')[-1]

                if atc_check==atc_code:
                    continue

                new_records.append(record+[atc_code, 'https://www.whocc.no/atc_ddd_index'+tags['href'].split('.')[1], tags.text])

    time.sleep(2)
    
    ## Third Tier 
    Third_Records=[]
    for record in new_records:
        atc_check1=record[2]
        atc_check2=record[5]
        url=record[-2]
        name=record[-1]

        response=r.get(url)
        soup=BeautifulSoup(response.text,"lxml")

        soup=BeautifulSoup(response.text,"lxml")
        for tags in soup.find('div',{"id":"content"})('a'):
            if "=no" in tags['href']:
                atc_code=tags['href'].split('&')[0].split('=')[-1]
                if atc_code==atc_check1 or atc_code==atc_check2:
                    continue
                Third_Records.append(record+[atc_code, 'https://www.whocc.no/atc_ddd_index'+tags['href'].split('.')[1],tags.text])
    
    time.sleep(2)   
    
    ############### Last Tier ###################### Ultimate Records ############################
    for record in Third_Records:
        url=record[-2]
        response=r.get(url)
        soup=BeautifulSoup(response.text,"lxml")
        atc_code_last=""
        name=""
        ddd=""
        u=""
        adm_r=""
        note=""

        try:  ## if any error occurs 
            tables=soup.find('table')('tr')
            for i in range(1, len(tables)):
                row=tables[i]('td')

                atc_code_last=row[0].text.replace('\xa0',"")
                name=row[1].text.replace('\xa0',"")
                ddd=row[2].text.replace('\xa0',"")
                u=row[3].text.replace('\xa0',"")
                adm_r=row[4].text.replace('\xa0',"")
                note= row[5].text.replace('\xa0',"")

                Ultlimate_Records.append(record+[atc_code_last, name, ddd, u, adm_r, note])
        except:
            Ultlimate_Records.append(record+[atc_code_last, name, ddd, u, adm_r, note])

        if random.randint(1,10)%3==0:
            time.sleep(1)
# Final Output - Master Sheet
Records=[]
for record in Ultlimate_Records:
    Records.append(Printing(record))

# file=pd.DataFrame(Records,columns=['sp_atc_code','sp_atc_description','sp_atc_code_l1','l1_name','sp_atc_code_l2','l2_name','sp_atc_code_l3','l3_name','sp_atc_code_l4','l4_name','DDD','U','Adm.R',' Note'])
file=pd.DataFrame(Records,columns=['web_atc_code_l1','web_l1_name','web_atc_code_l2','web_l2_name','web_atc_code_l3','web_l3_name','web_atc_code_l4','web_l4_name','web_atc_code_l5','web_l5_name','DDD','U','Adm.R','Note'])
file.to_csv('Final Output/ATC Main Output File.csv', index=False, encoding="utf-8-sig")

print("Completed")