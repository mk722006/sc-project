def global_report(month, year, selected_country, table_name, conn, curr):
    
    # Imports
    from openpyxl import load_workbook

    import pandas as pd
    from datetime import datetime, timedelta

    import shutil
    import os
    from dateutil.relativedelta import relativedelta


    # MTD, YTD - get_date_range | Dynamic Date Range - Function 

    def get_date_range(month, year):
        # Calculate the starting date range
        if month == 1:
            starting_date = datetime(year=year - 1, month=12, day=31)
        else:
            starting_date = datetime(year=year, month=month, day=1) - timedelta(days=1)

        # Calculate the ending date range
        if month == 12:
            ending_date = datetime(year=year + 1, month=1, day=1)
        else:
            ending_date = datetime(year=year, month=month + 1, day=1)

        # Format the dates as strings
        starting_date_str = starting_date.strftime('%Y-%m-%d')
        ending_date_str = ending_date.strftime('%Y-%m-%d')
        
        starting_date_ytd = datetime(year=year - 1, month=12, day=31)
        starting_date_ytd_str = starting_date_ytd.strftime('%Y-%m-%d')

        return starting_date_str, ending_date_str, starting_date_ytd_str, ending_date_str

    def convert_month_year(month, year):
        # Get the month abbreviation
        month_abbr = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        month_str = month_abbr[month]

        # Get the last two digits of the year
        year_str = str(year) #[-2:]

        # Combine the month abbreviation and year
        output = f"{month_str}-{year_str}"
        return output
    ## Important Variables
    start_date_mtd, end_date_mtd, start_date_ytd, end_date_ytd= get_date_range(month, year)

    ## All Market Info - market_df
    markets_loop = ['Global','Brazil','Chile','Colombia', 'Peru']
    currency_loop = ['USD','BRL','CLP','COP','PEN']
    curr_symbol_loop = ['$','R$','$','$','S/']
    market_df = pd.DataFrame({'country': markets_loop, 'currency': currency_loop, 'currency_symbol': curr_symbol_loop})


    ## Sheet Name, Currency Symbol, Country
    month_year_abbr = convert_month_year(month, year)
    sheet_name = f'{selected_country} Exe Sum' # Ex: Brazil Exe Sum
    curr_symbol = market_df[market_df['country']==selected_country]['currency'].values[0]
    print(f'Country: {selected_country} | Sheet Name: {sheet_name} | Symbol: {curr_symbol} | Month: {month_year_abbr}')


    print(f'Starting Date MTD: {start_date_mtd} | Ending Date MTD: {end_date_mtd}')
    print(f'Starting Date YTD: {start_date_ytd} | Ending Date YTD: {end_date_ytd}')
    # Making a Copy of the Template File - copy_wb
    file_name = f"executive_report-{selected_country}-{convert_month_year(month, year).split('-')[0]}-{year}.xlsx"

    if selected_country == 'Global':
        source_file_path = 'Template/Template - Global Exe Sum.xlsx'    
    else: # Brazil, Chile, Colombia, Peru
        source_file_path = 'Template/Template - Others.xlsx'
        
    # # Create a copy file path
    copy_directory = 'Copies'
    os.makedirs(copy_directory, exist_ok=True)
    copy_file_path = f'{copy_directory}/{file_name}'

    try: ## Removing the file if already available
        os.remove(copy_file_path)
    except:
        pass

    # # Copy the file
    # shutil.copy2(source_file_path, copy_file_path)

    # Open the copied file
    copy_wb = load_workbook(source_file_path)

    # if selected_country!= 'Global':
    #     for name in copy_wb.sheet_names:
    #         if name != sheet_name:
    #             copy_wb.sheets[name].delete()
    # Required Functions
    def calculate_previous_dates(month, year):
        # Create a datetime object for the first day of the given month and year
        date = datetime(year, month, 1)
        
        # Future date
        future_date = date - relativedelta(months=-1)
        current_date = date - relativedelta(months=0)
        one_month_prior = date - relativedelta(months=1)
        two_months_prior = date - relativedelta(months=2)
        three_months_prior = date - relativedelta(months=3)
        
        ### Month Names ####
        three_months_name = convert_month_year(three_months_prior.month, three_months_prior.year)
        two_months_name = convert_month_year(two_months_prior.month, two_months_prior.year)
        one_months_name = convert_month_year(one_month_prior.month, one_month_prior.year)
        current_month_name = convert_month_year(current_date.month, current_date.year)
        
        
        ## Start and End Date ##
        three_months_prior_start_date = three_months_prior - timedelta(days=1)
        three_months_prior_start_date = three_months_prior_start_date.strftime('%Y-%m-%d')
        three_months_prior_end_date = two_months_prior.strftime('%Y-%m-%d')

        two_months_prior_start_date = two_months_prior - timedelta(days=1)
        two_months_prior_start_date = two_months_prior_start_date.strftime('%Y-%m-%d')
        two_months_prior_end_date = one_month_prior.strftime('%Y-%m-%d')


        one_month_prior_start_date = one_month_prior - timedelta(days=1)
        one_month_prior_start_date = one_month_prior_start_date.strftime('%Y-%m-%d')
        one_month_prior_end_date = current_date.strftime('%Y-%m-%d')

        current_date_start_date = current_date - timedelta(days=1)
        current_date_start_date = current_date_start_date.strftime('%Y-%m-%d')
        current_date_end_date = future_date.strftime('%Y-%m-%d')

        return [[three_months_prior_start_date, three_months_prior_end_date, three_months_name], [two_months_prior_start_date, two_months_prior_end_date, two_months_name], [one_month_prior_start_date, one_month_prior_end_date, one_months_name], [current_date_start_date, current_date_end_date, current_month_name]]

    def getting_all_others_info(dictionary_info):
        df = pd.DataFrame(dictionary_info)
        percent_100 = (df[df.columns[1]].sum() / (df[df.columns[2]].sum())) * 100
        all_others = round(percent_100 - df[df.columns[1]].sum(), 2)
        all_others_percentage = round((100 - df[df.columns[2]].sum())/100, 3)
        return all_others, all_others_percentage

    def spend_percentage_value_function(dictionary_data, column_for_name, column_spend, column_percentage, copy_wb, cell_number, sheet_name, run):
        
        if column_for_name!=False: ## Set to False when you do not have column_for_name
            cell_name_A = f'{column_for_name}{cell_number}'
            cell_value_spend = list(dictionary_data[run].values())[0] # Name 
            copy_wb[sheet_name][cell_name_A].value = cell_value_spend
            

        if column_spend!= False:
        
            cell_name_B = f'{column_spend}{cell_number}'
            cell_value_spend = list(dictionary_data[run].values())[1] # Spend Value 
            copy_wb[sheet_name][cell_name_B].value = round(cell_value_spend, 2)
        
        if column_percentage!= False:
            cell_name_C = f'{column_percentage}{cell_number}'
            cell_value_percentage = round(list(dictionary_data[run].values())[2]/100, 3) # Percentage Value | Dividing by 100
            copy_wb[sheet_name][cell_name_C].value = cell_value_percentage
            

    three_months_prior, two_months_prior, one_month_prior, current_date = calculate_previous_dates(month, year)
    print(three_months_prior, two_months_prior, one_month_prior, current_date )
    # =_=_==_=_==_=_==_=_==_=_= {0} Some Headers - Global =_=_==_=_==_=_==_=_==_=_=
    ## A2 | Ex: Global Supply Chain Analytics - Apr 2023
    copy_wb[sheet_name]['A2'].value = f'Global Supply Chain Analytics - {month_year_abbr.split("-")[0]} 2023'

    ## B6, D6 | Ex: USD Spend MTD
    copy_wb[sheet_name]['B6'].value = f'{curr_symbol} Spend MTD'
    copy_wb[sheet_name]['D6'].value = f'{curr_symbol} Spend YTD'

    ## MTD Currency Headers: E62, E113, E167 | Ex: MTD USD Spend
    mtd_currency_spend_headers = ['E62', 'E113', 'E167']
    for head in mtd_currency_spend_headers:
        copy_wb[sheet_name][head].value = f'MTD {curr_symbol} Spend'
        
    ## YTD Currency Headers: F62, F113, F167 | Ex: YTD USD Spend
    ytd_currency_spend_headers = ['F62', 'F113', 'F167']
    for head in ytd_currency_spend_headers:
        copy_wb[sheet_name][head].value = f'YTD {curr_symbol} Spend'

    ## Headers: B86, E86, H86, K86, B101, E101, H101, K101, B139, E139, H139, K139, B155, E155, H155, K155, B179, E179, H179, K179, B194, E194, H194, K194 | Ex: USD
    headers_currency_list = ['B86','E86','H86','K86','B101','E101','H101','K101','B139','E139','H139','K139','B155','E155','H155','K155','B179','E179','H179','K179','B194','E194','H194','K194']
    for head in headers_currency_list:
        copy_wb[sheet_name][head].value = curr_symbol
        
    # copy_wb.save()
    # =_=_==_=_==_=_==_=_==_=_= {1} Executive Summary  [ROW-4] =_=_==_=_==_=_==_=_==_=_=

    ### **{1.1} USD Spend MTD [ROW-6]**

    ### **{1.2} USD Spend YTD [ROW-6]**

    ### **{1.3} Unique Suppliers [ROW-6]**

    ### **{1.4} Unique Manufacturers [ROW-6]**

    ### **{1.5} Unique SKUs [ROW-6]**
    def convert_to_millions_or_billions(value, selected_country, market_df):
        currency_to_use = market_df[market_df['country']==selected_country]['currency_symbol'].values[0]
        suffixes = ['', 'K', 'M', 'B', 'T']  # Suffixes for Thousand, Million, Billion, Trillion, etc.
        for suffix in suffixes[1:]:
            value /= 1000.0
            if abs(value) < 1000.0:
                return f"{currency_to_use}{value:.2f}{suffix}"
        return f"{currency_to_use}{value:.2f}{suffixes[-1]}"

    def executive_summary_query_values_function(table_name, start_date_mtd, end_date_mtd, start_date_ytd, end_date_ytd, curr):

        ## query_1_1_usd_spend_mtd ## 
        query_1_1_usd_spend_mtd = f"""select sum (dollar_total_cost) as USD_Spend_MTD from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}'"""
        curr.execute(query_1_1_usd_spend_mtd)

        value_query_1_1_usd_spend_mtd = curr.fetchall()
        value_query_1_1_usd_spend_mtd = convert_to_millions_or_billions(list(value_query_1_1_usd_spend_mtd[0].values())[0], selected_country, market_df)
        
        
        # print(value_query_1_1_usd_spend_mtd)

        ## query_1_2_usd_spend_ytd ## 
        query_1_2_usd_spend_ytd = f"""select sum (dollar_total_cost) as USD_Spend_YTD from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}'"""
        curr.execute(query_1_2_usd_spend_ytd)

        value_query_1_2_usd_spend_ytd = curr.fetchall()
        value_query_1_2_usd_spend_ytd = convert_to_millions_or_billions(list(value_query_1_2_usd_spend_ytd[0].values())[0], selected_country, market_df)
        
        # print(value_query_1_2_usd_spend_ytd)

        ## query_1_3_unique_suppliers ## 
        query_1_3_unique_suppliers = f"""select COUNT(DISTINCT(distributor_normalized)) as unique_suppliers from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}'"""
        curr.execute(query_1_3_unique_suppliers)

        value_query_1_3_unique_suppliers = curr.fetchall()
        # print(value_query_1_3_unique_suppliers)

        ## query_1_4_unique_suppliers ## 
        query_1_4_unique_manufacturers = f"""select COUNT(DISTINCT( mnf_dashboard_half)) as unique_manufacturers from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}'"""
        curr.execute(query_1_4_unique_manufacturers)

        value_query_1_4_unique_manufacturers = curr.fetchall()
        # print(value_query_1_4_unique_manufacturers)

        ## query_1_5_unique_skus ##
        query_1_5_unique_skus = f"""select COUNT(DISTINCT(sc_uhg_id)) as unique_SKUs from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}'"""
        curr.execute(query_1_5_unique_skus)

        value_query_1_5_unique_skus = curr.fetchall()
        # print(value_query_1_5_unique_skus)
        
        # return value_query_1_1_usd_spend_mtd, value_query_1_2_usd_spend_ytd, value_query_1_3_unique_suppliers, value_query_1_4_unique_manufacturers, value_query_1_5_unique_skus
        return value_query_1_1_usd_spend_mtd, value_query_1_2_usd_spend_ytd, list(value_query_1_3_unique_suppliers[0].values())[0], list(value_query_1_4_unique_manufacturers[0].values())[0], list(value_query_1_5_unique_skus[0].values())[0]

    ######## Running Function ########
    executive_summary_values = executive_summary_query_values_function(table_name, start_date_mtd, end_date_mtd, start_date_ytd, end_date_ytd, curr)

    ## Writing the Values on Excel File ##
    cell_names_executive_summary = ['B7', 'D7', 'F7', 'H7', 'J7']
    for j, value in enumerate(executive_summary_values):
        cell_value = value
        copy_wb[sheet_name][cell_names_executive_summary[j]].value = cell_value
    # =_=_==_=_==_=_==_=_==_=_= {2} Total Purchase Order Spend [ROW-17] =_=_==_=_==_=_==_=_==_=_=
    ### Q: Total P.O. Spend - MTD	| B21-B24 | total_po_spend_mtd
    ## {2.1}Total P.O. Spend - MTD	[ROW-19]
    query_mtd = f'''select geography ,sum (dollar_total_cost) as spend_MTD ,(spend_MTD / SUM(spend_MTD) OVER ()) * 100 AS percentage_spend_MTD from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' GROUP BY 1 ORDER BY 1'''
    curr.execute(query_mtd)
    total_po_spend_mtd = curr.fetchall()


    ## {2.2} Total P.O. Spend - YTD	[ROW-19]
    query_ytd = f"select geography ,sum (dollar_total_cost) as spend_YTD ,(spend_YTD / SUM(spend_YTD) OVER ()) * 100 AS percentage_spend_YTD from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' GROUP BY 1 ORDER BY 1"
    curr.execute(query_ytd)
    total_po_spend_ytd = curr.fetchall()
    ### W: Total P.O. Spend - Adding Values
    #### B20, D20 - Same Input [Ex:  Apr-23 ( BRL )] | Row 21 to 24 | Column: B,C,D,E 
    if selected_country == 'Global': 
        sheet_name = f'{selected_country} Exe Sum' # Ex: Brazil Exe Sum
        curr_symbol = market_df[market_df['country']==selected_country]['currency'].values[0]
        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Total PO Spend')
        
        # Writing in Cells - B20, D20  | Same for all countries
        B20_cell_value =  f'{month_year_abbr} ( {curr_symbol} )' # B20 and D20 Cell Values are the same
        D20_cell_value =  f'{month_year_abbr} ( {curr_symbol} )' # B20 and D20 Cell Values are the same
        copy_wb[sheet_name]['B20'].value = B20_cell_value
        copy_wb[sheet_name]['D20'].value = D20_cell_value
        
        for run in range(len(market_df)-1):
            
            ####### Row 21 to 24 ######## Column: B,C,D,E 
            cell_number = 21 + run 
            
            ##### B:  total_po_spend_mtd | cell_value_spend | B-21,22,23,24
            cell_name_B = f'B{cell_number}'
            cell_value_spend = list(total_po_spend_mtd[run].values())[1] # Spend Value # B
            copy_wb[sheet_name][cell_name_B].value = round(cell_value_spend, 2)
            
            ##### C:  total_po_spend_mtd | cell_value_percentage | C-21,22,23,24
            cell_name_C = f'C{cell_number}'
            cell_value_percentage = round(list(total_po_spend_mtd[run].values())[2]/100, 3) # Percentage Value # C | Dividing by 100
            copy_wb[sheet_name][cell_name_C].value = cell_value_percentage
            
            ###### D total_po_spend_ytd | cell_value_spend | D-21,22,23,24
            cell_name_D = f'D{cell_number}'
            cell_value_spend_ytd = round(list(total_po_spend_ytd[run].values())[1], 2) # Spend Value # D
            copy_wb[sheet_name][cell_name_D].value = round(cell_value_spend_ytd, 2)
            
            ###### E total_po_spend_ytd | cell_value_percentage | E-21,22,23,24
            cell_name_E = f'E{cell_number}'
            cell_value_percentage_ytd = round(list(total_po_spend_ytd[run].values())[2]/100, 3) # Percentage Value # E | Dividing by 100
            copy_wb[sheet_name][cell_name_E].value = cell_value_percentage_ytd
    # =_=_==_=_==_=_==_=_==_=_= {3}Total Category Spend by Market [ROW -28] =_=_==_=_==_=_==_=_==_=_=
    # {3.1} Non-Pharma [ROW-30]
    query_3_1_non_pharma = f"""select geography ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query_3_1_non_pharma)

    total_category_spend_by_market_non_pharma = curr.fetchall()

    # {3.2} Pharma [ROW-30]
    query_3_2_pharma = f"""select geography ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query_3_2_pharma)

    total_category_spend_by_market_pharma = curr.fetchall()

    # {3.3} Indirect [ROW-30]
    query_3_3_indirect = f"""select geography ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 1"""
    curr.execute(query_3_3_indirect)

    total_category_spend_by_market_indirect = curr.fetchall()

    # {3.4} Total PO Spend	[ROW-30]
    query_3_4_total_po_spend = f"""select geography ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' GROUP BY 1 ORDER BY 1"""
    curr.execute(query_3_4_total_po_spend)

    total_category_spend_by_market_total_po_spend = curr.fetchall()
    if selected_country == 'Global': 
        
        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Total Category Spend by Market')
        
        # Writing in Cells - B31, D31, F31, H31 | Headers
        table_header_value = f'YTD {year} ( {curr_symbol} )'
        table_header_cells = ['B31', 'D31', 'F31', 'H31']
        for cell in table_header_cells:
            copy_wb[sheet_name][cell].value = table_header_value
        
        for run in range(len(market_df)-1):
            
            ####### Row 32 to 35 ######## 
            cell_number = 32 + run 
            
            ## B & C:  total_category_spend_by_market_non_pharma | cell_value_spend | cell_value_percentage | B-32,33,34,35 | C-32,33,34,35 
            BC_Cell_values = spend_percentage_value_function(total_category_spend_by_market_non_pharma, False, 'B', 'C', copy_wb, cell_number, sheet_name, run)
            
            ## D & E:  total_category_spend_by_market_pharma | cell_value_spend | cell_value_percentage | D-32,33,34,35 | E-32,33,34,35 
            DE_Cell_values = spend_percentage_value_function(total_category_spend_by_market_pharma, False, 'D', 'E', copy_wb, cell_number, sheet_name, run)
            
            ## F & G:  total_category_spend_by_market_indirect | cell_value_spend | cell_value_percentage | F-32,33,34,35 | G-32,33,34,35 
            FG_Cell_values = spend_percentage_value_function(total_category_spend_by_market_indirect, False, 'F', 'G', copy_wb, cell_number, sheet_name, run)
            
            ## H & I:  total_category_spend_by_market_total_po_spend | cell_value_spend | cell_value_percentage | H-32,33,34,35 | I-32,33,34,35 
            HI_Cell_values = spend_percentage_value_function(total_category_spend_by_market_total_po_spend, False, 'H', 'I', copy_wb, cell_number, sheet_name, run)
    # =_=_==_=_==_=_==_=_==_=_= {4} Top Manufacturers by Category Spend [ROW-48]  =_=_==_=_==_=_==_=_==_=_=
    # {4.1} Non-Pharma [ROW-30]
    query_4_1_non_pharma = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 2 desc limit 5"""
    curr.execute(query_4_1_non_pharma)

    top_manufacturers_by_category_spend_non_pharma = curr.fetchall()

    # {4.2} Pharma [ROW-50]
    query_4_2_pharma = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 2 desc limit 5"""
    curr.execute(query_4_2_pharma)

    top_manufacturers_by_category_spend_pharma = curr.fetchall()

    # {4.3} Indirect-IT [ROW-50]
    query_4_3_indirect = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 2 desc limit 5"""
    curr.execute(query_4_3_indirect)

    top_manufacturers_by_category_spend_indirect = curr.fetchall()
    if selected_country == 'Global': 
        
        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Top Manufacturers by Category Spend')
        # Writing in Cells - B51, F51, F51 | Headers
        table_header_value = f'YTD {year} ( {curr_symbol} )'
        table_header_cells = ['B51', 'F51', 'J51']
        
        for cell in table_header_cells:
            copy_wb[sheet_name][cell].value = table_header_value
        
        for run in range(len(top_manufacturers_by_category_spend_non_pharma)):
            ####### Row 52 to 56 ######## 
            cell_number = 52 + run 
            
            ## A, B & C:  top_manufacturers_by_category_spend_non_pharma | cell_value_spend | cell_value_percentage | Row 52 to 56 | A,B,C
            ABC_Cell_values = spend_percentage_value_function(top_manufacturers_by_category_spend_non_pharma, 'A', 'B', 'C', copy_wb, cell_number, sheet_name, run)
            
            ## E, F & G:  top_manufacturers_by_category_spend_pharma | cell_value_spend | cell_value_percentage | Row 52 to 56 | E,F,G 
            EFG_Cell_values = spend_percentage_value_function(top_manufacturers_by_category_spend_pharma, 'E', 'F', 'G', copy_wb, cell_number, sheet_name, run)
            
            ## I, J & K:  top_manufacturers_by_category_spend_indirect | cell_value_spend | cell_value_percentage | Row 52 to 56 | I,J,K
            IJK_Cell_values = spend_percentage_value_function(top_manufacturers_by_category_spend_indirect, 'I', 'J', 'K', copy_wb, cell_number, sheet_name, run)
        
        ## Providing All Others Info | B57, F57, J57 ## 
        all_others, all_others_percentage = getting_all_others_info(top_manufacturers_by_category_spend_non_pharma)
        copy_wb[sheet_name]['B57'].value = all_others
        copy_wb[sheet_name]['C57'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(top_manufacturers_by_category_spend_pharma)
        copy_wb[sheet_name]['F57'].value = all_others
        copy_wb[sheet_name]['G57'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(top_manufacturers_by_category_spend_indirect)
        copy_wb[sheet_name]['J57'].value = all_others
        copy_wb[sheet_name]['K57'].value = all_others_percentage
    # =_=_==_=_== {5} Direct (Non-pharma) Spend [ROW-60]  =_=_==_=_==
    # {5.1} Total P.O. Spend - MTD [ROW-62]
    query_5_1_po_spend = f"""select geography ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query_5_1_po_spend)
    direct_pharma_spend_total_po_spend = curr.fetchall()

    # {5.2} MTD USD Spend [ROW-62] -------- No Need ---------- Taken from Excel Formula
    # {5.3} YTD USD Spend [ROW-62]  ------------- No Need ------------ Taken from Excel Formula


    # {5.4} Total P.O. Spend - YTD Trended [ROW -73]  | {5.4} Total P.O. Spend - YTD Trended [ROW -73] 

    ## 5.4.1 April YTD USD [ROW -74]
    query_5_4_1_ytd = f"""select geography ,sum (dollar_total_cost) as MTD_April_spend ,(MTD_April_spend / SUM(MTD_April_spend) OVER ()) * 100 AS percentage_MTD_April_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query_5_4_1_ytd)
    direct_non_pharma_spend_ytd  = curr.fetchall()

    ## 5.4.2 JAN -23 [ROW -74] | three_months_prior  
    starting_date, ending_date = three_months_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_non_pharma_spend_three_months_prior_value = curr.fetchall()


    ## 5.4.3 FEB -23 [ROW-74] |  two_months_prior  
    starting_date, ending_date = two_months_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_non_pharma_spend_two_months_prior_value = curr.fetchall()


    ## 5.4.4 MAR -23 [ROW-74] |  one_month_prior  
    starting_date, ending_date = one_month_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_non_pharma_spend_one_month_prior_value = curr.fetchall()


    ## 5.4.5 APR - 23 [ROW -74] | current_date 
    starting_date, ending_date = current_date[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_non_pharma_spend_current_date_value = curr.fetchall()
    if selected_country == 'Global': 
            
        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Direct (Non-pharma) Spend')
        
        # Writing in Cells - B63, B74 | Headers
        copy_wb[sheet_name]['B63'].value = f'{month_year_abbr} ( {curr_symbol} )'  # Ex: Apr-23 ( USD )
        copy_wb[sheet_name]['B74'].value = f"{month_year_abbr.split('-')[0]} YTD {curr_symbol}"  # Ex: Apr YTD USD 
        
        # Writing in Cells - D74, E74, F74, G74 | Headers 
        copy_wb[sheet_name]['D74'].value = three_months_prior[-1]  # Jan-23
        copy_wb[sheet_name]['E74'].value = two_months_prior[-1]  # Feb-23
        copy_wb[sheet_name]['F74'].value = one_month_prior[-1]  # Mar-23
        copy_wb[sheet_name]['G74'].value = current_date[-1]  # Apr-23
        
        
        for run in range(len(direct_pharma_spend_total_po_spend)):
            ####### Row 64 to 67 ######## 
            cell_number = 64 + run 
            ## B & C:  direct_non_pharma_spend_total_po_spend | cell_value_spend | cell_value_percentage | Row 64 to 67 | B,C
            BC_Cell_values = spend_percentage_value_function(direct_pharma_spend_total_po_spend, False, 'B', 'C', copy_wb, cell_number, sheet_name, run)
            
            cell_number = 75 + run
            ## B & C:  direct_non_pharma_spend_ytd | cell_value_spend | cell_value_percentage | Row 75 to 78 | A,B,C
            BC_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_ytd, False, 'B', 'C', copy_wb, cell_number, sheet_name, run)
            
            ## D:  direct_non_pharma_spend_three_months_prior_value | cell_value_spend | Row 75 to 78 | D
            D_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_three_months_prior_value, False, 'D', False, copy_wb, cell_number, sheet_name, run)
            
            ## E:  direct_non_pharma_spend_two_months_prior_value | cell_value_spend | Row 75 to 78 | E
            E_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_two_months_prior_value, False, 'E', False, copy_wb, cell_number, sheet_name, run)
            
            ## F:  direct_non_pharma_spend_one_month_prior_value | cell_value_spend | Row 75 to 78 | F
            F_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_one_month_prior_value, False, 'F', False, copy_wb, cell_number, sheet_name, run)
            
            ## G:  direct_non_pharma_spend_current_date_value | cell_value_spend | Row 75 to 78 | G
            G_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_current_date_value, False, 'G', False, copy_wb, cell_number, sheet_name, run)
    # =_=_==_=_== {6} Top Manufacturers | Top Suppliers | Top Product Categories | Direct (Non-Pharma) Spend=_=_==_=_==
    # {6.1} Top Manufacturers  Apr-23  MTD | Row 85
    query_top_manufacturers = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query_top_manufacturers)
    direct_non_pharma_spend_top_manufacturers_value_mtd = curr.fetchall()

    # {6.2} Top Manufacturers  Apr-23  YTD | Row 85
    query_top_manufacturers_ytd = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query_top_manufacturers_ytd)
    direct_non_pharma_spend_top_manufacturers_value_ytd = curr.fetchall()

    # {6.3} Top Suppliers  Apr-23  MTD | Row 85
    query_top_suppliers = f"""select distributor_normalized as suppliers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query_top_suppliers)
    direct_non_pharma_spend_top_suppliers_value_mtd = curr.fetchall()

    # {6.4} Top Suppliers  Apr-23  YTD | Row 85
    query_top_suppliers_ytd = f"""select distributor_normalized as suppliers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query_top_suppliers_ytd)
    direct_non_pharma_spend_top_suppliers_value_ytd = curr.fetchall()

    # {6.5} Top Product Categories  Apr-23  MTD | Row 100
    query_top_product_categories = f"""select unspsc_class_title as categories ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query_top_product_categories)
    direct_non_pharma_spend_top_product_categories_value_mtd = curr.fetchall()

    # {6.6} Top Product Categories  Apr-23  YTD	| Row 100			
    query_top_product_categories_ytd = f"""select unspsc_class_title as categories ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Non-Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query_top_product_categories_ytd)
    direct_non_pharma_spend_top_product_categories_value_ytd = curr.fetchall()
    if selected_country == 'Global': 

        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Direct (Non-pharma) Spend | Top Suppliers | Top Manufacturers | Top Product Categories')
        
        # Writing in Cells - A85, D85, G85, J85 | Headers | Suppliers, Manufacturers
        copy_wb[sheet_name]['A85'].value = f'Top Suppliers {month_year_abbr} MTD'  # Ex: Top Suppliers  Apr-23  MTD
        copy_wb[sheet_name]['D85'].value = f'Top Suppliers {month_year_abbr} YTD'  # Ex: Top Suppliers  Apr-23  YTD
        
        copy_wb[sheet_name]['G85'].value = f'Top Manufacturers {month_year_abbr} MTD'  # Ex: Top Manufacturers  Apr-23  MTD
        copy_wb[sheet_name]['J85'].value = f'Top Manufacturers {month_year_abbr} YTD'  # Ex: Top Manufacturers  Apr-23  YTD
        
        copy_wb[sheet_name]['A100'].value = f'Top Product Categories {month_year_abbr} MTD'  # Ex: Top Product Categories  Apr-23  MTD
        copy_wb[sheet_name]['G100'].value = f'Top Product Categories {month_year_abbr} YTD'  # Ex: Top Product Categories  Apr-23  YTD
        
        
        for run in range(len(direct_non_pharma_spend_top_manufacturers_value_mtd)): # Looping 10 Times 
            ####### Row 87 to 96 ######## 
            cell_number = 87 + run 
            
            ## A,B,C :  direct_non_pharma_spend_top_suppliers_value_mtd | cell_value_spend | cell_value_percentage | Row 87 - 96 | A,B,C
            ABC_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_suppliers_value_mtd, 'A', 'B', 'C', copy_wb, cell_number, sheet_name, run)

            ## D,E,F :  direct_non_pharma_spend_top_suppliers_value_ytd | cell_value_spend | cell_value_percentage | Row 87 - 96 | D,E,F
            DEF_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_suppliers_value_ytd, 'D', 'E', 'F', copy_wb, cell_number, sheet_name, run)
            
            ## G,H,I :  direct_non_pharma_spend_top_manufacturers_value_mtd | cell_value_spend | cell_value_percentage | Row 87 - 96 | G,H,I
            GHI_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_manufacturers_value_mtd, 'G', 'H', 'I', copy_wb, cell_number, sheet_name, run)
            
            ## J,K,L :  direct_non_pharma_spend_top_manufacturers_value_ytd | cell_value_spend | cell_value_percentage | Row 87 - 96 | J,K,L
            JKL_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_manufacturers_value_ytd, 'J', 'K', 'L', copy_wb, cell_number, sheet_name, run)
            
            
            ## Product Categories  | 102 - 107 | MTD | YTD | direct_non_pharma_spend_top_product_categories_value_mtd | direct_non_pharma_spend_top_product_categories_value_ytd
            start_value = 102
            cell_number = start_value + run 
            if cell_number<=start_value + 5 and run<=5:
                
                # A,B,C :  direct_non_pharma_spend_top_product_categories_value_mtd | cell_value_spend | cell_value_percentage | Row 102 - (102+5) | List Values from 0-5
                ABC_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_product_categories_value_mtd, 'A', 'B', 'C', copy_wb, cell_number, sheet_name, run)
                
                # G,H,I :  direct_non_pharma_spend_top_product_categories_value_ytd | cell_value_spend | cell_value_percentage | Row 102 - (102+5) | List Values from 0-5
                GHI_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_product_categories_value_ytd, 'G', 'H', 'I', copy_wb, cell_number, sheet_name, run)
            
            new_run = run + 6
            if cell_number<=start_value + 4 and new_run < 10:
                
                # D,E,F :  direct_non_pharma_spend_top_product_categories_value_mtd | cell_value_spend | cell_value_percentage | Row 102 - (102+4) | List Values from 6-9
                DEF_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_product_categories_value_mtd, 'D', 'E', 'F', copy_wb, cell_number, sheet_name, new_run) # new_run
                
                # J,K,L :  direct_non_pharma_spend_top_product_categories_value_ytd | cell_value_spend | cell_value_percentage | Row 102 - (102+4) | List Values from 6-9
                JKL_Cell_values = spend_percentage_value_function(direct_non_pharma_spend_top_product_categories_value_ytd, 'J', 'K', 'L', copy_wb, cell_number, sheet_name, new_run) # new_run
            
        
        ## Providing All Others Info | B, C, E, F, H, I, K, L  | Row 97 | Top Manufacturers | Top Suppliers - 'All Others' Value
        all_others, all_others_percentage = getting_all_others_info(direct_non_pharma_spend_top_suppliers_value_mtd) # | Top Suppliers
        copy_wb[sheet_name]['B97'].value = all_others
        copy_wb[sheet_name]['C97'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(direct_non_pharma_spend_top_suppliers_value_ytd) # | Top Suppliers
        copy_wb[sheet_name]['E97'].value = all_others
        copy_wb[sheet_name]['F97'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(direct_non_pharma_spend_top_manufacturers_value_mtd) # | Top Manufacturers
        copy_wb[sheet_name]['H97'].value = all_others
        copy_wb[sheet_name]['I97'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(direct_non_pharma_spend_top_manufacturers_value_ytd) # | Top Manufacturers
        copy_wb[sheet_name]['K97'].value = all_others
        copy_wb[sheet_name]['L97'].value = all_others_percentage
        
        ## Providing All Others Info | E, F, K, L | Row 106 | MTD | YTD | Top Product Categories 
        all_others, all_others_percentage = getting_all_others_info(direct_non_pharma_spend_top_product_categories_value_mtd) 
        copy_wb[sheet_name]['E106'].value = all_others
        copy_wb[sheet_name]['F106'].value = all_others_percentage

        all_others, all_others_percentage = getting_all_others_info(direct_non_pharma_spend_top_product_categories_value_ytd) 
        copy_wb[sheet_name]['K106'].value = all_others
        copy_wb[sheet_name]['L106'].value = all_others_percentage
    # =_=_==_=_== {7} Direct Pharma Spend | From [ROW-111]  =_=_==_=_==
    # {7.1} Total P.O. Spend - MTD [ROW-113]
    query = f"""select geography ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_pharma_spend_total_po_spend = curr.fetchall()

    # {7.2} MTD USD Spend [ROW-113] -------- No Need ---------- Taken from Excel Formula
    # {7.3} YTD USD Spend [ROW-113]  ------------- No Need ------------ Taken from Excel Formula


    # {7.4} Total P.O. Spend - YTD Trended [ROW - 124]  | {7 .4} Total P.O. Spend - YTD Trended [ROW - 124] | {7 .4} Total P.O. Spend - YTD Trended [ROW - 124] 

    ## 7.4.1 April YTD USD [ROW - 125] | Apr YTD  USD
    query = f"""select geography ,sum (dollar_total_cost) as MTD_April_spend ,(MTD_April_spend / SUM(MTD_April_spend) OVER ()) * 100 AS percentage_MTD_April_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_pharma_spend_ytd  = curr.fetchall()

    ## 7.4.2 JAN -23 [ROW -125] | three_months_prior  
    starting_date, ending_date = three_months_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_pharma_spend_three_months_prior_value = curr.fetchall()


    ## 7.4.3 FEB -23 [ROW - 125] |  two_months_prior  
    starting_date, ending_date = two_months_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_pharma_spend_two_months_prior_value = curr.fetchall()


    ## 7.4.4 MAR -23 [ROW- 125] |  one_month_prior  
    starting_date, ending_date = one_month_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_pharma_spend_one_month_prior_value = curr.fetchall()


    ## 7.4.5 APR - 23 [ROW - 125] | current_date 
    starting_date, ending_date = current_date[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    direct_pharma_spend_current_date_value = curr.fetchall()
    if selected_country == 'Global': 
            
        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Direct (pharma) Spend')
        
        # Writing in Cells - B114 | Headers
        copy_wb[sheet_name]['B114'].value = f'{month_year_abbr} ( {curr_symbol} )'  # Ex: Apr-23 ( USD )
        
        
        # Writing in Cells - B125, D125, E125, F125, G125 | Headers 
        copy_wb[sheet_name]['B125'].value = f"{month_year_abbr.split('-')[0]} YTD {curr_symbol}"  # Ex: Apr YTD USD 
        copy_wb[sheet_name]['D125'].value = three_months_prior[-1]  # Jan-23
        copy_wb[sheet_name]['E125'].value = two_months_prior[-1]  # Feb-23
        copy_wb[sheet_name]['F125'].value = one_month_prior[-1]  # Mar-23
        copy_wb[sheet_name]['G125'].value = current_date[-1]  # Apr-23
        
        
        for run in range(len(direct_pharma_spend_total_po_spend)):
            ####### Row 115 to 118 ######## 
            cell_number = 115 + run 
            ## B & C:  direct_pharma_spend_total_po_spend | cell_value_spend | cell_value_percentage | Row 115 to 118 | B,C
            BC_Cell_values = spend_percentage_value_function(direct_pharma_spend_total_po_spend, False, 'B', 'C', copy_wb, cell_number, sheet_name, run)
            
            cell_number = 126 + run
            ## B & C:  direct_pharma_spend_ytd | cell_value_spend | cell_value_percentage | Row 126 to 129 | A,B,C
            BC_Cell_values = spend_percentage_value_function(direct_pharma_spend_ytd, False, 'B', 'C', copy_wb, cell_number, sheet_name, run)
            
            ## D:  direct_pharma_spend_three_months_prior_value | cell_value_spend | Row 126 to 129 | D
            D_Cell_values = spend_percentage_value_function(direct_pharma_spend_three_months_prior_value, False, 'D', False, copy_wb, cell_number, sheet_name, run)
            
            ## E:  direct_pharma_spend_two_months_prior_value | cell_value_spend | Row 126 to 129 | E
            E_Cell_values = spend_percentage_value_function(direct_pharma_spend_two_months_prior_value, False, 'E', False, copy_wb, cell_number, sheet_name, run)
            
            ## F:  direct_pharma_spend_one_month_prior_value | cell_value_spend | Row 126 to 129 | F
            F_Cell_values = spend_percentage_value_function(direct_pharma_spend_one_month_prior_value, False, 'F', False, copy_wb, cell_number, sheet_name, run)
            
            ## G:  direct_pharma_spend_current_date_value | cell_value_spend | Row 126 to 129 | G
            G_Cell_values = spend_percentage_value_function(direct_pharma_spend_current_date_value, False, 'G', False, copy_wb, cell_number, sheet_name, run)
    # =_=_==_=_== {8} Top Manufacturers | Top Suppliers | Top Product Categories | Direct (Pharma) Spend | From Row [138]=_=_==_=_==
    # {8.1} Top Manufacturers  Apr-23  MTD | Row 138
    query = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    direct_pharma_spend_top_manufacturers_value_mtd = curr.fetchall()

    # {8.2} Top Manufacturers  Apr-23  YTD | Row 138
    query = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    direct_pharma_spend_top_manufacturers_value_ytd = curr.fetchall()

    # {8.3} Top Suppliers  Apr-23  MTD | Row 138
    query = f"""select distributor_normalized as suppliers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    direct_pharma_spend_top_suppliers_value_mtd = curr.fetchall()

    # {8.4} Top Suppliers  Apr-23  YTD | Row 138
    query = f"""select distributor_normalized as suppliers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    direct_pharma_spend_top_suppliers_value_ytd = curr.fetchall()

    # {8.5} Top Product Categories  Apr-23  MTD | Row 154
    query = f"""select unspsc_class_title as categories ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    direct_pharma_spend_top_product_categories_value_mtd = curr.fetchall()

    # {8.6} Top Product Categories  Apr-23  YTD	| Row 154			
    query = f"""select unspsc_class_title as categories ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Pharma)' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    direct_pharma_spend_top_product_categories_value_ytd = curr.fetchall()
    if selected_country == 'Global': 

        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Direct (Pharma) Spend | Top Suppliers | Top Manufacturers | Top Product Categories')
        
        # Writing in Cells - A138, D138, G138, J138 | Headers | Suppliers, Manufacturers
        copy_wb[sheet_name]['A138'].value = f'Top Suppliers {month_year_abbr} MTD'  # Ex: Top Suppliers  Apr-23  MTD
        copy_wb[sheet_name]['D138'].value = f'Top Suppliers {month_year_abbr} YTD'  # Ex: Top Suppliers  Apr-23  YTD
        
        copy_wb[sheet_name]['G138'].value = f'Top Manufacturers {month_year_abbr} MTD'  # Ex: Top Manufacturers  Apr-23  MTD
        copy_wb[sheet_name]['J138'].value = f'Top Manufacturers {month_year_abbr} YTD'  # Ex: Top Manufacturers  Apr-23  YTD
        
        copy_wb[sheet_name]['A154'].value = f'Top Product Categories {month_year_abbr} MTD'  # Ex: Top Product Categories  Apr-23  MTD
        copy_wb[sheet_name]['G154'].value = f'Top Product Categories {month_year_abbr} YTD'  # Ex: Top Product Categories  Apr-23  YTD
        
        
        for run in range(len(direct_pharma_spend_top_suppliers_value_mtd)): # Looping 10 Times 
            ####### Row 140 to 149 ######## 
            cell_number = 140 + run 
            
            ## A,B,C :  direct_pharma_spend_top_suppliers_value_mtd | cell_value_spend | cell_value_percentage | Row 140 to 149 | A,B,C
            ABC_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_suppliers_value_mtd, 'A', 'B', 'C', copy_wb, cell_number, sheet_name, run)

            ## D,E,F :  direct_pharma_spend_top_suppliers_value_ytd | cell_value_spend | cell_value_percentage | Row 140 to 149 | D,E,F
            DEF_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_suppliers_value_ytd, 'D', 'E', 'F', copy_wb, cell_number, sheet_name, run)
            
            ## G,H,I :  direct_pharma_spend_top_manufacturers_value_mtd | cell_value_spend | cell_value_percentage | Row 140 to 149 | G,H,I
            GHI_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_manufacturers_value_mtd, 'G', 'H', 'I', copy_wb, cell_number, sheet_name, run)
            
            ## J,K,L :  direct_pharma_spend_top_manufacturers_value_ytd | cell_value_spend | cell_value_percentage | Row 140 to 149 | J,K,L
            JKL_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_manufacturers_value_ytd, 'J', 'K', 'L', copy_wb, cell_number, sheet_name, run)
            
            
            ## Product Categories  | 156 - 161 | MTD | YTD | direct_pharma_spend_top_product_categories_value_mtd | direct_pharma_spend_top_product_categories_value_ytd
            start_value = 156
            cell_number = start_value + run 
            if cell_number<=start_value + 5 and run<=5:
                
                # A,B,C :  direct_pharma_spend_top_product_categories_value_mtd | cell_value_spend | cell_value_percentage | Row 156 - (156+5) | List Values from 0-5
                ABC_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_product_categories_value_mtd, 'A', 'B', 'C', copy_wb, cell_number, sheet_name, run)
                
                # G,H,I :  direct_pharma_spend_top_product_categories_value_ytd | cell_value_spend | cell_value_percentage | Row 156 - (156+5) | List Values from 0-5
                GHI_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_product_categories_value_ytd, 'G', 'H', 'I', copy_wb, cell_number, sheet_name, run)
            
            new_run = run + 6
            if cell_number<=start_value + 4 and new_run < 10:
                
                # D,E,F :  direct_pharma_spend_top_product_categories_value_mtd | cell_value_spend | cell_value_percentage | Row 156 - (156+4) | List Values from 6-9
                DEF_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_product_categories_value_mtd, 'D', 'E', 'F', copy_wb, cell_number, sheet_name, new_run) # new_run
                
                # J,K,L :  direct_pharma_spend_top_product_categories_value_ytd | cell_value_spend | cell_value_percentage | Row 156 - (156+4) | List Values from 6-9
                JKL_Cell_values = spend_percentage_value_function(direct_pharma_spend_top_product_categories_value_ytd, 'J', 'K', 'L', copy_wb, cell_number, sheet_name, new_run) # new_run
            
        
        ## Providing All Others Info | B, C, E, F, H, I, K, L  | Row 150 | Top Manufacturers | Top Suppliers - 'All Others' Value
        all_others, all_others_percentage = getting_all_others_info(direct_pharma_spend_top_suppliers_value_mtd) # | Top Suppliers
        copy_wb[sheet_name]['B150'].value = all_others
        copy_wb[sheet_name]['C150'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(direct_pharma_spend_top_suppliers_value_ytd) # | Top Suppliers
        copy_wb[sheet_name]['E150'].value = all_others
        copy_wb[sheet_name]['F150'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(direct_pharma_spend_top_manufacturers_value_mtd) # | Top Manufacturers
        copy_wb[sheet_name]['H150'].value = all_others
        copy_wb[sheet_name]['I150'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(direct_pharma_spend_top_manufacturers_value_ytd) # | Top Manufacturers
        copy_wb[sheet_name]['K150'].value = all_others
        copy_wb[sheet_name]['L150'].value = all_others_percentage
        
        ## Providing All Others Info | E, F, K, L | Row 160 | MTD | YTD | Top Product Categories 
        all_others, all_others_percentage = getting_all_others_info(direct_pharma_spend_top_product_categories_value_mtd)  # Top Product Categories MTD
        copy_wb[sheet_name]['E160'].value = all_others
        copy_wb[sheet_name]['F160'].value = all_others_percentage

        all_others, all_others_percentage = getting_all_others_info(direct_pharma_spend_top_product_categories_value_ytd)  # Top Product Categories YTD
        copy_wb[sheet_name]['K160'].value = all_others
        copy_wb[sheet_name]['L160'].value = all_others_percentage
    # =_=_==_=_== {9} Indirect Spend | From [ROW-165]  =_=_==_=_==
    # {9.1} Total P.O. Spend - MTD [ROW-168]
    query = f"""select geography ,sum (dollar_total_cost) as spend ,(spend / SUM(spend) OVER ()) * 100 AS percentage_spend from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    indirect_spend_total_po_spend = curr.fetchall()

    # {9.2} MTD USD Spend [ROW-113] -------- No Need ---------- Taken from Excel Formula
    # {9.3} YTD USD Spend [ROW-113]  ------------- No Need ------------ Taken from Excel Formula


    # {9.4} Total P.O. Spend - YTD Trended [ROW - 124]  | {9 .4} Total P.O. Spend - YTD Trended [ROW - 124] | {9 .4} Total P.O. Spend - YTD Trended [ROW - 124] 

    ## 9.4.1 April YTD USD [ROW - 168] | Apr YTD  USD
    query = f"""select geography ,sum (dollar_total_cost) as MTD_April_spend ,(MTD_April_spend / SUM(MTD_April_spend) OVER ()) * 100 AS percentage_MTD_April_spend from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    indirect_spend_ytd  = curr.fetchall()

    ## 9.4.2 JAN -23 [ROW -168] | three_months_prior  
    starting_date, ending_date = three_months_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    indirect_spend_three_months_prior_value = curr.fetchall()


    ## 9.4.3 FEB -23 [ROW - 168] |  two_months_prior  
    starting_date, ending_date = two_months_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    indirect_spend_two_months_prior_value = curr.fetchall()


    ## 9.4.4 MAR -23 [ROW- 168] |  one_month_prior  
    starting_date, ending_date = one_month_prior[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    indirect_spend_one_month_prior_value = curr.fetchall()


    ## 9.4.5 APR - 23 [ROW - 168] | current_date 
    starting_date, ending_date = current_date[0:2]
    query = f"""select geography ,sum (dollar_total_cost) as spend from {table_name} where date_of_purchase > '{starting_date}' and date_of_purchase < '{ending_date}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 1"""
    curr.execute(query)
    indirect_spend_current_date_value = curr.fetchall()

    if selected_country == 'Global': 
            
        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Indirect Spend')
        
        # Writing in Cells - B168 | Headers
        copy_wb[sheet_name]['B168'].value = f'{month_year_abbr} ( {curr_symbol} )'  # Ex: Apr-23 ( USD )
        
        
        # Writing in Cells - G168, I168, J168, K168, L168 | Headers 
        copy_wb[sheet_name]['G168'].value = f"{month_year_abbr.split('-')[0]} YTD {curr_symbol}"  # Ex: Apr YTD USD 
        copy_wb[sheet_name]['I168'].value = three_months_prior[-1]  # Jan-23
        copy_wb[sheet_name]['J168'].value = two_months_prior[-1]  # Feb-23
        copy_wb[sheet_name]['K168'].value = one_month_prior[-1]  # Mar-23
        copy_wb[sheet_name]['L168'].value = current_date[-1]  # Apr-23
        

        for run in range(len(indirect_spend_total_po_spend)):
            ####### Row 169 to 172 ######## 
            cell_number = 169 + run 
            
            ## B & C:  indirect_spend_total_po_spend | cell_value_spend | cell_value_percentage | Row 169 to 172
            BC_Cell_values = spend_percentage_value_function(indirect_spend_total_po_spend, False, 'B', 'C', copy_wb, cell_number, sheet_name, run)
            
            ## G & H:  indirect_spend_ytd | cell_value_spend | cell_value_percentage | Row 169 to 172 
            GH_Cell_values = spend_percentage_value_function(indirect_spend_ytd, False, 'G', 'H', copy_wb, cell_number, sheet_name, run)
        
            ## I:  indirect_spend_three_months_prior_value | cell_value_spend | Row 169 to 172
            I_Cell_values = spend_percentage_value_function(indirect_spend_three_months_prior_value, False, 'I', False, copy_wb, cell_number, sheet_name, run)
            
            ## J:  indirect_spend_two_months_prior_value | cell_value_spend | Row 169 to 172
            J_Cell_values = spend_percentage_value_function(indirect_spend_two_months_prior_value, False, 'J', False, copy_wb, cell_number, sheet_name, run)
            
            ## K:  indirect_spend_one_month_prior_value | cell_value_spend | Row 169 to 172
            K_Cell_values = spend_percentage_value_function(indirect_spend_one_month_prior_value, False, 'K', False, copy_wb, cell_number, sheet_name, run)
            
            ## L:  indirect_spend_current_date_value | cell_value_spend | Row 169 to 172
            L_Cell_values = spend_percentage_value_function(indirect_spend_current_date_value, False, 'L', False, copy_wb, cell_number, sheet_name, run)
    # =_=_==_=_== {10} Top Manufacturers | Top Suppliers | Top Product Categories | Indirect Spend | From Row [178]=_=_==_=_==
    # {10.1} Top Manufacturers  Apr-23  MTD | Row 178
    query = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    indirect_spend_top_manufacturers_value_mtd = curr.fetchall()

    # {10.2} Top Manufacturers  Apr-23  YTD | Row 178
    query = f"""select mnf_dashboard_half as manufacturers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    indirect_spend_top_manufacturers_value_ytd = curr.fetchall()

    # {10.3} Top Suppliers  Apr-23  MTD | Row 178
    query = f"""select distributor_normalized as suppliers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    indirect_spend_top_suppliers_value_mtd = curr.fetchall()

    # {10.4} Top Suppliers  Apr-23  YTD | Row 178
    query = f"""select distributor_normalized as suppliers ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    indirect_spend_top_suppliers_value_ytd = curr.fetchall()

    # {10.5} Top Product Categories  Apr-23  MTD | Row 193
    query = f"""select unspsc_class_title as categories ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_mtd}' and date_of_purchase < '{end_date_mtd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    indirect_spend_top_product_categories_value_mtd = curr.fetchall()

    # {10.6} Top Product Categories  Apr-23  YTD	| Row 193			
    query = f"""select unspsc_class_title as categories ,sum (dollar_total_cost) as USD_global ,(USD_global/ SUM(USD_global) OVER ()) * 100 AS percentage_USD_global from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Indirect' GROUP BY 1 ORDER BY 2 desc limit 10"""
    curr.execute(query)
    indirect_spend_top_product_categories_value_ytd = curr.fetchall()
    if selected_country == 'Global': 

        print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Indirect Spend | Top Suppliers | Top Manufacturers | Top Product Categories')
        
        # Writing in Cells - A178, D178, G178, J178, A193, G193 | Headers | Suppliers, Manufacturers
        copy_wb[sheet_name]['A178'].value = f'Top Suppliers {month_year_abbr} MTD'  # Ex: Top Suppliers  Apr-23  MTD
        copy_wb[sheet_name]['D178'].value = f'Top Suppliers {month_year_abbr} YTD'  # Ex: Top Suppliers  Apr-23  YTD
        
        copy_wb[sheet_name]['G178'].value = f'Top Manufacturers {month_year_abbr} MTD'  # Ex: Top Manufacturers  Apr-23  MTD
        copy_wb[sheet_name]['J178'].value = f'Top Manufacturers {month_year_abbr} YTD'  # Ex: Top Manufacturers  Apr-23  YTD
        
        copy_wb[sheet_name]['A193'].value = f'Top Product Categories {month_year_abbr} MTD'  # Ex: Top Product Categories  Apr-23  MTD
        copy_wb[sheet_name]['G193'].value = f'Top Product Categories {month_year_abbr} YTD'  # Ex: Top Product Categories  Apr-23  YTD
        

        for run in range(len(indirect_spend_top_suppliers_value_mtd)): # Looping 10 Times 
            ####### Row 180 to 189 ######## 
            cell_number = 180 + run 
            
            ## A,B,C :  indirect_spend_top_suppliers_value_mtd | cell_value_spend | cell_value_percentage | Row 180 to 189 | A,B,C
            ABC_Cell_values = spend_percentage_value_function(indirect_spend_top_suppliers_value_mtd, 'A', 'B', 'C', copy_wb, cell_number, sheet_name, run)

            ## D,E,F :  indirect_spend_top_suppliers_value_ytd | cell_value_spend | cell_value_percentage | Row 180 to 189 | D,E,F
            DEF_Cell_values = spend_percentage_value_function(indirect_spend_top_suppliers_value_ytd, 'D', 'E', 'F', copy_wb, cell_number, sheet_name, run)
            
            ## G,H,I :  indirect_spend_top_manufacturers_value_mtd | cell_value_spend | cell_value_percentage | Row 180 to 189 | G,H,I
            GHI_Cell_values = spend_percentage_value_function(indirect_spend_top_manufacturers_value_mtd, 'G', 'H', 'I', copy_wb, cell_number, sheet_name, run)
            
            ## J,K,L :  indirect_spend_top_manufacturers_value_ytd | cell_value_spend | cell_value_percentage | Row 180 to 189 | J,K,L
            JKL_Cell_values = spend_percentage_value_function(indirect_spend_top_manufacturers_value_ytd, 'J', 'K', 'L', copy_wb, cell_number, sheet_name, run)
            
            
            ## Product Categories  | 195 - 200 | MTD | YTD | indirect_spend_top_product_categories_value_mtd | indirect_spend_top_product_categories_value_ytd
            start_value = 195
            cell_number = start_value + run 
            if cell_number<=start_value + 5 and run<=5:
                
                # A,B,C :  indirect_spend_top_product_categories_value_mtd | cell_value_spend | cell_value_percentage | Row 195 - (195+5) | List Values from 0-5
                ABC_Cell_values = spend_percentage_value_function(indirect_spend_top_product_categories_value_mtd, 'A', 'B', 'C', copy_wb, cell_number, sheet_name, run)
                
                # G,H,I :  indirect_spend_top_product_categories_value_ytd | cell_value_spend | cell_value_percentage | Row 195 - (195+5) | List Values from 0-5
                GHI_Cell_values = spend_percentage_value_function(indirect_spend_top_product_categories_value_ytd, 'G', 'H', 'I', copy_wb, cell_number, sheet_name, run)
            
            new_run = run + 6
            if cell_number<=start_value + 4 and new_run < 10:
                
                # D,E,F :  indirect_spend_top_product_categories_value_mtd | cell_value_spend | cell_value_percentage | Row 195 - (195+4) | List Values from 6-9
                DEF_Cell_values = spend_percentage_value_function(indirect_spend_top_product_categories_value_mtd, 'D', 'E', 'F', copy_wb, cell_number, sheet_name, new_run) # new_run
                
                # J,K,L :  indirect_spend_top_product_categories_value_ytd | cell_value_spend | cell_value_percentage | Row 195 - (195+4) | List Values from 6-9
                JKL_Cell_values = spend_percentage_value_function(indirect_spend_top_product_categories_value_ytd, 'J', 'K', 'L', copy_wb, cell_number, sheet_name, new_run) # new_run
            

        
        ## Providing All Others Info | B, C, E, F, H, I, K, L  | Row 190 | Top Manufacturers | Top Suppliers - 'All Others' Value
        all_others, all_others_percentage = getting_all_others_info(indirect_spend_top_suppliers_value_mtd) # | Top Suppliers
        copy_wb[sheet_name]['B190'].value = all_others
        copy_wb[sheet_name]['C190'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(indirect_spend_top_suppliers_value_ytd) # | Top Suppliers
        copy_wb[sheet_name]['E190'].value = all_others
        copy_wb[sheet_name]['F190'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(indirect_spend_top_manufacturers_value_mtd) # | Top Manufacturers
        copy_wb[sheet_name]['H190'].value = all_others
        copy_wb[sheet_name]['I190'].value = all_others_percentage
        
        all_others, all_others_percentage = getting_all_others_info(indirect_spend_top_manufacturers_value_ytd) # | Top Manufacturers
        copy_wb[sheet_name]['K190'].value = all_others
        copy_wb[sheet_name]['L190'].value = all_others_percentage
        
        ## Providing All Others Info | E, F, K, L | Row 199 | MTD | YTD | Top Product Categories 
        all_others, all_others_percentage = getting_all_others_info(indirect_spend_top_product_categories_value_mtd)  # Top Product Categories MTD
        copy_wb[sheet_name]['E199'].value = all_others
        copy_wb[sheet_name]['F199'].value = all_others_percentage

        all_others, all_others_percentage = getting_all_others_info(indirect_spend_top_product_categories_value_ytd)  # Top Product Categories YTD
        copy_wb[sheet_name]['K199'].value = all_others
        copy_wb[sheet_name]['L199'].value = all_others_percentage
    # Sheet2 | Updating
    query = f"""WITH cte1 AS ( SELECT geography AS market, COUNT(DISTINCT mnf_dashboard_half) AS unique_Manufacturers, COUNT(DISTINCT distributor_normalized) AS unique_suppliers, COUNT(DISTINCT sc_uhg_id) AS unique_SKUs, SUM(dollar_total_cost) AS USD_Spend_MTD FROM {table_name} WHERE date_of_purchase > '{start_date_mtd}' AND date_of_purchase < '{end_date_mtd}' GROUP BY geography ORDER BY geography ), cte2 AS ( SELECT geography AS market, SUM(dollar_total_cost) AS USD_Spend_YTD FROM {table_name} WHERE date_of_purchase > '{start_date_ytd}' AND date_of_purchase < '{end_date_mtd}' GROUP BY geography ORDER BY geography ) SELECT cte1.market, cte1.unique_Manufacturers, cte1.unique_suppliers, cte1.unique_SKUs, cte1.USD_Spend_MTD, cte2.USD_Spend_YTD FROM cte1 JOIN cte2 ON cte1.market = cte2.market GROUP BY cte1.market, cte1.unique_Manufacturers, cte1.unique_suppliers, cte1.unique_SKUs, cte1.USD_Spend_MTD, cte2.USD_Spend_YTD ORDER BY cte1.market"""
    curr.execute(query)

    value = curr.fetchall()

    col_names = ['B', 'C', 'D', "E", 'F']
    for cell_number in range(2, 6):
        sum = 1
        for i, col_name in enumerate(col_names):
            cell_value = list(value[cell_number-2].values())[sum]
            copy_wb['Sheet2'][f'{col_name}{cell_number}'].value = cell_value
            sum = sum+1

    print(f'Completed Updating - Sheet2')
    # =_=_==_=_== {11} Total P.O. Spend - YoY Trend | H,J,L | Row: 21 to 25 =_=_==_=_==
    def get_date_ranges(year):
        year = int(year)
        date_ranges = []

        for i in range(year-3, year):
            date_range_start = datetime(i-1, 12, 31)  # End of the current year
            date_range_end = datetime(i+1, 1, 1)  # Start of the next year
            date_ranges.append((date_range_start.strftime('%Y-%m-%d'), date_range_end.strftime('%Y-%m-%d')))

        return date_ranges

    yoy_date_range = get_date_ranges(year) # Getting Date Range for Selected Year
    column_names_yoy_table = [yoy_date_range[1][0].split('-')[0], yoy_date_range[-1][0].split('-')[0], yoy_date_range[1][1].split('-')[0]]
    yoy_countries = ['Brazil', 'Chile', 'Colombia', 'Peru', 'Portugal']
    print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Total P.O. Spend - YoY Trend')

    yoy_records = []
    for yoy_country in yoy_countries:
        country_value = []
        for dt in yoy_date_range:
            start_yoy_date, end_yoy_date = dt
            
            if start_yoy_date == '2019-12-31' and end_yoy_date == '2021-01-01': # Query for 2020 Data
                
                query = f"""select sum (dollar_total_cost) as MTD_usd from {table_name} where geography = '{yoy_country}' and date_of_purchase > '{start_yoy_date}' and date_of_purchase < '{end_yoy_date}' and lower(spend_type_1) = 'overall'"""
            else:
                query = f"""select sum (dollar_total_cost) as MTD_usd_local from {table_name} where date_of_purchase > '{start_yoy_date}' and date_of_purchase < '{end_yoy_date}' and geography = '{yoy_country}'"""
            
            curr.execute(query)
            value = curr.fetchall()
            value = list(value[0].values())[0]
            country_value.append(value)

            # print(f'{query} | {value}')
        
        yoy_records.append(country_value)
        
    qf = pd.DataFrame(yoy_records, columns=[column_names_yoy_table])
    qf = qf.fillna(0)

    ## Row 21 to 25 | Including Portugal Information | columns = ['H', 'J', 'L']
    start_row = 21
    columns = ['H', 'J', 'L']
    for row in range(start_row, start_row + len(yoy_countries)):
        for j, col in enumerate(columns): 
            cell_name = f'{col}{row}'
            cell_value = qf.iloc[row - start_row][j]
            copy_wb[sheet_name][cell_name].value = cell_value
    # =_=_==_=_=={12} Non-Pharma Spend YoY Trend | B,D,F | Row 41 to 45 | H,J,L | Row 64 to 68=_=_==_=_==
    print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Non-Pharma Spend YoY Trend')

    yoy_records = []
    for yoy_country in yoy_countries:
        country_value = []
        for dt in yoy_date_range:
            start_yoy_date, end_yoy_date = dt
            
            query = f"""select sum (dollar_total_cost) as MTD_usd_local from {table_name} where date_of_purchase > '{start_yoy_date}' and date_of_purchase < '{end_yoy_date}' and geography = '{yoy_country}' and spend_type_1 = 'Direct (Non-Pharma)'"""
            curr.execute(query)
            value = curr.fetchall()
            value = list(value[0].values())[0]
            country_value.append(value)
            # print(f'{query} | {value}')
        
        yoy_records.append(country_value)
        
    qf = pd.DataFrame(yoy_records, columns=[column_names_yoy_table])
    qf = qf.fillna(0)

    ## Row 41 to 45 | Including Portugal Information | columns = ['B', 'D', 'F']
    start_row = 41
    columns = ['B', 'D', 'F']
    for row in range(start_row, start_row + len(yoy_countries)):
        for j, col in enumerate(columns): 
            cell_name = f'{col}{row}'
            cell_value = qf.iloc[row - start_row][j]
            copy_wb[sheet_name][cell_name].value = cell_value
            
    ## Row 64 to 68 | Including Portugal Information | columns = ['H', 'J', 'L']
    start_row = 64
    columns = ['H', 'J', 'L']
    for row in range(start_row, start_row + len(yoy_countries)):
        for j, col in enumerate(columns): 
            cell_name = f'{col}{row}'
            cell_value = qf.iloc[row - start_row][j]
            copy_wb[sheet_name][cell_name].value = cell_value
    # =_=_==_=_=={13} Pharma Spend YoY Trend | H,J,L | Row 41 to 45 | Row 115 to 119=_=_==_=_==
    print(f'{selected_country} - {sheet_name} - {curr_symbol} | Table: Pharma Spend YoY Trend')
    yoy_records = []
    for yoy_country in yoy_countries:
        country_value = []
        for dt in yoy_date_range:
            start_yoy_date, end_yoy_date = dt
            query = f"""select sum (dollar_total_cost) as MTD_usd_local from {table_name} where date_of_purchase > '{start_yoy_date}' and date_of_purchase < '{end_yoy_date}' and geography = '{yoy_country}' and spend_type_1 = 'Direct (Pharma)'"""
            curr.execute(query)
            value = curr.fetchall()
            value = list(value[0].values())[0]
            country_value.append(value)
            # print(f'{query} | {value}')
        
        yoy_records.append(country_value)
        
    qf = pd.DataFrame(yoy_records, columns=[column_names_yoy_table])
    qf = qf.fillna(0)

    ## Row 41 to 45 | Including Portugal Information | columns = ['H', 'J', 'L']
    start_row = 41
    columns = ['H', 'J', 'L']
    for row in range(start_row, start_row + len(yoy_countries)):
        for j, col in enumerate(columns): 
            cell_name = f'{col}{row}'
            cell_value = qf.iloc[row - start_row][j]
            copy_wb[sheet_name][cell_name].value = cell_value
            
    ## Row 115 to 119 | Including Portugal Information | columns = ['H', 'J', 'L']
    start_row = 115
    columns = ['H', 'J', 'L']
    for row in range(start_row, start_row + len(yoy_countries)):
        for j, col in enumerate(columns): 
            cell_name = f'{col}{row}'
            cell_value = qf.iloc[row - start_row][j]
            copy_wb[sheet_name][cell_name].value = cell_value
    # Single Values |  A,E,I | Row 59 | *Total unique mfg  3131	| *Total unique mfg  693	| *Total unique mfg  518
    ### Ex: *Total unique mfg  3131 | A | Non-Pharma

    query = f"""select count(distinct(mnf_dashboard_half)) as total_unique_mfg_for_NonPharma from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Non-Pharma)'"""
    curr.execute(query)
    value = curr.fetchall()
    # print(value)
    copy_wb[sheet_name]['A59'].value = f"*Total unique mfg {list(value[0].values())[0]}"

    ### Ex: *Total unique mfg  693 | E | Pharma

    query = f"""select count(distinct(mnf_dashboard_half)) as total_unique_mfg_for_Pharma from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Direct (Pharma)'"""
    curr.execute(query)
    value = curr.fetchall()
    # print(value)
    copy_wb[sheet_name]['E59'].value = f"*Total unique mfg {list(value[0].values())[0]}"

    ### Ex: *Total unique mfg  518 | I | Indirect

    query = f"""select count(distinct(mnf_dashboard_half)) as total_unique_mfg_for_NonPharma from {table_name} where date_of_purchase > '{start_date_ytd}' and date_of_purchase < '{end_date_ytd}' and spend_type_1 = 'Indirect'"""
    curr.execute(query)
    value = curr.fetchall()
    # print(value)
    copy_wb[sheet_name]['I59'].value = f"*Total unique mfg {list(value[0].values())[0]}"
    # Saving Everything | Closing
    copy_wb.save(copy_file_path)
    print(f'Completed: {file_name}')