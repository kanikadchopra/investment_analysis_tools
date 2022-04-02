#!/usr/bin/env python
# coding: utf-8

# In[97]:


import pdfplumber
import pandas as pd


# In[98]:


def extract_pages(file):
    lines = []

    with pdfplumber.open(file) as pdf:
        pages = pdf.pages
        for page in pdf.pages:
            text = page.extract_text()
            if 'Asset Review' in text: 
                pg = []
                for line in text.split('\n'):
                    pg.append(line)
                lines.append(pg)
    return lines


# In[99]:


def extract_info(file):
    
    pages = extract_pages(file)
    
    cdn_dfs = []
    us_dfs = [] 
    for page in pages:
        year = page[1][-4:] 
        if 'rrsp' in file.lower():
            currency = page[1][-11:-7]
            currency = currency.replace('(', '')
        else:
            currency = page[1][:3].upper()
        
        start = False
        asset_data = []
        for line in page:
            if 'Common Shares' in line or 'Foreign Securities' in line:
                start = True
            elif 'TotalValueofCommonShares' in line or 'TotalValueofForeignSecurities' in line:
                start = False

            if start == True: 
                asset_data.append(line)
                
        # Clean up the data
        asset_data = asset_data[1:]
        asset_data = [asset.split(' ') for asset in asset_data]
        
        # Create dataframe
        df = pd.DataFrame(asset_data)
        df.columns = ['Name', 'Symbol', 'Quantity', 'Price', 'Book Cost', 'Market Value']
        df.dropna(inplace=True)
        df.drop(columns = ['Name', 'Book Cost', 'Market Value'], inplace=True)
        df.set_index('Symbol', inplace=True)   
        
        # Change data types
        df.Quantity = df.Quantity.astype('float')
        df.Price = df.Price.astype('float')
        
        df.columns = [name + '_' + year for name in df.columns]
        
        if currency == 'CDN':
            cdn_dfs.append(df)
        
        else:
            us_dfs.append(df)
        
    suffix = '_' + file.split('_')[0]
    d = {'data_CDN' + suffix :  pd.concat(cdn_dfs, axis=0),
         'data_US' + suffix : pd.concat(us_dfs, axis=0)}
     
    return d


# ## For first time use to create the initial excel file

# In[100]:


investments_2020 = extract_info('non_registered_2020.pdf')
rrsp_2020 = extract_info('rrsp_2020.pdf')
investments_2020.update(rrsp_2020)


# In[102]:


from pandas import ExcelWriter
from openpyxl import load_workbook


def save_xls(dict_df, path):
    book = load_workbook(path)
    
    writer = ExcelWriter(path, engine='openpyxl')
    writer.book = book
    
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    
    
    for key in dict_df:
        dict_df[key].to_excel(writer, key)

    writer.save()


# In[103]:


save_xls(investments_2020, 'statement_analysis.xlsx')


# ## After first use to add the new data or accounts onto this

# In[105]:


def add_data(analysis_csv, new_pdf):
    # Read in new data
    data = extract_info(new_pdf)
    suffix = '_' + new_pdf.split('_')[0]
    
    # Canadian
    cdn_data = data['data_CDN' + suffix]
    old_cdn = pd.read_excel(analysis_csv, sheet_name = 'data_CDN' + suffix)
    old_cdn.set_index('Symbol', inplace=True)
    new_cdn = pd.concat([old_cdn, cdn_data], axis=1)
    
    cdn_analysis = new_cdn[[col for col in new_cdn.columns if 'Price' in col]].pct_change(axis=1)
    cdn_analysis *= 100 
    cdn_analysis.dropna(axis=1, inplace=True, how='all')

    # US
    us_data = data['data_US' + suffix]
    old_us = pd.read_excel(analysis_csv, sheet_name = 'data_US' + suffix)
    old_us.set_index('Symbol', inplace=True)
    new_us = pd.concat([old_us, us_data], axis=1)

    us_analysis = new_us[[col for col in new_us.columns if 'Price' in col]].pct_change(axis=1)
    us_analysis *= 100 
    years = [s[-4:] for s in us_analysis.columns]
    
    us_analysis.dropna(axis=1, inplace=True, how='all')
    
    i = 0
    pct_cols = []
    while i < len(years) - 1:
        col = years[i] + '_to_' + years[i+1]
        pct_cols.append(col)
        i += 1
        
    us_analysis.columns = pct_cols
    cdn_analysis.columns = pct_cols
    
    # Create the new dictionary 
    d = {'data_CDN' + suffix: new_cdn, 
         'analysis_CDN' + suffix : cdn_analysis,
         'data_US' + suffix : new_us,
         'analysis_US' + suffix: us_analysis}
    
    save_xls(d, 'statement_analysis.xlsx')
    
    
    
    print('Yay!')


# In[106]:


add_data('statement_analysis.xlsx', 'rrsp_2021.pdf')


# In[107]:


add_data('statement_analysis.xlsx', 'non_registered_2021.pdf')

