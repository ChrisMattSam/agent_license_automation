# -*- coding: utf-8 -*-
"""
Created on Sun Apr 12 12:43:59 2020

@author: CSR1
"""
import pandas as pd
from numpy import nan as nan
from mailmerge import MailMerge

def pad_the_df(df, max_rows = 12):
    '''
    If dataframe has less than max_row rows, add empty rows to the bottom 
    of it to make the new dataframe have that number of rows
    '''
    i = 0
    if df.shape[0] >= max_rows:
        return df
    else:
        while i <= 11:
            df = df.append(pd.Series(dtype = 'str'), ignore_index = True)
            i +=1
        return df.replace(nan,'place_holder').head(12)

def fix(number):
    if type(number) is not str:
        return str(round(number)).replace('.0','')
    else:
        return number
    
def split_the_data(df, n = 12):
    '''
    Takes a dataframe and returns a list of dataframes with n or less number of
    rows
    
    '''
    df.replace(nan,'not_provided', inplace = True)
    number_of_pages = int(df.shape[0]/n)
    holder = []
    i = 0
    while i <= number_of_pages :
        if df.shape[0] <n:
            holder.append(df)
        else:
            holder.append(df.head(n))
        df = df.iloc[n:]
        i +=1
    holder = [h for h in holder if h.shape[0] > 0]
    return holder

def template_fields(template, return_them = False):
        document = MailMerge(template)
        print(sorted(document.get_merge_fields()))
        if return_them:
            return sorted(document.get_merge_fields())

def pre_process(df):
    '''
    Split the first and last name into separate columns,
    
    '''
    df[['last_name','first_name']] = df.name.str.split(',', expand = True)
    df.drop(columns = 'name', inplace = True)
    df['first_name'] = df['first_name'].apply(lambda x: x.lstrip()) #remove leading space
    for col in ['license_number_1', 'license_number_2', 'license_number_3']:
        df[col]= df[col].astype('str').str.replace('p0','0').replace(' ','')
    return df


