# -*- coding: utf-8 -*-
"""
Created on Sat Apr 11 15:47:44 2020

Found this link to be very helpful: https://pbpython.com/python-word-template.html


@author: CSR1
"""

from __future__ import print_function
from mailmerge import MailMerge
import pandas as pd
from numpy import nan as nan
import os


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
    return holder

def pad_the_df(df, max_rows = 12):
    '''
    If dataframe has less than max_row rows, add empty rows to the bottom 
    of it to make the new dataframe have that number of rows
    '''
    i = 0
    while i <= (12-df.shape[0]):
        df = df.append(pd.Series(dtype = 'str'), ignore_index = True)
        i +=1
    df.replace(nan,'place_holder', inplace = True)
    return df
    
def create_page(df, office_phone_number, office_street_address, office_city_state_zip, 
                template ="12 cards template test2.docx", file_name = None,
                branch_license_number = None, company_license_number = None):
    '''
    Create the final printed page.  To-do: add the template, sample inputs and sample outputs
    '''
    df['street_address'] = office_street_address
    df['phone_number'] = office_phone_number
    df['city_state_zip'] = office_city_state_zip
    if company_license_number is None:
        df['company_license_number'] = '0226004377'
        print('Using the default company license number of 0226004377' )
    if branch_license_number is None:
        df['branch_license_number'] = '0226004375'
        print('Using the default branch license number of 0226004375\n')
    
    print('Creating the page of contact cards...')
    template = template
    document = MailMerge(template)
    
    
    'These are the first 12 rows of the dataset, representing 12 agents'
    locs = list(range(0,12))
    if df.shape[0] < 12:
        df = pad_the_df(df)
    
    
    l = [df.iloc[i] for i in locs]
    row_1 = l[0]
    row_2 = l[1]
    row_3 = l[2]
    row_4 = l[3]
    row_5 = l[4]
    row_6 = l[5]
    row_7 = l[6]
    row_8 = l[7]
    row_9 = l[8]
    row_10 = l[9]
    row_11 = l[10]
    row_12 = l[11]
    
    
    document = MailMerge(template)
    document.merge(
            
        state_1 = row_1['state'],
        state_2 = row_2['state'],
        state_3 = row_3['state'],
        state_4 = row_4['state'],
        state_5 = row_5['state'],
        state_6 = row_6['state'],
        state_7 = row_7['state'],
        state_8 = row_8['state'],
        state_9 = row_9['state'],
        state_10 = row_10['state'],
        state_11 = row_11['state'],
        state_12 = row_12['state'],
        
        #fill the branch license number for all 12 squares'
        branch_lic_num_1 = fix(row_1['branch_license_number']),
        branch_lic_num_2 = fix(row_2['branch_license_number']),
        branch_lic_num_3 = fix(row_3['branch_license_number']),
        branch_lic_num_4 = fix(row_4['branch_license_number']),
        branch_lic_num_5 = fix(row_5['branch_license_number']),
        branch_lic_num_6 = fix(row_6['branch_license_number']),
        branch_lic_num_7 = fix(row_7['branch_license_number']),
        branch_lic_num_8 = fix(row_8['branch_license_number']),
        branch_lic_num_9 = fix(row_9['branch_license_number']),
        branch_lic_num_10 = fix(row_10['branch_license_number']),
        branch_lic_num_11 = fix(row_11['branch_license_number']),
        branch_lic_num_12 = fix(row_12['branch_license_number']),
        
        comp_lic_num_1 = fix(row_1['company_license_number']),
        comp_lic_num_2 = fix(row_2['company_license_number']),
        comp_lic_num_3 = fix(row_3['company_license_number']),
        comp_lic_num_4 = fix(row_4['company_license_number']),
        comp_lic_num_5 = fix(row_5['company_license_number']),
        comp_lic_num_6 = fix(row_6['company_license_number']),
        comp_lic_num_7 = fix(row_7['company_license_number']),
        comp_lic_num_8 = fix(row_8['company_license_number']),
        comp_lic_num_9 = fix(row_9['company_license_number']),
        comp_lic_num_10 = fix(row_10['company_license_number']),
        comp_lic_num_11 = fix(row_11['company_license_number']),
        comp_lic_num_12 = fix(row_12['company_license_number']),
        
        exp_date_1 = row_1['expiration_date'],
        exp_date_2 = row_2['expiration_date'],
        exp_date_3 = row_3['expiration_date'],
        exp_date_4 = row_4['expiration_date'],
        exp_date_5 = row_5['expiration_date'],
        exp_date_6 = row_6['expiration_date'],
        exp_date_7 = row_7['expiration_date'],
        exp_date_8 = row_8['expiration_date'],
        exp_date_9 = row_9['expiration_date'],
        exp_date_10 = row_10['expiration_date'],
        exp_date_11 = row_11['expiration_date'],
        exp_date_12 = row_12['expiration_date'],
        
        first_1 = row_1['first_name'],
        first_2 = row_2['first_name'],
        first_3 = row_3['first_name'],
        first_4 = row_4['first_name'],
        first_5 = row_5['first_name'],
        first_6 = row_6['first_name'],
        first_7 = row_7['first_name'],
        first_8 = row_8['first_name'],
        first_9 = row_9['first_name'],
        first_10 = row_10['first_name'],
        first_11 = row_11['first_name'],
        first_12 = row_12['first_name'],
        
        last_1 = row_1['last_name'],
        last_2 = row_2['last_name'],
        last_3 = row_3['last_name'],
        last_4 = row_4['last_name'],
        last_5 = row_5['last_name'],
        last_6 = row_6['last_name'],
        last_7 = row_7['last_name'],
        last_8 = row_8['last_name'],
        last_9 = row_9['last_name'],
        last_10 = row_10['last_name'],
        last_11 = row_11['last_name'],
        last_12 = row_12['last_name'],
        
        lic_num_1 = fix(row_1['license_number']),
        lic_num_2 = fix(row_2['license_number']),
        lic_num_3 = fix(row_3['license_number']),
        lic_num_4 = fix(row_4['license_number']),
        lic_num_5 = fix(row_5['license_number']),
        lic_num_6 = fix(row_6['license_number']),
        lic_num_7 = fix(row_7['license_number']),
        lic_num_8 = fix(row_8['license_number']),
        lic_num_9 = fix(row_9['license_number']),
        lic_num_10 = fix(row_10['license_number']),
        lic_num_11 = fix(row_11['license_number']),
        lic_num_12 = fix(row_12['license_number']),
        
        street_address_1 = fix(row_1['street_address']),
        street_address_2 = fix(row_2['street_address']),
        street_address_3 = fix(row_3['street_address']),
        street_address_4 = fix(row_4['street_address']),
        street_address_5 = fix(row_5['street_address']),
        street_address_6 = fix(row_6['street_address']),
        street_address_7 = fix(row_7['street_address']),
        street_address_8 = fix(row_8['street_address']),
        street_address_9 = fix(row_9['street_address']),
        street_address_10 = fix(row_10['street_address']),
        street_address_11 = fix(row_11['street_address']),
        street_address_12 = fix(row_12['street_address']),
        
        city_state_zip_1 = fix(row_1['city_state_zip']),
        city_state_zip_2 = fix(row_2['city_state_zip']),
        city_state_zip_3 = fix(row_3['city_state_zip']),
        city_state_zip_4 = fix(row_4['city_state_zip']),
        city_state_zip_5 = fix(row_5['city_state_zip']),
        city_state_zip_6 = fix(row_6['city_state_zip']),
        city_state_zip_7 = fix(row_7['city_state_zip']),
        city_state_zip_8 = fix(row_8['city_state_zip']),
        city_state_zip_9 = fix(row_9['city_state_zip']),
        city_state_zip_10 = fix(row_10['city_state_zip']),
        city_state_zip_11 = fix(row_11['city_state_zip']),
        city_state_zip_12 = fix(row_12['city_state_zip']),
        
        phone_1 = fix(row_1['phone_number']),
        phone_2 = fix(row_2['phone_number']),
        phone_3 = fix(row_3['phone_number']),
        phone_4 = fix(row_4['phone_number']),
        phone_5 = fix(row_5['phone_number']),
        phone_6 = fix(row_6['phone_number']),
        phone_7 = fix(row_7['phone_number']),
        phone_8 = fix(row_8['phone_number']),
        phone_9 = fix(row_9['phone_number']),
        phone_10 = fix(row_10['phone_number']),
        phone_11 = fix(row_11['phone_number']),
        phone_12 = fix(row_12['phone_number']),
        )
    
    if not os.path.exists('output/'):
        os.mkdir('output/')
    
    df = df.loc[df['last_name'] != 'place_holder']
    if file_name is None:
        n = df.shape[0]            
        file_name = df['last_name'].iloc[0] + '_to_' + df['last_name'].iloc[n-1]
    document.write('output/' + file_name + '.docx')
    print('Done')

def template_fields(template, return_them = False):
        document = MailMerge(template)
        print(sorted(document.get_merge_fields()))
        if return_them:
            return sorted(document.get_merge_fields())
        
if __name__ == '__main__':
    
    'Generate files for the McLean office'
    df = pd.read_csv("test_licensure_workbook_mclean.csv")
    template_fields('template-one license.docx')
    
    mclean = ['6631 Old Dominion Dr','McLean, VA 22101','703.556.4222']
    pages = split_the_data(df)
    [create_page(df, office_phone_number = mclean[2], office_street_address = mclean[0],
                office_city_state_zip = mclean[1], template ="template-one license.docx")
     for df in pages]


    
    #[create_page(page) for page in pages]
    #test = pad_the_df(pages[5])
    