# -*- coding: utf-8 -*-
"""
Created on Sat Apr 11 15:47:44 2020

Found this link to be very helpful: https://pbpython.com/python-word-template.html


@author: CSR1
"""

from __future__ import print_function
from mailmerge import MailMerge
import pandas as pd
import os
from agent_license_automation.generation_functions import fix, pre_process, pad_the_df, template_fields, split_the_data

def one_address_page(df, template, file_name = None, branch_license_number = None, 
                company_license_number = None, p = None):
    '''
    Create the final printed page.  To-do: add the template, sample inputs and sample outputs
    '''

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
        
        phone_1 = row_1['phone_number'],
        phone_2 = row_2['phone_number'],
        phone_3 = row_3['phone_number'],
        phone_4 = row_4['phone_number'],
        phone_5 = row_5['phone_number'],
        phone_6 = row_6['phone_number'],
        phone_7 = row_7['phone_number'],
        phone_8 = row_8['phone_number'],
        phone_9 = row_9['phone_number'],
        phone_10 = row_10['phone_number'],
        phone_11 = row_11['phone_number'],
        phone_12 = row_12['phone_number'],
        )
    
    if p is None:
        p = 'output/one license/'
    if not os.path.exists(p):
        os.mkdir(p)
    
    df = df.loc[df['last_name'] != 'place_holder']
    if file_name is None:
        n = df.shape[0]            
        file_name = df['last_name'].iloc[0] + '_to_' + df['last_name'].iloc[n-1]
    document.write(p + file_name + '.docx')
    print('Done')
    
def two_address_page(df, template, file_name = None, branch_license_number = None, 
                company_license_number = None, p = None):
    '''
    Create the final printed page.  To-do: add the template, sample inputs and sample outputs
    '''

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
            
        state_1_1 = row_1['state_1'],
        state_2_1 = row_2['state_1'],
        state_3_1 = row_3['state_1'],
        state_4_1 = row_4['state_1'],
        state_5_1 = row_5['state_1'],
        state_6_1 = row_6['state_1'],
        state_7_1 = row_7['state_1'],
        state_8_1 = row_8['state_1'],
        state_9_1 = row_9['state_1'],
        state_10_1 = row_10['state_1'],
        state_11_1 = row_11['state_1'],
        state_12_1 = row_12['state_1'],
        
        state_1_2 = row_1['state_2'],
        state_2_2 = row_2['state_2'],
        state_3_2 = row_3['state_2'],
        state_4_2 = row_4['state_2'],
        state_5_2 = row_5['state_2'],
        state_6_2 = row_6['state_2'],
        state_7_2 = row_7['state_2'],
        state_8_2 = row_8['state_2'],
        state_9_2 = row_9['state_2'],
        state_10_2 = row_10['state_2'],
        state_11_2 = row_11['state_2'],
        state_12_2 = row_12['state_2'],
        
        
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
        
        exp_date_1_1 = row_1['expiration_date_1'],
        exp_date_2_1 = row_2['expiration_date_1'],
        exp_date_3_1 = row_3['expiration_date_1'],
        exp_date_4_1 = row_4['expiration_date_1'],
        exp_date_5_1 = row_5['expiration_date_1'],
        exp_date_6_1 = row_6['expiration_date_1'],
        exp_date_7_1 = row_7['expiration_date_1'],
        exp_date_8_1 = row_8['expiration_date_1'],
        exp_date_9_1 = row_9['expiration_date_1'],
        exp_date_10_1 = row_10['expiration_date_1'],
        exp_date_11_1 = row_11['expiration_date_1'],
        exp_date_12_1 = row_12['expiration_date_1'],
        
        exp_date_1_2 = row_1['expiration_date_2'],
        exp_date_2_2 = row_2['expiration_date_2'],
        exp_date_3_2 = row_3['expiration_date_2'],
        exp_date_4_2 = row_4['expiration_date_2'],
        exp_date_5_2 = row_5['expiration_date_2'],
        exp_date_6_2 = row_6['expiration_date_2'],
        exp_date_7_2 = row_7['expiration_date_2'],
        exp_date_8_2 = row_8['expiration_date_2'],
        exp_date_9_2 = row_9['expiration_date_2'],
        exp_date_10_2 = row_10['expiration_date_2'],
        exp_date_11_2 = row_11['expiration_date_2'],
        exp_date_12_2 = row_12['expiration_date_2'],
        
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
        
        lic_num_1_1 = fix(row_1['license_number_1']),
        lic_num_2_1 = fix(row_2['license_number_1']),
        lic_num_3_1 = fix(row_3['license_number_1']),
        lic_num_4_1 = fix(row_4['license_number_1']),
        lic_num_5_1 = fix(row_5['license_number_1']),
        lic_num_6_1 = fix(row_6['license_number_1']),
        lic_num_7_1 = fix(row_7['license_number_1']),
        lic_num_8_1 = fix(row_8['license_number_1']),
        lic_num_9_1 = fix(row_9['license_number_1']),
        lic_num_10_1 = fix(row_10['license_number_1']),
        lic_num_11_1 = fix(row_11['license_number_1']),
        lic_num_12_1 = fix(row_12['license_number_1']),
        
        lic_num_1_2 = fix(row_1['license_number_2']),
        lic_num_2_2 = fix(row_2['license_number_2']),
        lic_num_3_2 = fix(row_3['license_number_2']),
        lic_num_4_2 = fix(row_4['license_number_2']),
        lic_num_5_2 = fix(row_5['license_number_2']),
        lic_num_6_2 = fix(row_6['license_number_2']),
        lic_num_7_2 = fix(row_7['license_number_2']),
        lic_num_8_2 = fix(row_8['license_number_2']),
        lic_num_9_2 = fix(row_9['license_number_2']),
        lic_num_10_2 = fix(row_10['license_number_2']),
        lic_num_11_2 = fix(row_11['license_number_2']),
        lic_num_12_2 = fix(row_12['license_number_2']),
        
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
    
    if p is None:
        p = 'output/two licenses/'
    if not os.path.exists(p):
        os.mkdir(p)
    
    df = df.loc[df['last_name'] != 'place_holder']
    if file_name is None:
        n = df.shape[0]            
        
        file_name = df['last_name'].iloc[0] + '_to_' + df['last_name'].iloc[n-1]
    document.write(p + file_name + '.docx')
    print('Done')


if __name__ == '__main__':
    
    'Generate TRID cards for agents licensed in one state'
    df = pre_process(pd.read_csv("one license.csv"))
    [one_address_page(page, 'template-one license.docx') for page in split_the_data(df)]
    
    'Generate TRID cards for agents licensed in two states'
    df = pre_process(pd.read_csv("two licenses.csv"))
    [two_address_page(page, 'template-two licenses.docx') for page in split_the_data(df)]
    
    'Generate TRID cards for agents licensed in three states, breaking it up into 2-state and 1-state'
    df = pre_process(pd.read_csv("three licenses-1 of 2.csv"))
    [two_address_page(page, 'template-two licenses.docx', p = "output/three licenses/", file_name = "2 of three licenses") for page in split_the_data(df)]
    df = pre_process(pd.read_csv("three licenses-2 of 2.csv"))
    [one_address_page(page, 'template-one license.docx', p = "output/three licenses/", file_name = "1 of three licenses") for page in split_the_data(df)]
    
    
    
    