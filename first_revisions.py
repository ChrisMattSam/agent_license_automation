# -*- coding: utf-8 -*-
"""
Created on Sat Apr 18 12:54:16 2020

@author: CSR1
"""

from __future__ import print_function
from mailmerge import MailMerge
import pandas as pd
import numpy as np
import os
from agent_license_automation.generation_functions import fix, pre_process, pad_the_df, template_fields, split_the_data

#def two_address_page(df, template, file_name = None, branch_license_number = None, 
#                company_license_number = None, p = None):
df = pre_process(pd.read_csv("all_agent_info.csv"))

print('Creating the page of contact cards...')
template = 'new template.docx'
document = MailMerge(template)
print(sorted(document.get_merge_fields()))

df = df.head(12)
for col in ['_1', '_2', '_3']:
    df.loc[~df['license_number'+col].isna(),['state'+col]] = df['state'+col] + ' License #: '
    df.loc[~df['license_number'+col].isna(),['expiration_date'+col]] = ' Expiration Date: ' + df['expiration_date'+col]
    
df.replace(np.nan, '', inplace = True)
df.replace('nan', '', inplace = True)



l = df.reset_index()
box_1 = l.iloc[0]
box_2 = l.iloc[1]
box_3 = l.iloc[2]
box_4 = l.iloc[3]
box_5 = l.iloc[4]
box_6 = l.iloc[5]
box_7 = l.iloc[6]
box_8 = l.iloc[7]
box_9 = l.iloc[8]
box_10 = l.iloc[9]
box_11 = l.iloc[10]
box_12 = l.iloc[11]



document = MailMerge(template)
document.merge(
    
    #box 1
    first_1 = box_1['first_name'],
    last_1 = box_1['last_name'],
    
    state_1_1 = box_1['state_1'],
    lic_num_1_1 = fix(box_1['license_number_1']),
    exp_date_1_1 = box_1['expiration_date_1'],
    
    state_1_2 = box_1['state_2'],
    lic_num_1_2 = fix(box_1['license_number_2']),
    exp_date_1_2 = box_1['expiration_date_2'],
    
    state_1_3 = box_1['state_3'],
    lic_num_1_3 = fix(box_1['license_number_3']),
    exp_date_1_3 = box_1['expiration_date_3'],

    #box 2
    first_2 = box_2['first_name'],
    last_2 = box_2['last_name'],
    
    state_2_1 = box_2['state_1'],
    lic_num_2_1 = fix(box_2['license_number_1']),
    exp_date_2_1 = box_2['expiration_date_1'],
    
    state_2_2 = box_2['state_2'],
    lic_num_2_2 = fix(box_2['license_number_2']),
    exp_date_2_2 = box_2['expiration_date_2'],
    
    state_2_3 = box_2['state_3'],
    lic_num_2_3 = fix(box_2['license_number_3']),
    exp_date_2_3 = box_2['expiration_date_3'],

    #box 3
    first_3 = box_3['first_name'],
    last_3 = box_3['last_name'],
    
    state_3_1 = box_3['state_1'],
    lic_num_3_1 = fix(box_3['license_number_1']),
    exp_date_3_1 = box_3['expiration_date_1'],
    
    state_3_2 = box_3['state_2'],
    lic_num_3_2 = fix(box_3['license_number_2']),
    exp_date_3_2 = box_3['expiration_date_2'],
    
    state_3_3 = box_3['state_3'],
    lic_num_3_3 = fix(box_3['license_number_3']),
    exp_date_3_3 = box_3['expiration_date_3'],

    #box 4
    first_4 = box_4['first_name'],
    last_4 = box_4['last_name'],
    
    state_4_1 = box_4['state_1'],
    lic_num_4_1 = fix(box_4['license_number_1']),
    exp_date_4_1 = box_4['expiration_date_1'],
    
    state_4_2 = box_4['state_2'],
    lic_num_4_2 = fix(box_4['license_number_2']),
    exp_date_4_2 = box_4['expiration_date_2'],
    
    state_4_3 = box_4['state_3'],
    lic_num_4_3 = fix(box_4['license_number_3']),
    exp_date_4_3 = box_4['expiration_date_3'])
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    #)


document.write('output/revision_test.docx')
print('Done')