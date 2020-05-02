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
from agent_license_automation.generation_functions import fix, pre_process, pad_the_df, split_the_data

def format_to_print(df):
    for col in ['_1', '_2', '_3']:
        '''create copy of original state values because will alter those values to be "{state} License #:" and will
        need these values to append to the expiration date'''
        df['original_state'+col] = df['state'+col]

    for col in ['_1', '_2', '_3']:
        #Precede the chars "License #:" before the actual license number value
        df.loc[df['license_number'+col].str.len() >3,['state'+col]] = df['state'+col] + ' License #: '
        #If a license number doesnt have an expiration date, add "missing" to the expiration date"
        df.loc[(df['license_number'+col].str.len() > 3) & (df['expiration_date'+col].isna()), ['expiration_date'+col]] = 'missing'
        #If an agent is licensed in 3 states, precede the expiration date value with "Exp Date:"
        df.loc[df['license_number_3'].str.len()>3,['expiration_date'+col]] = df['original_state'+col] + ' Exp Date: ' + df['expiration_date'+col]

    #If an agent is licensed in 2 states, preced the expiration date value with "Expiration Date:"
    for col in ['_1', '_2']:
        df.loc[(df['license_number'+col].str.len()>3) & (df['license_number_3'].str.len() <4 ),
               ['expiration_date' + col]] = df['original_state'+col] + ' Expiration Date: ' + df['expiration_date'+col]
    
    '''if an agent is licensed in one state, move all "state 1" values to be "state 2" values because the template will list their info
    closer to the top of the card, as opposed to in the middle of the card''' 
    one_license = df.loc[(df['license_number_2'] == 'nan') & (df['license_number_3']== 'nan')]
    df.drop(one_license.index, inplace = True)
    for col in ['license_number_', 'state_','expiration_date_']:
        one_license[col+'2'] = one_license[col+'1']
        one_license[col+'1'] = 'nan'
    df = pd.concat([df, one_license])
    
    'based on what office they work in, print that state and branch license number on the back of the TRID card'
    df['office_state'] = 'VA'
    df.loc[df['office'] == 'rockville',['office_state']] = 'MD'
    df['office_state_license_number'] = '0226004375'
    df.loc[df['office_state'] == 'MD',['office_state_license_number']] = '640891'
    df.loc[df['office_state'] == 'DC', ['office_state_license_number']] = 'RE0200200243'
    
    df['office_state'] = df['office_state'] + ' Firm License #'
    
    'remove nan so that the final output doesnt have "nan" in it'
    df.replace(np.nan, '', inplace = True)
    df.replace('nan', '', inplace = True)
    
    df.drop(columns = [i for i in df.columns if 'original' in i], inplace = True)
    df.sort_values(by = 'last_name', inplace = True)
    return df

def make_trid_cards(df, output_dir, template = 'new template.docx', doc_name = None):
    print('Creating the page of contact cards...')
    
    document = MailMerge(template)
    #print(document.get_merge_fields())
    
    df = pad_the_df(df)
    box_1 = df.iloc[0]
    box_2 = df.iloc[1]
    box_3 = df.iloc[2]
    box_4 = df.iloc[3]
    box_5 = df.iloc[4]
    box_6 = df.iloc[5]
    box_7 = df.iloc[6]
    box_8 = df.iloc[7]
    box_9 = df.iloc[8]
    box_10 = df.iloc[9]
    box_11 = df.iloc[10]
    box_12 = df.iloc[11]
    
    document = MailMerge(template)
    document.merge(
        
        #box 1
        office_state_lic_num_1 = box_1['office_state_license_number'],
        office_state_1 = box_1['office_state'],
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
        office_state_lic_num_2 = box_2['office_state_license_number'],
        office_state_2 = box_2['office_state'],
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
        office_state_lic_num_3 = box_3['office_state_license_number'],
        office_state_3 = box_3['office_state'],
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
        office_state_lic_num_4 = box_4['office_state_license_number'],
        office_state_4 = box_4['office_state'],
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
        exp_date_4_3 = box_4['expiration_date_3'],
        
        #box 5
        office_state_lic_num_5 = box_5['office_state_license_number'],
        office_state_5 = box_5['office_state'],
        first_5 = box_5['first_name'],
        last_5 = box_5['last_name'],
        
        state_5_1 = box_5['state_1'],
        lic_num_5_1 = fix(box_5['license_number_1']),
        exp_date_5_1 = box_5['expiration_date_1'],
        
        state_5_2 = box_5['state_2'],
        lic_num_5_2 = fix(box_5['license_number_2']),
        exp_date_5_2 = box_5['expiration_date_2'],
        
        state_5_3 = box_5['state_3'],
        lic_num_5_3 = fix(box_5['license_number_3']),
        exp_date_5_3 = box_5['expiration_date_3'],
        
        #box 6
        office_state_lic_num_6 = box_6['office_state_license_number'],
        office_state_6 = box_6['office_state'],
        first_6 = box_6['first_name'],
        last_6 = box_6['last_name'],
        
        state_6_1 = box_6['state_1'],
        lic_num_6_1 = fix(box_6['license_number_1']),
        exp_date_6_1 = box_6['expiration_date_1'],
        
        state_6_2 = box_6['state_2'],
        lic_num_6_2 = fix(box_6['license_number_2']),
        exp_date_6_2 = box_6['expiration_date_2'],
        
        state_6_3 = box_6['state_3'],
        lic_num_6_3 = fix(box_6['license_number_3']),
        exp_date_6_3 = box_6['expiration_date_3'],
        
        #box 7
        office_state_lic_num_7 = box_7['office_state_license_number'],
        office_state_7 = box_7['office_state'],
        first_7 = box_7['first_name'],
        last_7 = box_7['last_name'],
        
        state_7_1 = box_7['state_1'],
        lic_num_7_1 = fix(box_7['license_number_1']),
        exp_date_7_1 = box_7['expiration_date_1'],
        
        state_7_2 = box_7['state_2'],
        lic_num_7_2 = fix(box_7['license_number_2']),
        exp_date_7_2 = box_7['expiration_date_2'],
        
        state_7_3 = box_7['state_3'],
        lic_num_7_3 = fix(box_7['license_number_3']),
        exp_date_7_3 = box_7['expiration_date_3'],
        
        #box 8
        office_state_lic_num_8 = box_8['office_state_license_number'],
        office_state_8 = box_8['office_state'],
        first_8 = box_8['first_name'],
        last_8 = box_8['last_name'],
        
        state_8_1 = box_8['state_1'],
        lic_num_8_1 = fix(box_8['license_number_1']),
        exp_date_8_1 = box_8['expiration_date_1'],
        
        state_8_2 = box_8['state_2'],
        lic_num_8_2 = fix(box_8['license_number_2']),
        exp_date_8_2 = box_8['expiration_date_2'],
        
        state_8_3 = box_8['state_3'],
        lic_num_8_3 = fix(box_8['license_number_3']),
        exp_date_8_3 = box_8['expiration_date_3'],
        
        #box 9
        office_state_lic_num_9 = box_9['office_state_license_number'],
        office_state_9 = box_9['office_state'],
        first_9 = box_9['first_name'],
        last_9 = box_9['last_name'],
        
        state_9_1 = box_9['state_1'],
        lic_num_9_1 = fix(box_9['license_number_1']),
        exp_date_9_1 = box_9['expiration_date_1'],
        
        state_9_2 = box_9['state_2'],
        lic_num_9_2 = fix(box_9['license_number_2']),
        exp_date_9_2 = box_9['expiration_date_2'],
        
        state_9_3 = box_9['state_3'],
        lic_num_9_3 = fix(box_9['license_number_3']),
        exp_date_9_3 = box_9['expiration_date_3'],
        
        #box 10
        office_state_lic_num_10 = box_10['office_state_license_number'],
        office_state_10 = box_10['office_state'],
        first_10 = box_10['first_name'],
        last_10 = box_10['last_name'],
        
        state_10_1 = box_10['state_1'],
        lic_num_10_1 = fix(box_10['license_number_1']),
        exp_date_10_1 = box_10['expiration_date_1'],
        
        state_10_2 = box_10['state_2'],
        lic_num_10_2 = fix(box_10['license_number_2']),
        exp_date_10_2 = box_10['expiration_date_2'],
        
        state_10_3 = box_10['state_3'],
        lic_num_10_3 = fix(box_10['license_number_3']),
        exp_date_10_3 = box_10['expiration_date_3'],
        
        #box 11
        office_state_lic_num_11 = box_11['office_state_license_number'],
        office_state_11 = box_11['office_state'],
        first_11 = box_11['first_name'],
        last_11 = box_11['last_name'],
        
        state_11_1 = box_11['state_1'],
        lic_num_11_1 = fix(box_11['license_number_1']),
        exp_date_11_1 = box_11['expiration_date_1'],
        
        state_11_2 = box_11['state_2'],
        lic_num_11_2 = fix(box_11['license_number_2']),
        exp_date_11_2 = box_11['expiration_date_2'],
        
        state_11_3 = box_11['state_3'],
        lic_num_11_3 = fix(box_11['license_number_3']),
        exp_date_11_3 = box_11['expiration_date_3'],
        
        #box 12
        office_state_lic_num_12 = box_12['office_state_license_number'],
        office_state_12 = box_12['office_state'],
        first_12 = box_12['first_name'],
        last_12 = box_12['last_name'],
        
        state_12_1 = box_12['state_1'],
        lic_num_12_1 = fix(box_12['license_number_1']),
        exp_date_12_1 = box_12['expiration_date_1'],
        
        state_12_2 = box_12['state_2'],
        lic_num_12_2 = fix(box_12['license_number_2']),
        exp_date_12_2 = box_12['expiration_date_2'],
        
        state_12_3 = box_12['state_3'],
        lic_num_12_3 = fix(box_12['license_number_3']),
        exp_date_12_3 = box_12['expiration_date_3'])
    
    n = df.shape[0]            
    file_name = df['last_name'].iloc[0] + '_to_' + df['last_name'].iloc[n-1]
    if doc_name is not None:
        if type(doc_name) is str: file_name = doc_name
    document.write(output_dir + '/' +file_name + '.docx')
    print('Done')


if __name__ == "__main__":
    df = format_to_print(pre_process(pd.read_csv("all_agent_info.csv")))
    
    dfs = split_the_data(df)
    
    path = '//192.168.50.245/Staff/2018/CSR Front Desk/MCL-RV-ARL Shared Folder/TRID Cards 2020'
    [make_trid_cards(df, template = 'new template.docx', output_dir = path) for df in dfs]
    
    sample_df = pd.concat([df.loc[df['office'] == 'mclean'].head(9),
                           df.loc[df['office'] == 'rockville'].head(3)])
    
    make_trid_cards(sample_df, template = 'new template.docx',output_dir = path, doc_name = '0-sample for Suzzette')        
        
        
        