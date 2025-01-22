# -*- coding: utf-8 -*-
"""
Created on Tue Jan 24 15:26:20 2023

@author: Andy Labella, PhD & Sang Hoon Chong, PhD
"""

import pydicom
import json
import pandas as pd
from datetime import datetime
from math import sqrt
import csv
from pandas import ExcelWriter
import os
import psutil

def ret_all_fl_series(inp_file):
    a = pydicom.filereader.dcmread(inp_file)
    b = a.to_json()
    raw_dcm = json.loads(b)
    
    d = raw_dcm['0040A730']['Value']
    
    res = [i for i in d if not (len(i) < 5)]
    
    paras = [i['0040A730']['Value'] for i in res]
    
    paras = [i for i in paras if not (len(i) < 20)]
    
    return raw_dcm, paras

def para_extract(sub_tree):
    key_list = list(sub_tree.keys())
    para_name = sub_tree['0040A043']['Value'][0]['00080104']['Value']
    para_val_dict = sub_tree[key_list[-1]]

    if 'Value' not in list(para_val_dict.keys()):
        para_val = para_val_dict[list(para_val_dict.keys())[0]]
    elif len(para_val_dict['Value'][0]) == 3 and type(para_val_dict['Value'][0]) is dict:
        para_val = para_val_dict['Value'][0]['00080104']['Value']
    elif len(para_val_dict['Value'][0]) == 2 and type(para_val_dict['Value'][0]) is dict:
        para_val = para_val_dict['Value'][0]['0040A30A']['Value']
    else:
        para_val = para_val_dict['Value'][0]
    
    if isinstance(para_val,str) == True:
        return para_name, para_val
    else:
        return para_name, para_val[0]
    
def calculatephantomAge(examDate,birthDate):

    today = examDate
    age = today.year - birthDate.year - ((today.month, today.day) < (birthDate.month, birthDate.day))
    if age < 1:
        ph_age = 1
    elif age >=1 and age < 5:
        ph_age = 2
    elif age >=5 and age < 10:
        ph_age = 3
    elif age >= 10 and age < 15:
        ph_age = 4
    elif age >= 15 and age < 18:
        ph_age = 5
    else:
        ph_age = 6
    return ph_age

#   Begin Main Processing 

try:

    dicom_file_path = "/Users/macindochi/Library/CloudStorage/Box-Box/BCH/Project/Angio CT Comparison Study/20231030 Document from Dr Maschietto/AS/20191115 XA/AS_XA.dcm"
    
    directory = os.path.dirname(dicom_file_path)
    file_name_with_ext = os.path.basename(dicom_file_path)
    
    file_name = os.path.splitext(file_name_with_ext)[0]
        
    raw_dcm, paras = ret_all_fl_series(dicom_file_path)
    
    dict_all_series = []
    se_all_series = []
    
    for i in paras:
        
        dict1={}
        
        for j in i:
            
            if len(j) == 4:
                
                para_name, para_val = para_extract(j)
                dict1[para_name[0]] = para_val
                
            elif len(j) == 5:
                
                sub_para = j[list(j.keys())[-1]]['Value'][0]
                para_name, para_val = para_extract(sub_para)
                dict1[para_name[0]] = para_val
                
            elif len(j) == 6:
                
                sub_para = j[list(j.keys())[-1]]['Value']
                para_name = j['0040A043']['Value'][0]['00080104']['Value']
    
                for k in sub_para:
                    
                    para_name, para_val = para_extract(k)
                    dict1[para_name[0]] = para_val
                    
            else:
                raise ValueError('A very specific bad thing happened.')
        
        dict_se = pd.Series(dict1)
        se_all_series.append(dict_se)
        dict_all_series.append(dict1)
    
    target_save = directory + '\\' + file_name + '_para.xlsx'
    
    se_all_series_concat = pd.concat(se_all_series,axis=1)
    
    # Use the xlsxwriter engine for ExcelWriter
    with ExcelWriter(target_save, engine='xlsxwriter') as writer:
        # Write DataFrame to Excel
        se_all_series_concat.to_excel(writer, sheet_name='sheetName', na_rep='NaN')
    
        # Access the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['sheetName']
    
        # Adjust column widths dynamically
        for col_idx, col in enumerate(se_all_series_concat.columns):
            # Calculate column width based on content
            column_length = max(
                se_all_series_concat[col].astype(str).map(len).max(),
                len(str(col))
            )
            # Set the column width
            worksheet.set_column(col_idx, col_idx, column_length)
    
    ncirf_all = []
    
    for i in dict_all_series:
        if i['Dose Area Product'] == 0:
            continue
        else:
            ncirf = []
            
            #ID
            k = 100001
            ncirf.append(k)
            k = k + 1
            
            # arm position
            ncirf.append(1)
            
            # phantom age
            if 'Value' in raw_dcm['00100030']:
                ncirf.append(calculatephantomAge(datetime.strptime(i['DateTime Started'].split('.',1)[0],'%Y%m%d%H%M%S').date(),datetime.strptime(raw_dcm['00100030']['Value'][0], '%Y%m%d').date()))
            else:
                ncirf.append(datetime.strptime(i['DateTime Started'].split('.',1)[0],'%Y%m%d%H%M%S').date())
                print('No patient birth date specified - default to EXAM DATE')
            # patient sex
            if 'Value' in raw_dcm['00100040']: 
                if raw_dcm['00100040']['Value'][0] == 'F':
                    ncirf.append(1)
                elif raw_dcm['00100040']['Value'][0] == 'M':
                    ncirf.append(2)
            else:
                ncirf.append(1)
                print('No sex specified in RDSR - default to F')
                
            # kVp
            ncirf.append(round(i['KVP']/10)*10)
            
            # beam quality (WITH PRE-LOADED BEAM QUALITIES, ASSUMING LOWEST SINCE FL)
            if round(i['KVP']/10)*10 == 50:
                ncirf.append(1.89)
                # add kVp-dependent attenuation factor here, implement in DAP
            elif round(i['KVP']/10)*10 == 60:
                ncirf.append(2.25)
            elif round(i['KVP']/10)*10 == 70:
                ncirf.append(2.61)
            elif round(i['KVP']/10)*10 == 80:
                ncirf.append(3.01)
            elif round(i['KVP']/10)*10 == 90:
                ncirf.append(3.38)
            elif round(i['KVP']/10)*10 == 100:
                ncirf.append(3.75)
            elif round(i['KVP']/10)*10 == 110:
                ncirf.append(4.11)
            elif round(i['KVP']/10)*10 == 120:
                ncirf.append(4.53)
            
            # SID
            ncirf.append(i['Distance Source to Isocenter']/10)
            
            # field width at isocenter (cm)
            # field height at isocenter (cm)
            if 'Collimated Field Height' in i and 'Collimated Field Width' in i:
                if i['Collimated Field Area'] == 0 or i['Collimated Field Height'] == 0 or i['Collimated Field Width'] == 0:
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100)
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100)
                else:
                    ncirf.append(i['Collimated Field Width']*100)
                    ncirf.append(i['Collimated Field Height']*100)
            elif 'Collimated Field Height' not in i or 'Collimated Field Width' not in i:
                if i['Collimated Field Area'] == 0:
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100)
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100)
                else:
                    ncirf.append(sqrt(i['Collimated Field Area'])*100)
                    ncirf.append(sqrt(i['Collimated Field Area'])*100)
            
            # DAP (Gy*cm2)
            # add attenuation factor
            ncirf.append(i['Dose Area Product'] * 10000)
            
            # PPA
            ncirf.append(i['Positioner Primary Angle'])
            
            # PSA
            ncirf.append(i['Positioner Secondary Angle'])
            
            if 'Lateral Beam Position' in i and 'Longitudinal Beam Position' in i:
            
            # start with IR14 (Philips)
            
                # Isocenter x (cm)
                ncirf.append(42+(i['Table Lateral Position'] - i['Lateral Beam Position'])/10)
                # Isocenter y (cm)
                ncirf.append((i['Table Height Position'] - i['Distance Source to Isocenter'])/10)
                # Isocenter z (cm)
                ncirf.append(163+(i['Table Longitudinal Position'] - i['Longitudinal Beam Position'])/10)
                # -8.9, 11.5, 7.8
                            
            elif i['Device Name'] == 'AXIS05187': 
                
                    if ncirf[2] == 1:
                # Isocenter x (cm)
                        ncirf.append(12.4+(i['Table Lateral Position'] - 594.2)/10)
                # Isocenter y (cm)
                        ncirf.append(7.2+(i['Table Height Position'] - 151.6)/10)
                # Isocenter z (cm)
                        ncirf.append(39.7+(i['Table Longitudinal Position'] - 5.1)/10)
                    elif ncirf[2] == 2:
                # Isocenter x (cm)
                        ncirf.append(19.5+(i['Table Lateral Position']/10 - 59.42)/10)
                # Isocenter y (cm)
                        ncirf.append(8.5+(i['Table Height Position'] - 151.6)/10)
                # Isocenter z (cm)
                        ncirf.append(67.7+(i['Table Longitudinal Position'] - 5.1)/10)
                    elif ncirf[2] == 3:
                # Isocenter x (cm)
                        ncirf.append(26.2+(i['Table Lateral Position']/10 - 59.42)/10)
                # Isocenter y (cm)
                        ncirf.append(10.1+(i['Table Height Position']/10 - 15.16)/10)
                # Isocenter z (cm)
                        ncirf.append(102.4+(i['Table Longitudinal Position']/10 - 0.51)/10)
                    elif ncirf[2] == 4:
                # Isocenter x (cm)
                        ncirf.append(34.4+(i['Table Lateral Position'] - 594.2)/10)
                # Isocenter y (cm)
                        ncirf.append(11.4+(i['Table Height Position'] - 151.6)/10)
                # Isocenter z (cm)
                        ncirf.append(130.7+(i['Table Longitudinal Position'] - 5.1)/10)
                    elif ncirf[2] == 5:
                # Isocenter x (cm)
                        ncirf.append(44+(i['Table Lateral Position'] - 594.2)/10)
                # Isocenter y (cm)
                        ncirf.append(14.6+(i['Table Height Position'] - 151.6)/10)
                # Isocenter z (cm)
                        ncirf.append(156.3+(i['Table Longitudinal Position'] - 5.1)/10)
                    elif ncirf[2] == 6:
                # Isocenter x (cm)
                        ncirf.append(47.3+(i['Table Lateral Position'] - 594.2)/10)
                # Isocenter y (cm)
                        ncirf.append(15.6+(i['Table Height Position'] - 151.6)/10)
                # Isocenter z (cm)
                        ncirf.append(165.3+(i['Table Longitudinal Position'] - 5.1)/10)
                # -8.9, 11.5, 7.8
                
            else:
            
                if i['Target Region'] == 'Abdomen':
                    if ncirf[2] == 1:
                        ncirf.extend((12.5,6.5,22))
                    elif ncirf[2] == 2:
                        ncirf.extend((19.5,7.5,40))
                    elif ncirf[2] == 3:
                        ncirf.extend((26,8.5,66))
                    elif ncirf[2] == 4:
                        ncirf.extend((34.5,9.5,87))
                    elif ncirf[2] == 5:
                        ncirf.extend((40.5,12.5,104))
                    elif ncirf[2] == 6:
                        ncirf.extend((44,13.5,105))
                elif i['Target Region'] == 'Chest' or i['Target Region'] == 'Heart' or i['Target Region'] == 'Coronary artery':
                    if ncirf[2] == 1:
                        ncirf.extend((12.5,6.5,31))
                    elif ncirf[2] == 2:
                        ncirf.extend((19.5,7.5,52))
                    elif ncirf[2] == 3:
                        ncirf.extend((26,8.5,79))
                    elif ncirf[2] == 4:
                        ncirf.extend((34.5,9.5,105))
                    elif ncirf[2] == 5:
                        ncirf.extend((40.5,12.5,121))
                    elif ncirf[2] == 6:
                        ncirf.extend((44,13.5,121))
                elif i['Target Region'] == 'Head':
                    if ncirf[2] == 1:
                        ncirf.extend((12.5,6.5,42))
                    elif ncirf[2] == 2:
                        ncirf.extend((19.5,7.5,69))
                    elif ncirf[2] == 3:
                        ncirf.extend((26,8.5,101.5))
                    elif ncirf[2] == 4:
                        ncirf.extend((34.5,9.5,131))
                    elif ncirf[2] == 5:
                        ncirf.extend((40.5,12.5,152.5))
                    elif ncirf[2] == 6:
                        ncirf.extend((44,13.5,154.5))
                elif i['Target Region'] == 'Extremity':
                    if ncirf[2] == 1:
                        ncirf.extend((4.3,6.5,11))
                    elif ncirf[2] == 2:
                        ncirf.extend((15.5,10,18))
                    elif ncirf[2] == 3:
                        ncirf.extend((21,12.5,29))
                    elif ncirf[2] == 4:
                        ncirf.extend((28,14.5,38))
                    elif ncirf[2] == 5:
                        ncirf.extend((32,20,45))
                    elif ncirf[2] == 6:
                        ncirf.extend((36,21,45))
                elif i['Target Region'] == 'Entire body':
                    if ncirf[2] == 1:
                        ncirf.extend((12.5,6.5,22))
                    elif ncirf[2] == 2:
                        ncirf.extend((19.5,7.5,40))
                    elif ncirf[2] == 3:
                        ncirf.extend((26,8.5,66))
                    elif ncirf[2] == 4:
                        ncirf.extend((34.5,9.5,87))
                    elif ncirf[2] == 5:
                        ncirf.extend((40.5,12.5,104))
                    elif ncirf[2] == 6:
                        ncirf.extend((44,13.5,105))
                else:
                    raise ValueError('"Target Region" not Defined')
            
            # MC History
            ncirf.append(1000000)
            
            # Number of threads
            
            physical_cores = psutil.cpu_count(logical=False)
            
            ncirf.append(physical_cores)
            #   ncirf.append(8) # for manual designation of the number of computing cores.
            
            ncirf_all.append(ncirf)
            
    #   Save the output file
    
    target_save = directory + '/' + file_name + '.csv'
    
    with open(target_save, 'w') as f:
        fc = csv.writer(f, lineterminator='\n')
        fc.writerows(ncirf_all)
        
except Exception as e:
    print(f"Error: {e}")