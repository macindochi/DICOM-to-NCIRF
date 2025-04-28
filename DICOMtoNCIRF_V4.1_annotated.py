# -*- coding: utf-8 -*-
"""
Created on Tue Jan 24 15:26:20 2023
Ver 4 is an upgraded version from Ver 1.3 programmed by Andy LaBella, PhD
Ver 4.1 includes correction of field width/height accounting for SID and SDD

@author: Sang Hoon Chong, PhD
"""

import pydicom
import json
import pandas as pd
from datetime import datetime
from math import sqrt
import csv
from scipy.interpolate import RegularGridInterpolator
import numpy as np
import os
import psutil
import sys

# Function to read DICOM file and extract relevant fluoroscopy series
def ret_all_fl_series(inp_file):
    a = pydicom.filereader.dcmread(inp_file)
    b = a.to_json()
    raw_dcm = json.loads(b)
    
    d = raw_dcm['0040A730']['Value']
    
    res = [i for i in d if not (len(i) < 5)]
    
    paras = [i['0040A730']['Value'] for i in res]
    
    paras = [i for i in paras if not (len(i) < 20)]
    
    return raw_dcm, paras

# Function to extract a parameter name and value from a sub-tree of the DICOM JSON
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

# Function to calculate phantom age group based on birthdate and exam date
# Returns group 1-6 depending on patient age    
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

# Function to estimate beam quality (HVL) based on kVp and filter information
# Interpolates based on a copper filter HVL database
def estimatebeamquality(kvp, flt_material, flt_thickness):
    
    ncirf_hvl_dict = {
        50: [1.89, 2.8, 3.3, 3.75],
        60: [2.25, 3.42],
        70: [2.61, 4.05, 6.83],
        80: [3.01, 4.61, 5.57, 6.38, 7.7],
        90: [3.38, 5.18],
        100: [3.75, 5.71],
        110: [4.11, 6.18, 7.33, 8.23, 9.68],
        120: [4.53, 6.52]
        }
    
    kvp_ncirf = round(kvp/10)*10
    
    if 'copper' in flt_material.lower():
        
        #   Designate the path to the HVL database
        beam_quality_path = "[Path to HVL Database]/hvl_copper_filter.xlsx"
    # Add 'elif' or 'else' statement to include another kind of filter material.
    
    hvl_db = pd.read_excel(beam_quality_path, index_col = 0)
    
    thickness_array = hvl_db.index.to_numpy()
    kvp_array = hvl_db.columns.to_numpy().astype(float)
    hvl_array = hvl_db.values
    
    interpolator = RegularGridInterpolator((thickness_array, kvp_array), hvl_array)
    
    point = np.array([[kvp, flt_thickness]])
    hvl = interpolator(point)[0]
    
    ncirf_hvl_array = np.array(ncirf_hvl_dict[kvp_ncirf])
    
    abs_diff_hvl_array = np.abs( ncirf_hvl_array - hvl )
    
    hvl_ncirf = ncirf_hvl_array[ abs_diff_hvl_array == np.min(abs_diff_hvl_array) ][0]
    
    return kvp_ncirf, hvl_ncirf

# Function to pre-set isocenter coordinates based on target body region and phantom age
# Used as a backup method if user input is missing
def presetisocenter(target_region, phantom_age_group):
    region_coords = {
        'Abdomen': {
            1: [12.5, 6.5, 22], 2: [19.5, 7.5, 40], 3: [26, 8.5, 66],
            4: [34.5, 9.5, 87], 5: [40.5, 12.5, 104], 6: [44, 13.5, 105]
        },
        'Chest': {
            1: [12.5, 6.5, 31], 2: [19.5, 7.5, 52], 3: [26, 8.5, 79],
            4: [34.5, 9.5, 105], 5: [40.5, 12.5, 121], 6: [44, 13.5, 121]
        },
        'Head': {
            1: [12.5, 6.5, 42], 2: [19.5, 7.5, 69], 3: [26, 8.5, 101.5],
            4: [34.5, 9.5, 131], 5: [40.5, 12.5, 152.5], 6: [44, 13.5, 154.5]
        },
        'Extremity': {
            1: [4.3, 6.5, 11], 2: [15.5, 10, 18], 3: [21, 12.5, 29],
            4: [28, 14.5, 38], 5: [32, 20, 45], 6: [36, 21, 45]
        },
    }

    # 'Heart' and 'Coronary artery' use Chest coordinates
    if target_region in ['Heart', 'Coronary artery']:
        target_region = 'Chest'
    elif target_region == 'Entire body':
        target_region = 'Abdomen'

    try:
        coord = region_coords[target_region][phantom_age_group]
    except KeyError:
        raise ValueError(f'Invalid target region or phantom age group: {target_region}, {phantom_age_group}')
    
    iso_x, iso_y, iso_z = coord
    return iso_x, iso_y, iso_z


# Begin Main Processing block (protected by try-except)
try:

    #   Please designate the dicom file path here before running the entire script.
    dicom_file_path = "/Users/macindochi/Library/CloudStorage/Box-Box/BCH/Project/Angio CT Comparison Study/20231030 Document from Dr Maschietto/AS/20191115 XA/AS_XA.dcm" 
    
    directory = os.path.dirname(dicom_file_path)
    file_name_with_ext = os.path.basename(dicom_file_path)
    
    file_name = os.path.splitext(file_name_with_ext)[0]
        
    raw_dcm, paras = ret_all_fl_series(dicom_file_path)
    
    dict_all_series = []
    se_all_series = []
    
    # Patient demographic info
    pat_birth_date = datetime.strptime(raw_dcm['00100030']['Value'][0], '%Y%m%d')
    pat_study_date = datetime.strptime(raw_dcm['00080020']['Value'][0], '%Y%m%d')
    phantom_group = calculatephantomAge(pat_study_date, pat_birth_date)
    
    # Check birthdate info; if missing, request user input
    if 'Value' in raw_dcm['00100030']:
        pat_birth_date = datetime.strptime(raw_dcm['00100030']['Value'][0], '%Y%m%d')
        pat_study_date = datetime.strptime(raw_dcm['00080020']['Value'][0], '%Y%m%d')
        phantom_group = calculatephantomAge(pat_study_date, pat_birth_date)
    else:
        phantom_group = int(input("No patient birth date specified. Please choose phantom age group.\nEnter 1 for age < 1\nEnter 2 for 1<= age < 5\nEnter 3 for 5<= age < 10/nEnter 4 for 10<= age < 15/nEnter 5 for 15<= age < 18/nEnter 6 for age >= 18 :"))
    
    # if raw_dcm['00100040']['Value'][0] == 'F':
    #     patient_sex = 'female'
    # elif raw_dcm['00100040']['Value'][0] == 'M':
    #     patient_sex = 'male'
    # else:
    #     print('Patient sex is not assigned properly. The execution stops.')
    #     sys.exit()
    
    # Patient sex
    if 'Value' in raw_dcm['00100040']: 
        if raw_dcm['00100040']['Value'][0] == 'F':
            patient_sex = 1
        elif raw_dcm['00100040']['Value'][0] == 'M':
            patient_sex = 2
    else:
        patient_sex = 1 # Default to female
        print('No sex specified in RDSR - default to F')
    
    arm_position = int(input("Please choose phantom posture depending on arm position. 1 = Arm-raised, 2 = Arm-lowered, 3 = Arm-rotated.: "))
    
    # Interpret arm position into text
    if arm_position == 1:
        position_statement = 'raised'
    elif arm_position == 2:
        position_statement = 'lowered'
    elif arm_position == 3:
        position_statement = 'rotated'
    else:
        print('The arm position is not determined clearly. The execution is stopped now.')
        sys.exit()
    
    # Confirm to user phantom and position
    if patient_sex == 1:
        print(f'The patient is female, and the phantom group is {phantom_group} with arms {position_statement}.')
    else:
        print(f'The patient is male, and the phantom group is {phantom_group} with arms {position_statement}.')
    
    # Ask for isocenter coordinates manually
    print('Open NCIRF to decide the isocenter coordinate and enter them in the following order: x, y, and z.')
    
    iso_x = input('Enter the coordinate for x in cm: ')
    iso_y = input('Enter the coordinate for y in cm: ')
    iso_z = input('Enter the coordinate for z in cm: ')
    
    # Convert to float if not empty
    if iso_x.strip():
        iso_x = float(iso_x)
        
    if iso_y.strip():
        iso_y = float(iso_y)
        
    if iso_z.strip():
        iso_z = float(iso_z)
    
    # Other user inputs
    patient_id = int(input('Enter a patient ID. (This is not equivalent to the MRN number): '))
    
    history_num = input('Enter the number of photon history for each irradiation event. If nothing is entered the default nunber is 10M: ')
    
    if history_num.strip():
        history_num = int(history_num)
    else:
        history_num = 10000000
    
    cpu_core_num = input('Enter the number of CPU cores for simulation process. The maximum number of available cores will be used if nothing is entered: ')
    
    if cpu_core_num.strip():
        cpu_core_num = int(cpu_core_num)
    else:
        cpu_core_num = psutil.cpu_count(logical=False)
    
    # Process each irradiation event series
    for i in paras:
        
        # i = paras[4]
        
        dict1={}
        
        for j in i:
            # Extract depending on length of j
            if len(j) == 4:
                
                para_name, para_val = para_extract(j)
                dict1[para_name[0]] = para_val
                
            elif len(j) == 5:
                
                if j['0040A043']['Value'][0]['00080104']['Value'][0] == 'X-Ray Filters':
                    sub_para = j[list(j.keys())[-1]]['Value']
                    para_name = j['0040A043']['Value'][0]['00080104']['Value']
        
                    for k in sub_para:
                        
                        para_name, para_val = para_extract(k)
                        dict1[para_name[0]] = para_val
                    
                else:
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
    
    # (Optional) save extracted parameters to Excel for debugging

    
    ################################################################################################
    # # Activate this part of code to generate an Excel output with the available dictionary data.
    
    # target_save = directory + '\\' + file_name + '_para.xlsx'
    
    # se_all_series_concat = pd.concat(se_all_series,axis=1)
    
    # # Extracted parameters (more than needed for NCIRF) are saved for troubleshooting.
    # # Use the xlsxwriter engine for ExcelWriter
    # with ExcelWriter(target_save, engine='xlsxwriter') as writer:
    #     # Write DataFrame to Excel
    #     se_all_series_concat.to_excel(writer, sheet_name='sheetName', na_rep='NaN')
    
    #     # Access the workbook and worksheet
    #     workbook = writer.book
    #     worksheet = writer.sheets['sheetName']
    
    #     for col_idx, col in enumerate(se_all_series_concat.columns):
    #         # Calculate column width based on content
    #         # column_length = max(
    #         #     se_all_series_concat[col].astype(str).map(len).max(),
    #         #     len(str(col))
    #         # )
    #         # Set the column width
    #         column_length = 30
    #         worksheet.set_column(col_idx, col_idx, column_length)
    #################################################################################################
    
    # Now prepare NCIRF batch input rows
    ncirf_all = []    
    
    for i in dict_all_series:
        
        # i = dict_all_series[0]
        
        if i['Dose Area Product'] == 0:
            continue
        else:
            ncirf = []
            
            #ID
            ncirf.append(patient_id)
            
            # arm position
            ncirf.append(arm_position)
            
            # phantom age group
            ncirf.append(phantom_group)

            # phantom sex
            ncirf.append(patient_sex)
                
            # kVp & beam quality
            filter_material = i['X-Ray Filter Material']
            
            filter_min_thickness = i['X-Ray Filter Thickness Minimum']
            filter_max_thickness = i['X-Ray Filter Thickness Maximum']
            
            if filter_min_thickness != filter_max_thickness:
                raise ValueError('X-ray filter is not flat.')
            else:
                filter_thickness = filter_min_thickness
            
            kvp_ncirf, hvl_ncirf = estimatebeamquality(i['KVP'], filter_material, filter_thickness)
            
            # kVp
            ncirf.append(kvp_ncirf)
            
            # HVL
            ncirf.append(hvl_ncirf)
            
            # SID
            sid = i['Distance Source to Isocenter']
            ncirf.append(sid/10)
            
            # field width at isocenter (cm)
            # field height at isocenter (cm)
            
            if 'Distance Source to Reference Point' in i:
                srd = i['Distance Source to Reference Point']                
            else:
                srd = sid - 150
            
            sdd = i['Distance Source to Detector']
            cf_srd = sid/srd #  Correction factor for reference point to isocenter point.
            cf_sdd = sid/sdd #  Correction factor for image recepter point to isocenter point.
            
            if 'Collimated Field Height' in i and 'Collimated Field Width' in i:
                if i['Collimated Field Area'] == 0 or i['Collimated Field Height'] == 0 or i['Collimated Field Width'] == 0:
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100*cf_srd)
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100*cf_srd)
                else:
                    ncirf.append(i['Collimated Field Width']/10*cf_sdd) #  /10 because Collimated Field Width in mm
                    ncirf.append(i['Collimated Field Height']/10*cf_sdd)
            elif 'Collimated Field Height' not in i or 'Collimated Field Width' not in i:
                if i['Collimated Field Area'] == 0:
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100*cf_srd)
                    ncirf.append(sqrt(i['Dose Area Product']/i['Dose (RP)'])*100*cf_srd)
                else:
                    ncirf.append(sqrt(i['Collimated Field Area'])*100*cf_sdd) # *100 because Collimated Field Area in m^2.
                    ncirf.append(sqrt(i['Collimated Field Area'])*100*cf_sdd)
            
            # DAP (Gy*cm2)
            # add attenuation factor 
            ncirf.append(i['Dose Area Product'] * 10000)
            
            # Positioner Primary Angle (PPA)
            ncirf.append(i['Positioner Primary Angle'])
            
            # Positioner Secondary Angle (PSA)
            ncirf.append(i['Positioner Secondary Angle'])
            
            # Isocenter Coordinate
            
            #######################################################################################
            # # If this if-atatement is activated, the preset value deduced from the target region 
            # # with arm position raised when any of the user input coordiates is found empty,   
            # if iso_x == '' or iso_y == '' or iso_z == '':           
            #     iso_x, iso_y, iso_z = presetisocenter(i['Target Region'], ncirf[2])
            #     ncirf[1] = 1
            #######################################################################################
            
            ncirf.extend((iso_x, iso_y, iso_z))
            
            # MC History
            ncirf.append(history_num)
            
            # Number of threads
            ncirf.append(cpu_core_num)

            # Add the parameters in a row                    
            ncirf_all.append(ncirf)
            
    #   Save the output file
    target_save = directory + '/' + file_name + '.csv'
    
    with open(target_save, 'w') as f:
        fc = csv.writer(f, lineterminator='\n')
        fc.writerows(ncirf_all)
        
except Exception as e:
    print(f"Error: {e}")
