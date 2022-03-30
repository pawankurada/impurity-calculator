import os
import re
from tqdm import tqdm
from datetime import datetime
import xlrd
import xlutils.copy
import glob

import pandas as pd
import numpy as np
import camelot
import tabula

# template sheet
assay_template_input = xlrd.open_workbook(os.path.join(os.getcwd(), "data", "Templates",'Assay-template.xls'), formatting_info=True)
assay_template = xlutils.copy.copy(assay_template_input)

# Table headers
# chrom_headers = ['Peak#','Name','Ret. Time','Area','Area%','RRT']
# area_headers = ['Title', 'Ret. Time', 'Area', 'Area%', 'NTP', 'Tailing Factor']
chrom_headers_shimadzu = ['Name','Area']
area_headers_shimadzu = ['Title','Area',]
chrom_headers_empower = ['Name','Area\n(µV*sec)']
area_headers_empower = ['SampleName','Area' ]


def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell

def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx

def calc_results (df_peak, compound, average_area, constant_1, constant_2, unit):
    area = float(df_peak['Area'][df_peak['Name'].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0])
    assay = (area/average_area) * constant_1 * constant_2 * unit
    return assay

def shift_row_to_top(df, index_to_shift):
    idx = df.index.tolist()
    idx.remove(index_to_shift)
    df = df.reindex([index_to_shift] + idx)
    return df

def table_extratcor(tables, headers):
    df_result_table =''
    result_tables = []
    for table in tables:
        df_table = table.df
        search = df_table.where(df_table==headers[0]).dropna(how='all').dropna(axis=1)
        inx = list(search.index)
        if(inx):
            inx= inx[0]
            new_header = df_table.iloc[inx]
            new_start_inx = inx+1
            df_table = df_table[new_start_inx:]
            df_table.columns = new_header
            df_table = df_table[headers]
            result_tables.append(df_table)
        else:
            continue
    try:
        df_result_table = pd.concat(result_tables, ignore_index=True)
    except ValueError as ve:
        print("No tables/values found in this file\n")
        return pd.DataFrame()

    return df_result_table

def table_extratcor_with_range(tables, headers, rt_range):
    df_result_table =''
    result_tables = []
    for table in tables:
        df_table = table.df
        search = df_table.where(df_table==headers[0]).dropna(how='all').dropna(axis=1)
        inx = list(search.index)
        if(inx):
            inx= inx[0]
            new_header = df_table.iloc[inx]
            new_start_inx = inx+1
            df_table = df_table[new_start_inx:]
            df_table.columns = new_header
            df_table = df_table[headers]
            sample_rt = float(df_table['Ret. Time'].values.tolist()[0])
            if(sample_rt >= rt_range[0] and sample_rt <= rt_range[1]):
                result_tables.append(df_table)
            else:
                continue
        else:
            continue
    try:
        df_result_table = pd.concat(result_tables, ignore_index=True)
    except ValueError as ve:
        print("No tables/values found in this file\n")
        return pd.DataFrame()

    return df_result_table

def fill_rs_sheet(output_sheet, df_area_table, df_peak_table, sample_input_list, input_list):
    average_area = float(df_area_table["Area"][df_area_table["Title"] == "Average"].values.tolist()[0])
    area_input = list(df_area_table['Area'])

    #poject name
    setOutCell(output_sheet, 2, 3, input_list[15])
    #Date
    setOutCell(output_sheet, 2, 4, input_list[11])
    #Method
    setOutCell(output_sheet, 2, 5, input_list[12])
    # WS ID No.
    setOutCell(output_sheet, 1, 10, input_list[13])
    # potency
    setOutCell(output_sheet, 3, 10, input_list[10])
    # use before date
    setOutCell(output_sheet, 5, 10, input_list[14])
    # Average area
    setOutCell(output_sheet, 7, 10, average_area)
    # std_wt
    setOutCell(output_sheet, 2, 11, input_list[0])
    #  v1
    setOutCell(output_sheet, 2, 12, input_list[1])
    # v2
    setOutCell(output_sheet, 4, 11, input_list[2])
    #  v3
    setOutCell(output_sheet, 4, 12, input_list[3])
    #  v4
    setOutCell(output_sheet, 6, 11,  input_list[4])
    # v5
    setOutCell(output_sheet, 6, 12, input_list[5])
    # v6
    setOutCell(output_sheet, 8, 11, input_list[6])
    # v7
    setOutCell(output_sheet, 8, 12, input_list[7])
    # factor
    setOutCell(output_sheet, 9, 11, input_list[8])
    # factor
    setOutCell(output_sheet, 9, 12, input_list[9])

#     areas
    setOutCell(output_sheet, 12, 5, area_input[0])
    setOutCell(output_sheet, 12, 6, area_input[1])
    setOutCell(output_sheet, 12, 7, area_input[2])
    setOutCell(output_sheet, 12, 8, area_input[3])
    setOutCell(output_sheet, 12, 9, area_input[4])
    setOutCell(output_sheet, 12, 10, area_input[5])
    setOutCell(output_sheet, 12, 11, area_input[6])
    setOutCell(output_sheet, 12, 12, area_input[7])
    setOutCell(output_sheet, 12, 13, area_input[8])
#   Impurity table
    table_row = 17
    for index, row in df_peak_table.iterrows():
        if(table_row > 216):
            break
        # AR NO
        setOutCell(output_sheet, 1, table_row, '')
        # Batch NO
        setOutCell(output_sheet, 2, table_row, row[0] )
        # Condition
        setOutCell(output_sheet, 3, table_row, row[1])
        # Label Claim
        setOutCell(output_sheet, 9, table_row, row[6])
        # per unit
        setOutCell(output_sheet, 10, table_row, row[7])
        # sample_wt
        setOutCell(output_sheet, 5, table_row, row[2])
        #  v1
        setOutCell(output_sheet, 6, table_row, row[3])
        # v2
        setOutCell(output_sheet, 7, table_row, row[4])
        #  v3
        setOutCell(output_sheet, 8, table_row, row[5])
        # Area
        setOutCell(output_sheet, 11, table_row, row[8])
        #  Assay%
        setOutCell(output_sheet, 12, table_row, row[9])
        table_row +=1

def initiate_report_creation(compound,concentration, df_sample_prep, chrom_inputs, area_input, input_list, rt_range, software):
    cond_1 = df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)
    cond_2 = df_sample_prep["concentration"] == concentration
    sample_qty = df_sample_prep['sample quantity'][cond_1 & cond_2].values.tolist()[0]
    sample_v1 = df_sample_prep['v1'][cond_1 & cond_2].values.tolist()[0]
    sample_v2 = df_sample_prep['v2'][cond_1 & cond_2].values.tolist()[0]
    sample_v3 = df_sample_prep['v3'][cond_1 & cond_2].values.tolist()[0]
    label_claim = df_sample_prep['label claim'][cond_1 & cond_2].values.tolist()[0]
    unit = df_sample_prep['per unit'][cond_1 & cond_2].values.tolist()[0]
    constant_1 = (input_list[0]/input_list[1]) * (input_list[2]/input_list[3]) * (input_list[4]/input_list[5])*(input_list[6]/input_list[7]) * (input_list[8]/input_list[9])
    constant_2 = (sample_v1/sample_qty) * (sample_v3/sample_v2) * (input_list[10]/label_claim)

    # area table extraction
    tables = camelot.read_pdf(area_input, pages= 'all', line_scale =30)
    area_headers = area_headers_shimadzu if software == 'Shimadzu' else area_headers_empower
    if(rt_range):
        df_area_table = table_extratcor_with_range(tables, ['Title','Area', 'Ret. Time'], rt_range)
    else:
        df_area_table = table_extratcor(tables, area_headers)
    df_area_table.columns = area_headers_shimadzu
    df_area_table['Title'] = ['Standard Solution_01','Standard Solution_02','Standard Solution_03','Standard Solution_04','Standard Solution_05','Standard Solution_06','Average', '%RSD','Standard Deviation']
    df_area_table = df_area_table[['Title','Area']]
    average_area = float(df_area_table["Area"][df_area_table["Title"] == "Average"].values.tolist()[0])


    batch_size = len(chrom_inputs)
    worksheets = assay_template._Workbook__worksheets
    peak_master = []
    chrom_headers = chrom_headers_shimadzu if software == 'Shimadzu' else chrom_headers_empower
    for index, chrom_input in enumerate(chrom_inputs):
        worksheet_name = chrom_input.split("\\")[-1].strip(".pdf")
        print(worksheet_name)
        # peak tables extratcion
        tables = camelot.read_pdf(chrom_input, pages= 'all', line_scale =30)
        df_peak_table = table_extratcor(tables, chrom_headers)
        df_peak_table.columns = chrom_headers_shimadzu
        if (df_peak_table.empty):
            continue
        df_peak_table = df_peak_table.drop_duplicates(keep="first")
        try:
            inx_to_shift = df_peak_table[df_peak_table["Name"].str.contains(compound, flags = re.IGNORECASE)].index[0]
        except IndexError as ie:
            print("\"{}\" might not be present in the tables of the file {}.Please check this file".format(compound,worksheet_name))
            continue
        df_peak_table = shift_row_to_top(df_peak_table, inx_to_shift)
        cond_1 = df_peak_table["Name"] == ''
        cond_2 = df_peak_table["Name"] == np.nan
        cond_3 = df_peak_table["Name"].str.contains(compound, flags = re.IGNORECASE)
        inxs_to_remove = df_peak_table[cond_1 | cond_2 | ~cond_3].index
        df_peak_table = df_peak_table.drop(inxs_to_remove)

        # impurity calculation
        assay = calc_results(df_peak_table, compound, average_area, constant_1, constant_2, unit)
        df_peak_table['Assay%'] = assay
        try:
            df_peak_table['B.no'] = worksheet_name.split("_")[0]
            df_peak_table['Condition'] = worksheet_name.split("_")[1]
        except IndexError as ie:
            df_peak_table['B.no'] = worksheet_name
            df_peak_table['Condition'] = ''
        df_peak_table["Sample quantity"] = sample_qty
        df_peak_table["V1"] = sample_v1
        df_peak_table["V2"] = sample_v2
        df_peak_table["V3"] = sample_v3
        df_peak_table["Label claim"] = label_claim
        df_peak_table["Per Unit"] = unit
        df_peak_table = df_peak_table[['B.no', 'Condition', 'Sample quantity', 'V1', 'V2', 'V3', 'Label claim', 'Per Unit', 'Area', 'Assay%']]
        peak_master.append(df_peak_table)

        # writing to output sheet
    df_peak_master = pd.concat(peak_master)
    assay_template_sheet = assay_template.get_sheet(0)
    sample_input_list = [sample_qty, sample_v1, sample_v2, sample_v3,label_claim, unit]
    fill_rs_sheet(assay_template_sheet, df_area_table, df_peak_master, sample_input_list, input_list)
    # worksheets[0].name = worksheet_name
    # assay_template.active_sheet = 0

if __name__ == '__main__':
    special_compounds = ['Heparin', 'Bumetanide', 'Labetalol']
    software = int(input("Enter 1 for Shimadzu\nEnter 2 for Empower"))
    software = 'Shimadzu' if software == 1 else 'Empower'
    compound = input("Enter the compund name [As mentioned in the chromatogram] ")
    path = os.path.join(compound, 'Assay')
    rt_range =[]
    if(compound in special_compounds):
        if(compound == 'Heparin'):
            option = int(input("Enter 1 for Heparin assay\nEnter 2 for Methyl Paraben assay\nEnter 3 for Propyl Paraben assay\nEnter 4 for Benzyl Alcohol assay\nEnter 5 for  P-Hydroxybenzoic acid assay: "))
            if (option == 2):
                compound = 'Methyl Paraben'
                path = os.path.join("Heparin", "Methyl Paraben Assay")
                rt_range = [7,10]
            elif (option == 3):
                compound = 'Propyl Paraben'
                path = os.path.join("Heparin", "Propyl Paraben Assay")
                rt_range = [11,14]
            elif (option == 4):
                compound = 'Benzyl Alcohol'
                path = os.path.join("Heparin", "Benzyl Alcohol Assay")
            elif (option == 5):
                compound = 'P-Hydroxy benzoic acid'
                path = os.path.join("Heparin", "P-Hydroxy benzoic acid Assay")
                rt_range = [4,6.5]
        if(compound == 'Bumetanide'):
            option = int(input("Enter 1 for Bumetanide assay\nEnter 2 for Benzyl Alcohol assay: "))
            if (option == 2):
                compound = 'Benzyl Alcohol'
                path = os.path.join("Bumetanide", "Benzyl Alcohol Assay")
        if(compound == 'Labetalol'):
            option = int(input("Enter 1 for Labetalol Assay\nEnter 2 for Dextrose Assay: "))
            if(option == 2):
                compound = 'Dextrose'
                path = os.path.join("Labetalol", "Dextrose Assay")

    concentration = float(input("Enter the concentration "))
    year = str(datetime.today().year)
    # input data sources
    df_sample_prep = pd.read_excel(os.path.join(os.getcwd(), 'data', 'Templates', 'Assay-sample-preparation.xlsx'))
    input_list = [0]*16

    """ Methyl Paraben Heparin"""
    # input_list = [30.14,20,2,20,1,1,1,1,1,1,100,'12.01.22', 'AMD/21/HEP/PS-00', 'K52228657', '28/06/2022', compound]
    # input_list = [30.49,20,5,50,1,1,1,1,1,1,100,'01.03.22', '1060-010', 'K52228657', '31/08/2022', compound]

    """ Propyl Paraben Heparin"""
    # input_list = [15.43,50,1,20,1,1,1,1,1,1,100,'12.01.22', 'AMD/21/HEP/PS-00', '331171102', '28/06/2022', compound]
    # input_list = [15.38,50,2.5,50,1,1,1,1,1,1,100,'01.03.22', '1060-010', '331171102', '28/06/2022', compound]

    """P-Hydroxybenzoic Acid Heparin """
    # input_list = [12.66,50,5,50,1,1,1,1,1,1,99.8,'01.03.22', '1060-010', 'STBK0362', 'Dec-2024', compound]

    """" Milrinone"""
    # input_list = [20.50,100,2,20,1,1,1,1,1,1,100,'17.01.22', 'STP-T-0009-00', '20200805', '09-08-2024', compound]

    """Benzyl Alcohol Bumetanide"""
    # input_list = [528.61,50,2,20,1,1,1,1,1,1,99.8,'10.11.21','AMD/20/BUM/PS-00', '311076928AT', 'DEC-2021', compound]
    """Acetaminophen"""
    input_list = [20.02,20,5,50,1,1,1,1,1,1,99.6,'17.02.22','test-method','APAP/20091561', 'Aug-2025', compound]

    """Hydralazine HCl """
    # input_list = [20.48,20,1,50,1,1,1,1,1,1,99.94,'23.02.22','STP-T-0017-00', 'HHC/WS/001/21', 'NOV-22',compound]

    """ Labetalol"""
    # input_list = [45.64,20,1,1,1,1,1,1,1,1,98.0,'15.02.22','STP-T-0019-00', '2021072333', 'Dec-2023', compound]
    # input_list[0] = float(input("Enter the Weight taken "))
    # input_list[1] = float(input("Enter the standard preparation v1 "))
    # input_list[2] = float(input("Enter the standard preparation v2 "))
    # input_list[3] = float(input("Enter the standard preparation v3 "))
    # input_list[4] = float(input("Enter the standard preparation v4 "))
    # input_list[5] = float(input("Enter the standard preparation v5 "))
    # input_list[6] = float(input("Enter the standard preparation v6 "))
    # input_list[7] = float(input("Enter the standard preparation v7 "))
    # input_list[8] = float(input("Enter the standard preparation factor 1 "))
    # input_list[9] = float(input("Enter the standard preparation factor 2 "))
    # input_list[10] = float(input("Enter the standard preparation Potency "))
    # input_list[11] = input("Enter the date of analysis (dd/mm/yyyy) ")
    # input_list[12] = input("Enter the method of reference ")
    # input_list[13] = input("Enter WSID number ")
    # input_list[14] = input("Enter the use before date (dd/mm/yyyy) ")
    # input_list[15] = compound
    data_path = 'data' if software == 'Shimadzu' else os.path.join('data','empower-data')
    area_input = os.path.join(os.getcwd(), data_path, year, path, input_list[11], "{}-areas.pdf".format(compound))
    chrom_inputs = glob.glob(os.path.join(os.getcwd(), data_path, year,path, input_list[11],  '*.pdf'))
    try:
        chrom_inputs.remove(area_input)
    except ValueError as ve:
        print("Check the name of the RSD file. Make sure it is in the format: <compound name>-areas")
        exit()
    initiate_report_creation(compound, concentration, df_sample_prep, chrom_inputs, area_input, input_list, rt_range, software)
    assay_template.save(os.path.join(os.getcwd(), data_path, year, path, input_list[11], '{}-Assay.xls'.format(compound)))

    print("Reports saved successfully, check Output folder.")
