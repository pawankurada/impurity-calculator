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
rs_template_input = xlrd.open_workbook(os.path.join(os.getcwd(), "data", "Templates",'RS-template-batch.xls'), formatting_info=True)
rs_template = xlutils.copy.copy(rs_template_input)

# Table headers
# chrom_headers = ['Peak#','Name','Ret. Time','Area','Area%','RRT']
# area_headers = ['Title', 'Ret. Time', 'Area', 'Area%', 'NTP', 'Tailing Factor']
chrom_headers = ['Name','Ret. Time','Area']
area_headers = ['Title','Area']


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

def calc_results (df_peak, df_rrf, compound, average_area, constant_1, constant_2, unit):
    base_rt = float(df_peak['Ret. Time'][df_peak['Name'].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0])
    rrt_master = []
    impurity_master = []
    rrf_master = []
    for index, row in df_peak.iterrows():
        name =row[0]
        if(name.lower() == compound.lower()):
            impurity_master.append(0)
            rrt_master.append(1)
            rrf_master.append(0)
            continue
        if(name == np.nan or name == ''):
            continue
#             impurity_master.append(0)
#             rrt_master.append(0)

        area = float(row[2])
        rrf_cond_1 = df_rrf['Compound'].str.contains(compound, flags = re.IGNORECASE)
        rrf_cond_2 = df_rrf['Impurity/Active Name'].str.contains(name, flags = re.IGNORECASE)
        rrf = df_rrf['RRF'][rrf_cond_1 & rrf_cond_2].values.tolist()
        rt = float(row[1])

        if(not(rrf)):
            print("ignoring {}".format(name))
            impurity_master.append(0)
            rrt_res = round(rt/base_rt, ndigits=2)
            rrt_master.append(rrt_res)
            rrf_master.append(0)
            continue

        rrf = float(rrf[0])
        impurity = round((area/average_area) * constant_1 * constant_2 * (unit/rrf), ndigits=2)
        rrt_res = round(rt/base_rt, ndigits=2)
        impurity_master.append(impurity)
        rrt_master.append(rrt_res)
        rrf_master.append(rrf)

    return impurity_master, rrt_master, rrf_master

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
    df_result_table = pd.concat(result_tables, ignore_index=True)
    return df_result_table


def fill_rs_sheet(output_sheet, df_area_table, df_peak_table, sample_input_list):
    average_area = float(df_area_table["Area"][df_area_table["Title"] == "Average"].values.tolist()[0])
    area_input = list(df_area_table['Area'])

    #poject name
    setOutCell(output_sheet, 2, 3, '')
    #Date
    setOutCell(output_sheet, 2, 4, '')
    #Method
    setOutCell(output_sheet, 2, 5, '')
    # WS ID No.
    setOutCell(output_sheet, 1, 9, '')
    # potency
    setOutCell(output_sheet, 3, 9, input_list[-1])
    # use before date
    setOutCell(output_sheet, 5, 9, '')
    # Average area
    setOutCell(output_sheet, 7, 9, average_area)
    # std_wt
    setOutCell(output_sheet, 2, 10, input_list[0])
    #  v1
    setOutCell(output_sheet, 2, 11, input_list[1])
    # v2
    setOutCell(output_sheet, 4, 10, input_list[2])
    #  v3
    setOutCell(output_sheet, 4, 11, input_list[3])
    #  v4
    setOutCell(output_sheet, 6, 10,  input_list[4])
    # v5
    setOutCell(output_sheet, 6, 11, input_list[5])
    # v6
    setOutCell(output_sheet, 8, 10, input_list[6])
    # v7
    setOutCell(output_sheet, 8, 11, input_list[7])
    # factor
    setOutCell(output_sheet, 9, 10, input_list[8])
    # factor
    setOutCell(output_sheet, 9, 11, input_list[9])
    # AR NO
    setOutCell(output_sheet, 1, 14, '')
    # Batch NO
    setOutCell(output_sheet, 3, 14, '')
    # Condition
    setOutCell(output_sheet, 4, 14, '')
    # Label Claim
    setOutCell(output_sheet, 5, 14, sample_input_list[8])
    # per unit
    setOutCell(output_sheet, 7, 14, sample_input_list[9])
    # sample_wt
    setOutCell(output_sheet, 2, 15, sample_input_list[0])
    #  v1
    setOutCell(output_sheet, 2, 16, sample_input_list[1])
    # v2
    setOutCell(output_sheet, 4, 15, sample_input_list[2])
    #  v3
    setOutCell(output_sheet, 4, 16, sample_input_list[3])
    #  v4
    setOutCell(output_sheet, 6, 15, sample_input_list[4])
    # v5
    setOutCell(output_sheet, 6, 16, sample_input_list[5])
    # v6
    setOutCell(output_sheet, 8, 15, sample_input_list[6])
    # v7
    setOutCell(output_sheet, 8, 16, sample_input_list[7])
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
    table_row = 20
    for index, row in df_peak_table.iterrows():
        if(table_row > 60):
            break
        setOutCell(output_sheet, 1, table_row, row[0])
        setOutCell(output_sheet, 2, table_row, row[1])
        setOutCell(output_sheet, 3, table_row, row[2])
        setOutCell(output_sheet, 4, table_row, row[3])
        setOutCell(output_sheet, 5, table_row, row[4])
        setOutCell(output_sheet, 6, table_row, row[5])
        table_row +=1

    sum_of_impurities = round(df_peak_table["% w/w"].sum(), ndigits=2)
    setOutCell(output_sheet, 6, 62, sum_of_impurities)

def initiate_report_creation(compound, df_rrf, df_sample_prep, chrom_inputs, area_input, input_list):
    sample_wt = df_sample_prep['vials'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v1 = df_sample_prep['v1'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v2 = df_sample_prep['v2'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v3 = df_sample_prep['v3'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v4 = df_sample_prep['v4'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v5 = df_sample_prep['v5'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v6 = df_sample_prep['v6'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v7 = df_sample_prep['v7'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    label_claim = df_sample_prep['label claim'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    unit = df_sample_prep['per unit'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]

    constant_1 = (input_list[0]/input_list[1]) * (input_list[2]/input_list[3]) * (input_list[4]/input_list[5])*(input_list[6]/input_list[7]) * (input_list[8]/input_list[9])
    constant_2 = (sample_v1/sample_wt) * (sample_v3/sample_v2) * (sample_v5/sample_v4) * (sample_v7/sample_v6) * (input_list[10]/label_claim)
    # area table extraction
    tables = camelot.read_pdf(area_input, pages= 'all',flavor='stream')
    df_area_table = table_extratcor(tables, area_headers)
    df_area_table = df_area_table[['Title','Area']]
    average_area = float(df_area_table["Area"][df_area_table["Title"] == "Average"].values.tolist()[0])


    batch_size = len(chrom_inputs)
    outputs = []
    worksheets = rs_template._Workbook__worksheets
    for index, chrom_input in enumerate(chrom_inputs):
        # print(chrom_input)
        # peak tables extratcion
        tables = camelot.read_pdf(chrom_input, pages= 'all',flavor='stream')
        df_peak_table = table_extratcor(tables, chrom_headers)
        df_peak_table = df_peak_table.drop_duplicates(keep="first")
        inx_to_shift = df_peak_table[df_peak_table["Name"].str.contains(compound, flags = re.IGNORECASE)].index[0]
        df_peak_table = shift_row_to_top(df_peak_table, inx_to_shift)
        cond_1 = df_peak_table["Name"] == ''
        cond_2 = df_peak_table["Name"] == np.nan
        inxs_to_remove = df_peak_table[cond_1 | cond_2].index
        df_peak_table = df_peak_table.drop(inxs_to_remove)

        # impurity calculation
        impurities, rrts, rrfs = calc_results(df_peak_table, df_rrf, compound, average_area, constant_1, constant_2, unit)
        df_peak_table['RRT'] = rrts
        df_peak_table['RRF'] = rrfs
        df_peak_table["% w/w"] = impurities
        df_peak_table = df_peak_table[['Name', 'Ret. Time','RRT', 'RRF', 'Area', '% w/w']]

        # writing to output sheet

        rs_template_sheet = rs_template.get_sheet(index)
        sample_input_list = [sample_wt, sample_v1, sample_v2, sample_v3, sample_v4, sample_v5, sample_v6, sample_v7, label_claim, unit]
        fill_rs_sheet(rs_template_sheet, df_area_table, df_peak_table, sample_input_list)
        worksheets[index].name = chrom_input.split("\\")[-1].strip(".pdf")

    rs_template._Workbook__worksheets = [worksheet for worksheet in rs_template._Workbook__worksheets if "Sheet" not in worksheet.name ]
    rs_template.active_sheet = 0

if __name__ == '__main__':
    # Bumetanide
    # Acyclovir
    # Famotidine
    # ketorolacTromethamine
    # LabetalolHCl
    # compound = 'Acyclovir'
    # input_list = [50.43,100,5,50,5,50,1,1,1,1,94.4]
    compound = input("Enter the compund name [As mentioned in the chromatogram] ")

    # input data sources
    df_rrf = pd.read_excel(os.path.join(os.getcwd(), 'data', 'Templates', 'RRF-template.xlsx'))
    df_sample_prep = pd.read_excel(os.path.join(os.getcwd(), 'data', 'Templates', 'Sample Preparation.xlsx'))
    area_input = os.path.join(os.getcwd(), "data", "RS", compound, "{}-areas.pdf".format(compound))
    chrom_inputs = glob.glob(os.path.join(os.getcwd(), "data", "RS", compound, '*.pdf'))
    chrom_inputs.remove(area_input)
    input_list = [0]*11
    input_list[0] = float(input("Enter the Weight taken "))
    input_list[1] = float(input("Enter the standard preparation v1 "))
    input_list[2] = float(input("Enter the standard preparation v2 "))
    input_list[3] = float(input("Enter the standard preparation v3 "))
    input_list[4] = float(input("Enter the standard preparation v4 "))
    input_list[5] = float(input("Enter the standard preparation v5 "))
    input_list[6] = float(input("Enter the standard preparation v6 "))
    input_list[7] = float(input("Enter the standard preparation v7 "))
    input_list[8] = float(input("Enter the standard preparation factor 1 "))
    input_list[9] = float(input("Enter the standard preparation factor 2 "))
    input_list[10] = float(input("Enter the standard preparation Potency "))
    initiate_report_creation(compound, df_rrf, df_sample_prep, chrom_inputs, area_input, input_list)
    rs_template.save(os.path.join(os.getcwd(), "data", 'output', '{}-RS.xls'.format(compound)))
    print("Reports saved successfully, check Output folder.")
