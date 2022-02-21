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
ivi_template_input = xlrd.open_workbook(os.path.join(os.getcwd(), "data", "Templates",'Imp-vs-imp-template.xls'), formatting_info=True)
ivi_template = xlutils.copy.copy(ivi_template_input)

# Table headers
# chrom_headers = ['Peak#','Name','Ret. Time','Area','Area%','RRT']
# area_headers = ['Title', 'Ret. Time', 'Area', 'Area%', 'NTP', 'Tailing Factor']
chrom_headers = ['Name','Ret. Time','Area']
area_headers = ['Name','Area']


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

def calc_results (df_peak, df_rrf, sample_input_list, inputs, compound, df_area_table, base_rt, unit):
    rrt_master = []
    impurity_master = []
    rrf_master = []
    ignore_compounds = []
    sample_wt, sample_v1, sample_v2, sample_v3, sample_v4, sample_v5, sample_v6, sample_v7, label_claim, unit = tuple(sample_input_list)

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
        try:
            average_area = round(df_area_table['Area'][df_area_table['Name'].str.contains(name, re.IGNORECASE)].mean())
        except:
            average_area = round(df_area_table['Area'][df_area_table['Name'].str.contains(compound, re.IGNORECASE)].mean())

        rt = float(row[1])
        rrf_cond_1 = df_rrf['Compound'].str.contains(compound, flags = re.IGNORECASE)
        rrf_cond_2 = df_rrf['Impurity/Active Name'].str.contains(name, flags = re.IGNORECASE)
        rrf = df_rrf['RRF'][rrf_cond_1 & rrf_cond_2].values.tolist()
        if(not(rrf)):
            impurity_master.append(0)
            rrt_res = round(rt/base_rt, ndigits=2)
            rrt_master.append(rrt_res)
            rrf_master.append(0)
            ignore_compounds.append(name)
            continue
        rrf = float(rrf[0])
        input_list =  inputs[name] if name.lower() != 'unknown' else inputs[compound]
        constant_1 = (input_list[0]/input_list[1]) * (input_list[2]/input_list[3]) * (input_list[4]/input_list[5])
        constant_2 = (sample_v1/sample_wt) * (sample_v3/sample_v2) * (sample_v5/sample_v4) * (sample_v7/sample_v6) * (input_list[6]/label_claim)
        impurity = round((area/average_area) * constant_1 * constant_2 * (unit/rrf), ndigits=2)
        rrt_res = round(rt/base_rt, ndigits=2)
        impurity_master.append(impurity)
        rrt_master.append(rrt_res)
        rrf_master.append(rrf)

    return impurity_master, rrt_master, rrf_master, ignore_compounds

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


def fill_rs_sheet(output_sheet, df_area_table, df_peak_table, sample_input_list, inputs):
    table_col = 2
    keys = ['Related Compound-A','Related Compound-B','Related Compound-C','Related Compound-D','Ketorolac Tromethamine']
    #poject name
    setOutCell(output_sheet, 2, 3, 'Ketorolac Tromethamine')
    #Date
    setOutCell(output_sheet, 2, 4, inputs['Details'][0])
    #Method
    setOutCell(output_sheet, 2, 5, inputs['Details'][1])
    for key in keys:
        df_comp_areas = df_area_table[df_area_table['Name'].str.contains(key, flags= re.IGNORECASE)]
        average_area = round(df_comp_areas['Area'].mean())
        SD = round(df_comp_areas['Area'].std())
        RSD = round((SD/average_area)*100, ndigits=2)

        area_input = list(df_comp_areas['Area'])
        input_list = inputs[key]
        # WS ID No.
        setOutCell(output_sheet, table_col, 9, input_list[7])
        # std_wt
        setOutCell(output_sheet, table_col, 10, input_list[0])
        #  v1
        setOutCell(output_sheet, table_col, 11, input_list[1])
        # v2
        setOutCell(output_sheet, table_col, 12, input_list[2])
        #  v3
        setOutCell(output_sheet, table_col, 13, input_list[3])
        #  v4
        setOutCell(output_sheet, table_col, 14,  input_list[4])
        # v5
        setOutCell(output_sheet, table_col, 15, input_list[5])
        # potency
        setOutCell(output_sheet, table_col, 16, input_list[6])
        # areas
        setOutCell(output_sheet, table_col, 18, area_input[0])
        setOutCell(output_sheet, table_col, 19, area_input[1])
        setOutCell(output_sheet, table_col, 20, area_input[2])
        setOutCell(output_sheet, table_col, 21, area_input[3])
        setOutCell(output_sheet, table_col, 22, area_input[4])
        setOutCell(output_sheet, table_col, 23, area_input[5])
        setOutCell(output_sheet, table_col, 24, average_area)
        setOutCell(output_sheet, table_col, 25, SD)
        setOutCell(output_sheet, table_col, 26, RSD)

        table_col +=1

    # AR NO
    setOutCell(output_sheet, 1, 30, '')
    # Batch NO
    setOutCell(output_sheet, 3, 30, '')
    # Condition
    setOutCell(output_sheet, 4, 30, '')
    # Label Claim
    setOutCell(output_sheet, 5, 30, sample_input_list[8])
    # per unit
    setOutCell(output_sheet, 7, 30, sample_input_list[9])
    # sample_wt
    setOutCell(output_sheet, 2, 31, sample_input_list[0])
    #  v1
    setOutCell(output_sheet, 2, 32, sample_input_list[1])
    # v2
    setOutCell(output_sheet, 4, 31, sample_input_list[2])
    #  v3
    setOutCell(output_sheet, 4, 32, sample_input_list[3])
    #  v4
    setOutCell(output_sheet, 6, 31, sample_input_list[4])
    # v5
    setOutCell(output_sheet, 6, 32, sample_input_list[5])
    # v6
    setOutCell(output_sheet, 8, 31, sample_input_list[6])
    # v7
    setOutCell(output_sheet, 8, 32, sample_input_list[7])

#   Impurity table
    table_row = 37
    for index, row in df_peak_table.iterrows():
        if(table_row > 76):
            break
        setOutCell(output_sheet, 1, table_row, row[0])
        setOutCell(output_sheet, 2, table_row, row[1])
        setOutCell(output_sheet, 3, table_row, row[2])
        setOutCell(output_sheet, 4, table_row, row[3])
        setOutCell(output_sheet, 5, table_row, row[4])
        setOutCell(output_sheet, 6, table_row, row[5])
        table_row +=1

    sum_of_impurities = str(round(df_peak_table["% w/w"].sum(), ndigits=2))
    setOutCell(output_sheet, 6, 77, sum_of_impurities)

def initiate_report_creation(compound, df_rrf, df_sample_prep, chrom_inputs, area_inputs, inputs):
    sample_wt = df_sample_prep['Sample Volume'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v1 = df_sample_prep['v1'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v2 = df_sample_prep['v2'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v3 = df_sample_prep['v3'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v4 = df_sample_prep['v4'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v5 = df_sample_prep['v5'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v6 = df_sample_prep['v6'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_v7 = df_sample_prep['v7'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    label_claim = df_sample_prep['label claim'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    unit = df_sample_prep['per unit'][df_sample_prep["Compound"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0]
    sample_input_list = [sample_wt, sample_v1, sample_v2, sample_v3, sample_v4, sample_v5, sample_v6, sample_v7, label_claim, unit]
    # area tables extraction
    area_tables =  []
    for area_input in area_inputs:
        tables = camelot.read_pdf(area_input, pages= 'all', line_scale = 30)
        df_area = table_extratcor(tables, area_headers)
        df_area = df_area[['Name','Area']]
        area_tables.append(df_area)
    df_area_table = pd.concat(area_tables)
    df_area_table = df_area_table.drop(df_area_table[df_area_table['Name'] == ''].index)
    df_area_table["Area"] = df_area_table["Area"].astype(float)
    batch_size = len(chrom_inputs)
    outputs = []
    worksheets = ivi_template._Workbook__worksheets
    for index, chrom_input in enumerate(chrom_inputs):
        worksheet_name = chrom_input.split("\\")[-1].strip(".pdf")
        print(worksheet_name)
        # peak tables extratcion
        tables = camelot.read_pdf(chrom_input, pages= 'all', line_scale =30)
        df_peak_table = table_extratcor(tables, chrom_headers)
        if (df_peak_table.empty):
            continue
        df_peak_table = df_peak_table.drop_duplicates(keep="first")
        try:
            base_rt = float(df_peak_table['Ret. Time'][df_peak_table['Name'].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0])
        except IndexError as ie:
            print("\"{}\" might not be present in the tables of the file {}.Please check this file".format(compound,worksheet_name))
            continue
        cond_1 = df_peak_table["Name"] == ''
        cond_2 = df_peak_table["Name"] == np.nan
        cond_3 = df_peak_table["Name"].str.contains(compound, flags = re.IGNORECASE)

        inxs_to_remove = df_peak_table[cond_1 | cond_2 | cond_3].index
        df_peak_table = df_peak_table.drop(inxs_to_remove)

        # impurity calculation
        impurities, rrts, rrfs, ignore_compounds = calc_results(df_peak_table, df_rrf, sample_input_list, inputs, compound, df_area_table, base_rt, unit)
        for ic in ignore_compounds:
            inx_to_remove = df_peak_table[df_peak_table['Name'] == ic].index
            df_peak_table = df_peak_table.drop(inx_to_remove)

        df_peak_table['RRT'] = rrts
        df_peak_table['RRF'] = rrfs
        df_peak_table["% w/w"] = impurities
        df_peak_table = df_peak_table[['Name', 'Ret. Time','RRT', 'RRF', 'Area', '% w/w']]

        # writing to output sheet
        ivi_template_sheet = ivi_template.get_sheet(index)
        fill_rs_sheet(ivi_template_sheet, df_area_table, df_peak_table, sample_input_list, inputs)
        worksheets[index].name = worksheet_name

    ivi_template._Workbook__worksheets = [worksheet for worksheet in ivi_template._Workbook__worksheets if "Sheet" not in worksheet.name ]
    ivi_template.active_sheet = 0

if __name__ == '__main__':
    # Bumetanide
    # Acyclovir
    # Famotidine
    # ketorolacTromethamine
    # LabetalolHCl
    # compound = 'Acyclovir'
    # input_list = [50.43,100,5,50,5,50,1,1,1,1,94.4]
    compound = 'Ketorolac Tromethamine'
    year = str(datetime.today().year)
    # input data sources
    df_rrf = pd.read_excel(os.path.join(os.getcwd(), 'data', 'Templates', 'RRF-template.xlsx'))
    df_sample_prep = pd.read_excel(os.path.join(os.getcwd(), 'data', 'Templates', 'RS-sample-preparation.xlsx'))
    inputs = {
    'Ketorolac Tromethamine': [10.83,100,1,20,2,50,100, 'test-number'],
    'Related Compound-A': [1.019,10,1,20,2,50,92.73, 'test-number'],
    'Related Compound-B': [1.0955,10,1,20,2,50,99.31, 'test-number'],
    'Related Compound-C': [1.0957,10,1,20,2,50,99.54, 'test-number'],
    'Related Compound-D': [1.1931,10,1,20,2,50,98.14, 'test-number'],
    }
    # for key in inputs:
    #     input_list[key] = [0]*8
    # input_list[key][0] = float(input("Enter the Weight taken for {} ".format(key)))
    # input_list[key][1] = float(input("Enter the standard preparation v1 for {} ".fromat(key)))
    # input_list[key][2] = float(input("Enter the standard preparation v2 for {} ".fromat(key)))
    # input_list[key][3] = float(input("Enter the standard preparation v3 for {} ".fromat(key)))
    # input_list[key][4] = float(input("Enter the standard preparation v4 for {} ".fromat(key)))
    # input_list[key][5] = float(input("Enter the standard preparation v5 for {} ".fromat(key)))
    # input_list[key][6] = float(input("Enter the standard preparation Potency for {} ".fromat(key)))
    # input_list[key][7] = input("Enter the WSID number for {} ".fromat(key))
    inputs['Details'] = [0]*2
    inputs['Details'][0] = input("Enter the date of analysis dd.mm.yyyy ")
    inputs['Details'][1] = input("Enter the method of reference ")
    chrom_inputs = glob.glob(os.path.join(os.getcwd(), "data", year, compound, "RS", inputs['Details'][0], '*.pdf'))
    area_inputs = glob.glob(os.path.join(os.getcwd(),"data", year, compound, "RS", inputs['Details'][0], '*standard*.pdf'))
    chrom_inputs = [chrom_input for chrom_input in chrom_inputs if chrom_input not in area_inputs]
    initiate_report_creation(compound, df_rrf, df_sample_prep, chrom_inputs, area_inputs, inputs)
    ivi_template.save(os.path.join(os.getcwd(), "data", year, compound, "RS",  inputs['Details'][0], '{}-RS.xls'.format(compound)))
    print("Reports saved successfully, check Output folder.")
