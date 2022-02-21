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
area_norm_template_input = xlrd.open_workbook(os.path.join(os.getcwd(), "data", "Templates",'area-norm-template.xls'), formatting_info=True)
area_norm_template = xlutils.copy.copy(area_norm_template_input)

# Table headers
chrom_headers = ['Name','Ret. Time','Area']


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

def calc_results (df_peak, compound, main_peak, factor, base_rt):
    rrt_master = []
    area_percent__master = []
    sum_of_areas = round(df_peak["Area"].sum(), ndigits=2)
    constant_1 = (factor * main_peak) + sum_of_areas
    sum =0
    for index, row in df_peak.iterrows():
        rt = float(row[1])
        area = row[2]
        area_res = (area/constant_1)*100
        sum+=area_res
        rrt_res = round(rt/base_rt, ndigits=2)
        rrt_master.append(rrt_res)
        area_percent__master.append(round(area_res, ndigits=2))

    return  rrt_master, area_percent__master, round(sum, ndigits=2)

def shift_row_to_top(df, index_to_shift):
    idx = df.index.tolist()
    idx.remove(index_to_shift)
    df = df.reindex([index_to_shift] + idx)
    return df

def check_transpose(df_table):
    df_table_t = df_table.T
    search = df_table_t.where(df_table_t==chrom_headers[1]).dropna(how='all').dropna(axis=1)
    inx = list(search.index)
    if(inx):
        inx= inx[0]
        new_header = df_table_t.iloc[inx]
        if('Ret. Time' and 'Area' in list(new_header)):
            if(inx == df_table_t.shape[0]-1):
                return df_table.T.reindex(index=df_table_t.index[::-1]).reset_index()
            else:
                return df_table.T
        else:
            return df_table
    else:
        return df_table

def table_extratcor(tables, headers):
    df_result_table =''
    result_tables = []
    for table in tables:
        df_table = table.df
        df_table = check_transpose(df_table)
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


def fill_rs_sheet(output_sheet, df_peak_table, user_input_list, ap_sum):
    # ARR no
    setOutCell(output_sheet, 1, 0, '')
    # B.No & Condition
    setOutCell(output_sheet, 3, 0, '')
    # main peak area
    setOutCell(output_sheet, 5, 4, user_input_list[0])
    # Base RT
    setOutCell(output_sheet, 3, 3, user_input_list[5])
    # dilution factor 1
    setOutCell(output_sheet, 9, 1, user_input_list[1])
    # dilution factor 1
    setOutCell(output_sheet, 9, 2, user_input_list[2])
    # Factor
    setOutCell(output_sheet, 12, 2, user_input_list[4])
#   RRT table
    table_row = 4
    for index, row in df_peak_table.iterrows():
        if(table_row > 102):
            break
        setOutCell(output_sheet, 0, table_row, row[0])
        setOutCell(output_sheet, 1, table_row, row[1])
        setOutCell(output_sheet, 2, table_row, row[2])
        setOutCell(output_sheet, 3, table_row, row[3])
        setOutCell(output_sheet, 4, table_row, row[4])
        table_row +=1

    sum_of_areas = round(df_peak_table["Area"].sum(), ndigits=2)
    sum_of_percentage = df_peak_table["Area%"].sum()
    sum_of_rrt = round(((user_input_list[4]*user_input_list[0]) / ((user_input_list[4] * user_input_list[0])+sum_of_areas))*100, ndigits=1)

    setOutCell(output_sheet, 1, table_row, "SUM")
    setOutCell(output_sheet, 2, table_row, sum_of_areas)
    setOutCell(output_sheet, 3, table_row, ap_sum)
    setOutCell(output_sheet, 4, table_row  , sum_of_rrt)

def initiate_report_creation(compound, chrom_inputs, input_list):
    dil_s01 = input_list[1]
    dil_s02 = input_list[2]
    factor = dil_s01/dil_s02
    batch_size = len(chrom_inputs)
    worksheets = area_norm_template._Workbook__worksheets
    for index, chrom_input in enumerate(chrom_inputs):
        worksheet_name = chrom_input.split("\\")[-1].strip(".pdf")
        input_list[0] = float(input("Enter main peak area for {} ".format(worksheet_name)))
        main_peak = input_list[0]
        # peak tables extratcion
        tables = camelot.read_pdf(chrom_input, pages= 'all', line_scale = 30)
        df_peak_table = table_extratcor(tables, chrom_headers)
        if(df_peak_table.empty):
            continue
        df_peak_table = df_peak_table.drop_duplicates(keep="first")
        try:
            base_rt = float(df_peak_table['Ret. Time'][df_peak_table["Name"].str.contains(compound, flags = re.IGNORECASE)].values.tolist()[0])
        except IndexError as ie:
            print("\"{}\" might not be present in the tables of the file {}.Please check this file".format(compound,worksheet_name))
            continue
        compound_row = df_peak_table[df_peak_table["Name"].str.contains(compound, flags = re.IGNORECASE)].index
        df_peak_table = df_peak_table.drop(compound_row)
        cond_1 = df_peak_table["Name"] == ''
        cond_2 = df_peak_table["Name"] == np.nan
        inxs_to_remove = df_peak_table[cond_1 | cond_2].index
        df_peak_table = df_peak_table.drop(inxs_to_remove)
        df_peak_table["Area"] = [float(area) for area in df_peak_table["Area"]]
        # df_peak_table["Area%"] = [float(area) for area in df_peak_table["Area%"]]

        # RRT calculation
        rrts,area_percents,ap_sum = calc_results(df_peak_table, compound, main_peak, factor, base_rt)
        df_peak_table['RRT'] = rrts
        df_peak_table["Area%"] = area_percents
        df_peak_table = df_peak_table[['Ret. Time','Name','Area', 'Area%', 'RRT']]

        # writing to output sheet
        area_norm_sheet = area_norm_template.get_sheet(index)
        user_input_list = input_list + [factor,base_rt]
        fill_rs_sheet(area_norm_sheet, df_peak_table, user_input_list, ap_sum)
        worksheets[index].name = worksheet_name

    area_norm_template._Workbook__worksheets = [worksheet for worksheet in area_norm_template._Workbook__worksheets if "Sheet" not in worksheet.name ]
    area_norm_template.active_sheet = 0

if __name__ == '__main__':
    # Bumetanide
    # Acyclovir
    # Famotidine
    # ketorolacTromethamine
    # LabetalolHCl
    # compound = 'Acyclovir'
    # input_list = [50.43,100,5,50,5,50,1,1,1,1,94.4]
    year = str(datetime.today().year)
    compound = input("Enter the compund name [As mentioned in the chromatogram] ")

    # input data sources
    input_list = [0]*4
    input_list[3] = input("Enter the date of analysis ")
    input_list[1] = float(input("Enter the dilution factor of sample 01 "))
    input_list[2] = float(input("Enter the dilution factor of sample 02 "))
    chrom_inputs = glob.glob(os.path.join(os.getcwd(), "data", year, compound, "RS", input_list[3], '*.pdf'))
    initiate_report_creation(compound,chrom_inputs, input_list)
    area_norm_template.save(os.path.join(os.getcwd(), "data", year, compound, "RS", input_list[3], '{}-RS-area-norm.xls'.format(compound)))
    print("Reports saved successfully, check Output folder.")
