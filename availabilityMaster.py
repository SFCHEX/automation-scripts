import pandas as pd
import openpyxl
import argparse 
from datetime import datetime , timedelta
from openpyxl import load_workbook
import os

avaSheetMasterName="C:\\Users\\swx1283483\\Desktop\\tes\\Sep-Network Availability Dashboard-2023-18 (F).xlsx"
excludedSitesSheet="C:\\Users\\swx1283483\\Desktop\\tes\\Excluded List From Cells Breakdown.xlsx"


outputName="excludedAvailibilityOutsideCairo.xlsx"




def get_files_in_current_directory():
    try:
        current_directory = os.getcwd()
        files = os.listdir(current_directory)
        return [x for x in files if x.endswith(".xlsx") and outputName != x]
    except Exception as e:
        print(f"Error getting files in current directory: {str(e)}")
        return []

def merge_files(input_files):
    merged_data = pd.DataFrame()
    for input_file in input_files:
        df = pd.read_excel(input_file,"OC Cells AVA (All Tech)")
        merged_data =pd.concat([merged_data,df])
    return merged_data


def excludeAVA(avaSheet ,outputName):
    headers=["DAYID","CELL","REG","SITE",'TECH',"VENDOR","DOWN_TIME","AVA (%)"]

    avaSheet=avaSheet[avaSheet["TECH"]!="4G"]
    avaSheet=avaSheet[headers]
    excludedSheet= pd.read_excel(excludedSitesSheet, sheet_name="Sheet1")
    ZTEUpdateSheet= pd.read_excel(excludedSitesSheet, sheet_name="ZTE UPDATE")

   
    excluded=excludedSheet["SITE"].tolist()
    ZTEUpdate=ZTEUpdateSheet["Site ID"].tolist()
    
    avaSheet=avaSheet[~avaSheet["SITE"].isin(excluded)]

    avaSheetZTE=avaSheet[avaSheet["VENDOR"]=="Z"]
    avaSheetREST=avaSheet[avaSheet["VENDOR"]!="Z"]
    avaSheetZTE=avaSheetZTE[avaSheet["SITE"].isin(ZTEUpdate)]
    avaSheet=pd.concat([avaSheetZTE,avaSheetREST])


    avaSheet.to_excel(outputName, sheet_name='OC Cells AVA (All Tech)', index=False)
 
def updateNetworkAva(avaSheetName):
    try:
        avaWBMaster=openpyxl.load_workbook(avaSheetMasterName)
        sheetnum=2
        avaSheetMaster=avaWBMaster[f'OC Cells AVA (All Tech)-{sheetnum}']
        avaWBSheet=openpyxl.load_workbook(avaSheetName)
        avaSheet=avaWBSheet['Sheet1']
        print("opened AVA master sheet")
        for row in avaSheet.iter_rows(min_row=2, values_only=True):
                avaSheetMaster.append(row)
        avaWBMaster.save(avaSheetMasterName)
        print("Added values to master sheet")
    except Exception as e:
        print(e)

def main():
    parser = argparse.ArgumentParser(description="DataFornetAv")
    parser.add_argument("-a", "--avaSheetName", help="availability Update")
    args =parser.parse_args()

    if args.avaSheetName:
        excludeAVA(args.avaSheetName,"CellBreakdownExcluded.xlsx")
        print("outputed CellBreakdownExcluded.xlsx")
    else:
        input_files=get_files_in_current_directory()
        avaSheet=merge_files(input_files)
        excludeAVA(avaSheet ,outputName)
        print("outputed CellBreakdownExcluded.xlsx")

 



if __name__ == "__main__":

    main()
