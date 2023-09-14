import pandas as pd
import openpyxl
import argparse 
import datetime

avaSheetMasterName=""
powerSheetMasterName=""
siteDataMasterName=""

def loadSheets(avaSheetName,powerSheetName):
    avaSheet=pd.read_excel(avaSheetName,"")
    powerSheet=pd.read_excel(powerSheetName,"")
    return avaSheet, powerSheet

def updateNetworkAva(avaSheetName):
    try:
        avaWBMaster=openpyxl.load_workbook(avaSheetMasterName)
        avaSheetMaster=avaWBMaster['']
        avaWBSheet=openpyxl.load_workbook(avaSheetName)
        avaSheet=avaWBSheet['']
        for row in avaSheet.iter_rows(min_row=2, values_only=True):
                avaSheetMaster.append(row)
        avaWBMaster.save(avaSheetMasterName)
    except Exception as e:
        print(e)



def updatePowerAlarm(powerSheetName):
    try:
        powerWBMaster=openpyxl.load_workbook(powerSheetMasterName)
        powerSheetMaster=powerWBMaster['']
        powerWBSheet=openpyxl.load_workbook(powerSheetName)
        powerSheet=powerWBSheet['']
        for row in powerSheet.iter_rows(min_row=2, values_only=True):
                powerSheetMaster.append(row)
        powerWBMaster.save(powerSheetMasterName)
    except Exception as e:
        print(e)



def getAVA(avaSheet):
    avaSheet=avaSheet[avaSheet[]=="yes"]
    AVA = pd.pivot_table(avaSheet, values='Site ID', index='', columns='ava', aggfunc='avg')
    blank_columns = pd.DataFrame('', index=range(len(avaSheet)), columns=[''] * 3)  # Insert 3 blank columns
    avaSheet=pd.concat([avaSheet,blank_columns,avaSheet],axis=1)
      
    return AVA

def getPADPAR(powerSheet):
    powerSheet=powerSheet[powerSheet[] is in ["2g","3g"]]
    PAD = pd.pivot_table(powerSheet, values='Site ID', index='', columns='MTTR', aggfunc='sum')
    PAR = pd.pivot_table(powerSheet, values='Site ID', index='', columns='MTTR', aggfunc='count')
    blank_columns = pd.DataFrame('', index=range(len(powerSheet)), columns=[''] * 3)  # Insert 3 blank columns
    powerSheet=pd.concat([powerSheet,blank_columns,PAD,blank_columns,PAR],axis=1)
 
    return PAD,PAR


def select_and_replace_column(df, col_name, type_index, new_column_data):
    type_row = df.iloc[2]
    cols_with_name = [col for col in df.columns if col == col_name]
    selected_col = next((col for col in cols_with_name if type_row[df.columns.get_loc(col)] == type_index), None)
    
    if selected_col:
        df[selected_col] = df['Site ID'].map(new_column_data.set_index('Site ID')['Pivot Column'])
        print(f"Column '{selected_col}' replaced with new data and saved.")
    else:
        print(f"No column with type '{type_index}' found.")

def updateSiteData(PAR,AVA,PAD):
    date = (datetime.now() - datetime.timedelta(days=1)).strftime("%m-%d-%y")
    siteDataMaster= pd.read_excel(siteDataMasterName, sheet_name="sheet1")
    select_and_replace_column(siteDataMaster,date,"AVA",AVA)
    select_and_replace_column(siteDataMaster,date,"PAR",PAR)
    select_and_replace_column(siteDataMaster,date,"PAD",PAD)
    siteDataMaster.to_excel(siteDataMasterName, sheet_name="sheet1", index=False)
    

def main():
    parser = argparse.ArgumentParser(description="DataFornetAv")
    parser.add_argument("-a", "avaSheetName", nargs="+", help="availability Update")

    parser.add_argument("-p", "powerSheetName", nargs="+", help="power update")
    args =parser.parse_args()

    avaSheet,powerSheet=loadSheets(args.avaSheetName,args.powerSheetName)

    updateNetworkAva(args.avaSheetName)
    updatePowerAlarm(args.powerSheetName)

    AVA=getAVA(avaSheet)
    PAD,PAR=getPADPAR(powerSheet)

    updateSiteData(PAR,PAD,AVA)
 

if __name__ == "__main__":

    main()
