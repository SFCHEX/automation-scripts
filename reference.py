

import pandas as pd
import argparse
import re
from datetime import datetime, timedelta
from fuzzywuzzy import fuzz


custom_rules = {
    "BSS": {
        "Reason Category": "BSS",
        "Reason Sub-Category": "BSS_HW"
    },

    # Add more rules here...
}

desired_header_order = ["Date", "Region", "BSC", "Site Name", "Site ID", "2G/3G", "ID","Site Category","Priority","Technical Area","Site Layer - Qism","Type Of Sharing","Host/Guest","Alarm Occurance Time","Fault Occurance Time","Fault Clearance Time","MTTR","Hybrid Down Time","SLA Status","Site Type", "Reason Category","Reason Sub-Category","Comment", "Owner","Access Type","Cascaded To", "Final", "Office","Corp","Generator","Vendor"]



desired_header_order_mt= ["Site", "TECH", "ID", "Controller", "RBSType", "Region", "nm_tier","EventTime","CeaseTime","MTTR","Duration","Reason Category","Reason Sub-Category","RootCause","DateOfOutage","Vendor","FilledBy","MR_CP","Office","Site Type"]


original_headers= {"Reason Category":["BSS", "High_temp", "Others", "Power", "TX",""],

                       "Reason Sub-Category":['BSS_HW','BSS_SW','High_Temp','High_Temp_Dependency','Others_Illegal_Intervention','Others_ROT','Others_Dependency_Illegal Intervention','Others_Unknown Reason','Others_Dependency_Unknown Reason','Others_Wrong_Action','Power_Commercial','Power_Dependency_Commercial','Power_Generator','Power_Dependency_Generator','Power_HW (Cct Breakers, Cables)','Power_Dependency_HW (Cct Breakers, Cables)','Power_Power Criteria','Power_Dependency_Power Criteria','Power_Solar Cell','Power_Dependency_Solar Cell','TX_Bad Performance','TX_Dependency_Bad Performance','TX_HW Failure','TX_Dependency_HW Failure','TX_LOS','TX_Dependency_LOS','TX_Physical Connection','TX_Dependency_Physical Connection','TX_Power Supply','TX_Dependency_Power Supply','TX_SW','TX_Telecom Egypt','TX_Dependency_Telecom Egypt','BSS/License',""],

                       "Owner":["BSS","FM","GD","ROT","RT","TD","TX","E///","ZTE",""],
                       }


def merge_categories(original_categories, new_categories):
    # Initialize a dictionary to store mappings
    mappings = {}
    
    # Remove empty strings and non-string objects from input lists
    original_categories = [x for x in original_categories if isinstance(x, str) and x]
    new_categories = [x for x in new_categories if isinstance(x, str) and x]

    # Iterate through each new category
    for new_cat in new_categories:
        best_similarity = 0
        best_orig_cat = None
        
        # Iterate through each original category
        for orig_cat in original_categories:
            # Calculate the similarity score between the two categories
            similarity_score = fuzz.ratio(new_cat.lower(), orig_cat.lower())

            # If the similarity score is above a threshold (e.g., 80) and higher than the current best,
            # consider them a match
            if similarity_score >= 70 and similarity_score > best_similarity:
                best_similarity = similarity_score
                best_orig_cat = orig_cat
        
        # If a match is found, add it to the mapping
        if best_orig_cat:
            mappings[new_cat] = best_orig_cat
    mappings={key: value for key,value in mappings.items() if key!=value and key!=""}
    return mappings

def apply_cascaded_to_rule(row):
    comment = row["Comment"]
    cas_matches = re.findall(r'cas', comment, re.IGNORECASE)

    if cas_matches and row["Reason Category"]!="Others":
        match = re.search(r'\d{4}', comment)
        if match:
            site_id_prefix = row["Site ID"][:3]
            site_id_postfix = row["Site ID"][-4:]
            cascaded_to_digits = match.group()

            if site_id_postfix == cascaded_to_digits:
                row["Cascaded To"] = ""
            else:
                row["Cascaded To"] = site_id_prefix + cascaded_to_digits

            # Add a print statement here to describe the modification
#            print(f"Modified Cascaded To in row {row.name}: '{row['Cascaded To']}'")

    return row


def apply_custom_rules(row, custom_rules):
    comment = row["Comment"]
    for keyword, rule in custom_rules.items():
        if keyword.lower() in comment.lower():
            if rule["Reason Category"]!="":
                row["Reason Category"] = rule["Reason Category"]

            if rule["Reason Sub-Category"]!="":
                row["Reason Sub-Category"] = rule["Reason Sub-Category"]
            # Add a print statement here to describe the modification
#            print(f"Modified Reason Category in row {row.name}: '{row['Reason Category']}'")
#            print(f"Modified Reason Sub-Category in row {row.name}: '{row['Reason Sub-Category']}'")
            break  # Once a keyword is found, we apply the rule and break out of the loop
    return row

def apply_tcr_logic(row):
    comment = row["Comment"]
    tcr_match = re.search(r'tcr\s*\d{6,7}', comment, re.IGNORECASE)

    if pd.notna(row["SLA Status"]):  # Check if "SLA Status" cell is not empty
        if tcr_match:
            sla_status = row["SLA Status"]
            if isinstance(sla_status, str):
                sla_status_words = sla_status.split()
                if sla_status_words and sla_status_words[0].lower() != "planned":
                    row["SLA Status"] = "Planned " + " ".join(sla_status_words[1:])
                    # Add a print statement here to describe the modification
#                    print(f"Modified SLA Status in row {row.name}: '{row['SLA Status']}'")
        else:
            sla_status = row["SLA Status"]
            if isinstance(sla_status, str):
                sla_status_words = sla_status.split()
                if sla_status_words and sla_status_words[0].lower() != "unplanned":
                    row["SLA Status"] = "Unplanned " + " ".join(sla_status_words[1:])
                    # Add a print statement here to describe the modification
#                    print(f"Modified SLA Status in row {row.name}: '{row['SLA Status']}'")

    return row



def fixes(data, custom_rules):
    data = data.apply(apply_cascaded_to_rule, axis=1)
    data = data.apply(lambda row: apply_custom_rules(row, custom_rules), axis=1)
    data = data.apply(apply_tcr_logic, axis=1)
    return data
def process_excel_files(input_files,input_files2 ,output_file):
    combined_data = pd.DataFrame(columns=desired_header_order)  
    for input_file in input_files:
        data = pd.read_excel(input_file)
        for header in desired_header_order:
            if header not in data.columns:
                data[header]=None

        data["Site ID"] = data["SiteCode"]
        data["2G/3G"] = data["Tech"]
        data["Site Layer - Qism"] = data["Site Layer Qism"]
        data["Hybrid Down Time"] = data["Down Time"]
         

        data = data[desired_header_order]


        combined_data = pd.concat([combined_data, data])








    combined_data_mt = pd.DataFrame(columns=desired_header_order_mt)  # Initialize with desired order
    for input_file2 in input_files2:
        data2 = pd.read_excel(input_file2)
        for header in desired_header_order_mt:
            if header not in data.columns:
                data[header]=None

        data2["Site Type"] = data2["RBSType"] + " " + data2["ID/OD"]
        data2["ID"] = data2["Site"].str[-4:]
        data2["TECH"] = data2["Site"].apply(lambda x: "4G" if x[0] == "L" else "3G" if x[0] == "U" else "2G")
        data2["Site"] = data2["Site"].str[-7:]
        data2["Reason Category"] = data2["Category"]
        data2["Reason Sub-Category"] = data2["SubCategory"]

        for header in desired_header_order:
            if header not in data2.columns:
                data2[header]=None
        
         

        data2["Date"] = (datetime.now()- timedelta(days=1))
        data2["BSC"]=data2["Controller"]
        data2["2G/3G"]=data2["TECH"]

        data2["Priority"]=data2["nm_tier"]
        data2["Site ID"]=data2["Site"]
        data2["Site Name"]=data2["Site"]

        data2["Technical Area"]=data2["Region"]
        data2["Site Layer - Qism"]=data2["Region"]

        data2["Fault Clearance Time"]=data2["CeaseTime"]
        data2["Alarm Occurance Time"]=data2["EventTime"]
        data2["Fault Occurance Time"]=data2["EventTime"]
        data2["Hybrid Down Time"]=data2["Duration"]
        data2["Comment"]=data2["RootCause"]

        data2 = data2[desired_header_order]

        combined_data_mt = pd.concat([combined_data_mt, data2])




    combined_data = pd.concat([combined_data, combined_data_mt])
    combined_data["Date"]=combined_data["Date"].apply(lambda x: x.strftime("%m-%d-%Y"))
    combined_data=combined_data[desired_header_order]

    for key,header in original_headers.items():
        mcd=merge_categories(header,combined_data[key].unique().tolist())
        print(combined_data[key].unique().tolist())
        print(f"\n{key}") 
        for og,nw in mcd.items():
            print(f"{og}:{nw}")
            combined_data[key] = combined_data[key].str.replace(og,nw)

    combined_data=fixes(combined_data,custom_rules)


    for key,header in original_headers.items():
        mcd=merge_categories(header,combined_data[key].unique().tolist())
        print(combined_data[key].unique().tolist())
        print(f"\n{key}") 
        for og,nw in mcd.items():
            print(f"{og}:{nw}")
            combined_data[key] = combined_data[key].str.replace(og,nw)


    combined_data.to_excel(output_file, index=False)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process multiple Excel files.")
    parser.add_argument("-i","--input1", nargs="+", help="List of input Excel files")
    parser.add_argument("-m","--input2", nargs="+", help="List of input Excel files")
    parser.add_argument("-o","--output", help="Output Excel file")
    args = parser.parse_args()

    process_excel_files(args.input1,args.input2 ,args.output)
