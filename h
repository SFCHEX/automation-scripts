def apply_custom_classification_rules(row):
    if pd.notna(row["Comment"]):
        for keyword, rule in custom_rules.items():
            if keyword.lower() in row["Comment"].lower() and row["Reason Category"] != rule["Reason Category"]:
                row["Reason Category"] = rule["Reason Category"]
                row["Reason Sub-Category"] = rule["Reason Sub-Category"]
                print(f"Changing Reason Category to '{rule['Reason Category']}' and Reason Sub-Category to '{rule['Reason Sub-Category']}' ")
    return row
