
def updateSiteData(date_str, pivot_table_ava, pivot_table_pad, pivot_table_par, xlsb_file_path):
    with pyxlsb.open_workbook(xlsb_file_path) as wb:
        with pyxlsb.get_sheet(wb, sheet_name='AVA') as sheet_ava:
            data_ava = []
            for row in sheet_ava.rows():
                data_ava.append([item.v for item in row])

        with pyxlsb.get_sheet(wb, sheet_name='PAD') as sheet_pad:
            data_pad = []
            for row in sheet_pad.rows():
                data_pad.append([item.v for item in row])

        with pyxlsb.get_sheet(wb, sheet_name='PAR') as sheet_par:
            data_par = []
            for row in sheet_par.rows():
                data_par.append([item.v for item in row])

    section_data_ava = pd.DataFrame(data_ava[2:], columns=data_ava[1])
    section_data_pad = pd.DataFrame(data_pad[2:], columns=data_pad[1])
    section_data_par = pd.DataFrame(data_par[2:], columns=data_par[1])

    # Ensure that the 'SITE ID' values in the pivot tables match the XLSB file
    pivot_table_ava = pivot_table_ava[pivot_table_ava['SITE ID'].isin(section_data_ava['SITE ID'])]
    pivot_table_pad = pivot_table_pad[pivot_table_pad['SITE ID'].isin(section_data_pad['SITE ID'])]
    pivot_table_par = pivot_table_par[pivot_table_par['SITE ID'].isin(section_data_par['SITE ID'])]

    # Update the Data column for the specified date in each section
    section_data_ava[date_str] = pivot_table_ava.set_index('SITE ID')['Data']
    section_data_pad[date_str] = pivot_table_pad.set_index('SITE ID')['Data']
    section_data_par[date_str] = pivot_table_par.set_index('SITE ID')['Data']

    # Save the updated data back to the XLSB file
    with pyxlsb.open_workbook(xlsb_file_path, 'w') as wb:
        with wb.get_sheet(sheet_name='AVA') as sheet_ava:
            sheet_ava.writer.writerows([section_data_ava.columns] + section_data_ava.values.tolist())

        with wb.get_sheet(sheet_name='PAD') as sheet_pad:
            sheet_pad.writer.writerows([section_data_pad.columns] + section_data_pad.values.tolist())

        with wb.get_sheet(sheet_name='PAR') as sheet_par:
            sheet_par.writer.writerows([section_data_par.columns] + section_data_par.values.tolist())
