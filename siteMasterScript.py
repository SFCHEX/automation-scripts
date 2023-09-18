
def updateSiteData(date_str, pivot_table_ava, pivot_table_pad, pivot_table_par, xlsb_file_path):
    with pyxlsb.open_workbook(xlsb_file_path) as wb:
        for sheet_name in wb.sheets:
            with wb.get_sheet(sheet_name) as sheet:
                data = []
                for row in sheet.rows():
                    data_row = [item.v for item in row]
                    data.append(data_row)
                section_data = pd.DataFrame(data[2:], columns=data[1])

            with pyxlsb.open_workbook(xlsb_file_path, use_memory_mmap=True) as wb:
                writer = pd.ExcelWriter(xlsb_file_path, engine='pyxlsb')
                writer.book = wb

                for row1_value, pivot_table in zip(['AVA', 'PAD', 'PAR'], [pivot_table_ava, pivot_table_pad, pivot_table_par]):
                    if row1_value == sheet_name:
                        date_col_idx = section_data.columns.get_loc(date_str)

                        section_data[date_str] = pivot_table.set_index('SITE ID')['Data']
                        section_data.to_xls(writer, sheet_name, index=False, header=False, startrow=2)

                writer.save()
