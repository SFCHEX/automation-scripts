
def write_pivot_to_xlsb(row1_values, date_str, pivot_table_ava, pivot_table_pad, pivot_table_par, xlsb_file_path):
    with pyxlsb.open_workbook(xlsb_file_path) as wb:
        section_data = {}
        for sheet in wb.sheets:
            data = []
            for row in sheet.rows():
                data_row = []
                for cell in row:
                    data_row.append(cell.v)
                data.append(data_row)
            section_data[sheet.name] = pd.DataFrame(data[2:], columns=data[1])

        writer = pd.ExcelWriter(xlsb_file_path, engine='openpyxl')
        writer.book = load_workbook(writer.path)
        
        for row1_value, pivot_table in zip(row1_values, [pivot_table_ava, pivot_table_pad, pivot_table_par]):
            if row1_value in section_data:
                section_df = section_data[row1_value]
                date_col_idx = section_df.columns.get_loc(date_str)

                writer.sheets = {ws.title: ws for ws in writer.book.worksheets}
                writer.sheets[row1_value] = writer.book.create_sheet(title=row1_value)

                section_df[date_str] = pivot_table.set_index('SITE ID')['Data']
                section_df.to_excel(writer, row1_value, startcol=date_col_idx, index=False, header=False)

        writer.save()
        writer.close()
