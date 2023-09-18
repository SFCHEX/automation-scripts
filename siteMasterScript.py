import pyxlsb
from openpyxl import load_workbook

def updateSiteData(AVA, PAD, PAR, date, filename):
    with pyxlsb.open_workbook(filename) as wb:
        with wb.get_sheet(1) as sheet:
            date_column_index = None
            for i, row in enumerate(sheet.rows()):
                if i == 1:
                    for j, cell in enumerate(row):
                        if cell.v == date:
                            date_column_index = j
                            break
                    if date_column_index is not None:
                        break

            if date_column_index is None:
                print(f"Date '{date}' not found in the sheet.")
                return

            site_id_column_index = None
            for i, row in enumerate(sheet.rows()):
                if i == 2:
                    for j, cell in enumerate(row):
                        if cell.v == "SITE ID":
                            site_id_column_index = j
                            break
                    if site_id_column_index is not None:
                        break

            if site_id_column_index is None:
                print("SITE ID subheader not found.")
                return

            for row in sheet.rows():
                site_id = row[site_id_column_index].v

                if site_id in AVA:
                    sheet.cell(row=row[site_id_column_index].r, column=date_column_index, value=AVA[site_id])

                if site_id in PAD:
                    sheet.cell(row=row[site_id_column_index].r, column=date_column_index, value=PAD[site_id])

                if site_id in PAR:
                    sheet.cell(row=row[site_id_column_index].r, column=date_column_index, value=PAR[site_id])

    wb.save(filename)
