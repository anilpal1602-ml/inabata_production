import os
import pandas as pd
from openpyxl import load_workbook

def fix_entitas_nomor_aju_to_text(input_excel):
    if not os.path.isfile(input_excel):
        raise FileNotFoundError(f"Input file not found: {input_excel}")

    output_excel = input_excel  # overwrite same file

    wb = load_workbook(input_excel)
    sheet_names = wb.sheetnames

    sheets = {}
    for sheet in sheet_names:
        if sheet == "ENTITAS":
            df = pd.read_excel(
                input_excel,
                sheet_name=sheet,
                dtype={
                    "NOMOR AJU": str,
                    "NOMOR IDENTITAS": str
                }
            )
        else:
            df = pd.read_excel(input_excel, sheet_name=sheet)

        sheets[sheet] = df

    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        for sheet, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

            if sheet == "ENTITAS":
                workbook = writer.book
                worksheet = writer.sheets[sheet]
                text_format = workbook.add_format({"num_format": "@"})

                for col in ["NOMOR AJU", "NOMOR IDENTITAS"]:
                    if col in df.columns:
                        col_idx = df.columns.get_loc(col)
                        worksheet.set_column(col_idx, col_idx, None, text_format)

    return output_excel
