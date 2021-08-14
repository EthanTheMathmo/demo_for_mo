import xlwings as xw
sheet_info_name = "test_info_sheet.txt"

def input_link():
    xw_app = xw.apps.active
    wb = xw_app.books.active
    sht = wb.sheets.active

    book_name = wb.name 
    sheet_name = sht.name

    

    x=1


if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('Demo.xlsm').set_mock_caller()
    input_link()
