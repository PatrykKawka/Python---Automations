import win32com.client as win32
from datetime import datetime
import pythoncom

file_path = r'sciezkadopliku'
excel = win32.Dispatch('Excel.Application')
excel.Visible = True

try:
    workbook = excel.Workbooks.Open(file_path)
    workbook.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()

    now = datetime.now()
    date_str = now.strftime('%Y-%m-%d')
    new_file_path = rf'C:\Users\pkawk\OneDrive\Pulpit\Excel\Test\Kursy_walut_{date_str}.xlsx'
    excel.DisplayAlerts = False
    workbook.SaveAs(new_file_path)
    excel.DisplayAlerts = True

    print(f'Plik zapisany jako: {new_file_path}')

except Exception as e:
    print(f'Wystąpił błąd: {e}')
finally:
    workbook.Close(SaveChanges=False)
    excel.Quit()
