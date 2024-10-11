import report_data as rp
import xlwings as xw
import os
from datetime import datetime
import warnings
import datetime
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")
# Get the current working directory
current_directory = os.getcwd()
output_file=os.path.join(current_directory,'output')
# Find the first .xlsx file in the current directory
for file_name in os.listdir(output_file):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(output_file, file_name)
            break
        else:
            pass
        
data_files=[]        
folder_path = os.path.join(os.getcwd(), 'sale_summary')  # Update with your file path        
for data_file in os.listdir(folder_path):
        if data_file.endswith('.xlsx'):
            data_files.append(os.path.join(folder_path, data_file))    
        else:
            raise FileNotFoundError("No .xlsx file found in the data directory.")
workbook=xw.Book(file_path)
weekly_dd=0
for sheet in workbook.sheets:
    if "DO NOT TOUCH" in sheet.name:
        sheet.range('D6').options(convert_to_number=True).number_format = '0.00'
        sheet.range('E6').options(convert_to_number=True).number_format = '0.00'
        continue
    else:
        # Access the sheet by name
        sheet = workbook.sheets[sheet]
        date=str(sheet.range('C3').value.date())
        for file in data_files:
            if date in file:
                if datetime.datetime.strptime(date, '%Y-%m-%d').weekday()==6:
                    weekly_dd+=rp.get_dd(file)
                    sheet['F11'].value=weekly_dd
                    sheet['F11'].number_format='0.00'
                    weekly_dd=0
                else:
                     weekly_dd+=rp.get_dd(file)
                if sheet['E4'].value==rp.get_location(file) and date == rp.get_date(file):
                    sheet['D6'].value=rp.get_def_amount(file)
                    sheet['D6'].number_format='0.00'
                    sheet['E6'].number_format='0.00'
                    sheet['D7'].value=rp.get_gc(file)
                    sheet['D8'].value=rp.get_saleWOTip(file)
                    sheet['D15'].value=rp.get_cctips(file)
                    sheet['E15'].value=rp.get_credit(file)
                    sheet['E21'].value=abs(rp.get_pettycash(file))
                    #sheet['C25'].value=rp.get_cash(file)
                    sheet['D22'].value=rp.get_ub_Tip(file)
                    sheet['D23'].value=rp.get_dd_Tip(file)
                    sheet['D24'].value=rp.get_gh_Tip(file)
                    #third-party
                    sheet['D10'].value=rp.get_uber(file)
                    sheet['D11'].value=rp.get_dd(file)
                    sheet['D12'].value=rp.get_gh(file)
                    sheet['D25'].value=''
                    if sheet['E21'].value!=0:
                        print(f"successfully filled out: {date}--a petty cash recorded")
                    else:
                         print(f"successfully filled out: {date}")
                else:
                     print(f"this file has wrong location: {date} {file}")
workbook.save(file_path)
workbook.close()

                        
