from django.shortcuts import render
import openpyxl

import os
SETTINGS_PATH = os.path.dirname(os.path.dirname(__file__))

def home(request):
    if "GET" == request.method:
        return render(request,SETTINGS_PATH+'/hello/templates/index.html', {})
    else:
        excel_file = request.FILES["excel_file"]

        # you may put validations here to check extension or file size

        wb = openpyxl.load_workbook(excel_file)

        # getting a particular sheet by name out of many sheets
        worksheet = wb["Sheet1"]
        print(worksheet)

        excel_data = list()
        # iterating over the rows and
        # getting value from each cell in row
        for row in worksheet.iter_rows():
            row_data = list()
            for cell in row:
                row_data.append(str(cell.value))
            excel_data.append(row_data)

        return render(request, SETTINGS_PATH+'/hello/templates/index.html', {"excel_data":excel_data})
