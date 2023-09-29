from django.http import HttpResponse
from .models import *


def generate_excel_report(request):
    file_path = "my_book.xlsx"
    generated_excel = ExcelPropertyGenerator(file_path)
    generated_excel.generate_report()

    with open(file_path, 'rb') as excel_file:
        response = HttpResponse(excel_file.read(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response['Content-Disposition'] = f"attachment; filename=отчет.xlsx"

    return response

