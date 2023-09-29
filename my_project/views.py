from django.http import HttpResponse
from .generator import ExcelPropertyGenerator
from django.views import View


class MyView(View):
    def get(self, *args, **kwargs):
        result = ExcelPropertyGenerator("my_project/my_book.xlsx")
        result.generate_report()
        return HttpResponse("Hi world!")




