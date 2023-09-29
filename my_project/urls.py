from django.urls import path
from .views import generate_excel_report

urlpatterns = [

    path('generate_excel_report/', generate_excel_report, name='generate_excel_report'),

]
