from django.contrib import admin

from .models import Customer, CustomerIndividual, CustomerLegalEntity


admin.site.register(Customer)
admin.site.register(CustomerIndividual)
admin.site.register(CustomerLegalEntity)
