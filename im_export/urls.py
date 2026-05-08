from . import views
from django.urls import path

urlpatterns = [
    path("exports/insurees", views.export_insurees),
    path("imports/insurees", views.import_insurees),
    path("imports/exim_bank", views.import_exim_bank),
    path("imports/bdc_bank", views.import_bdc_bank),
    path("imports/other_payment", views.import_other_payment_method),
]
