from django.urls import path
from .views import *

urlpatterns = [

    path('createsgs', createSgs2),
    # path('createsinglesgs', c)
    path('createxlsx', createXlsx),
    path('sample', sample)
]
