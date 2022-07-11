from django.urls import path
from . import views

urlpatterns = [
    #path('', views.main, name='main'),
    path('view', views.view, name='view'),
    path('filter', views.filter_data, name='filter')
    #path('mail', views.send_file_to_email, name='mail')
    # path('daegu', views.daegu_view, name='daegu'),
    # path('busan', views.busan_view, name='busan'),
    # path('seoul', views.seoul_view, name='seoul')
]