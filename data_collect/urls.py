from django.contrib import admin
from django.urls import path, include
from data_collect import views


urlpatterns = [
    path('', views.upload_file , name='upload_file_url' )
]
