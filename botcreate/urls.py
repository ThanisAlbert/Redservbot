"""redservcommitee URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.1/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from . import views

app_name="bot"

urlpatterns = [
    #==================Index========================
    path('', views.index, name='index'),
    path('create.html',views.createbot,name='create'),
    path('getfile.html/<int:param>/',views.file_download, name='getfile'),
    path('getfilenameslist.html/<int:param>/',views.file_getfilenamelist, name='getfilenameslist'),

    #===================Bot View===========================
    path('botviewquery.html', views.botviewquery, name='botviewquery'),
    path('botview_getfile.html/<int:param>/',views.botview_file_download, name='botview_getfile'),
    path('botview_getfilenameslist.html/<int:param>/',views.botview_file_getfilenamelist, name='botview_getfilenameslist'),
    path('downloadreport.html',views.downloadreport, name='downloadreport'),
    path('mailreport.html',views.mailreport, name='mailreport'),
   #====================Chart View===========================

    path('test.html',views.test,name='testview'),
    path('chart.html',views.sample_bar_chart,name='chartview'),
    path('changechart.html',views.changechartview,name='changechartview'),

]
