import csv

from django.contrib import admin
from django.http import HttpResponse
from import_export import resources
from import_export.admin import ImportExportModelAdmin

from .models import Bot, Process, Subprocess, Workstatus, Botstatus, Kaizenstatus, Kaizenawardedyear, Developermail, \
    Mailreport_to, Mailreport_cc, dbMailrecipient, BotHist, TicketTrackingTable
from django.contrib.admin.models import LogEntry
from django.contrib.auth.models import Group

# Register your models here.

class subprocessresource(resources.ModelResource):
    class Meta:
        model=Subprocess

class subprocessadmin(ImportExportModelAdmin):
    resource_class = subprocessresource
    list_display = ['subprocessname']

class processresource(resources.ModelResource):
    class Meta:
        model=Process

class processadmin(ImportExportModelAdmin):
    resource_class = processresource
    list_display = ['processname']

class botresource(resources.ModelResource):
    class Meta:
        model=Bot
        exclude=('id',)
        import_id_fields = ('Botno',)

class botadmin(ImportExportModelAdmin):
    resource_class = botresource
    search_fields = ['Botno']
    list_display = ['Botno', 'Botname', 'Process', 'Subprocess', 'Spocname', 'Requestormail', 'Teamleadmail',
                    'Managermail', 'Developermail', 'Technologyused', 'Creationdate', 'Startdate', 'Enddate',
                    'Workstatus', 'Botstatus', 'Manualtimespend', 'Automationtimespend', 'Totaltimesaved',
                    'Totaldaysaved', 'Kaizenawardstatus', 'Kaizenawardyear', 'Botdesc', 'Mailrecipient', 'Mailnotes',
                    'Mailsend','enhancestartdate','enhanceenddate','businessunit','livestatus','priority','categorization']


class bothistresource(resources.ModelResource):
    class Meta:
        model=BotHist


class bothistadmin(ImportExportModelAdmin):
    resource_class = bothistresource
    search_fields = ['botno']
    list_display = ['botno', 'botname', 'Developermail', 'botstatus', 'creationdate', 'startdate', 'enddate',
                    'enhancestartdate', 'enhanceenddate', 'livestatus', 'remarks', 'last_updated_datetime']


#class botadmin(admin.ModelAdmin):
#    list_display = ['Botno','Botname','Process','Subprocess','Spocname','Requestormail','Teamleadmail','Managermail','Developermail','Technologyused','Creationdate','Startdate','Enddate','Workstatus','Botstatus','Manualtimespend','Automationtimespend','Totaltimesaved','Totaldaysaved','Kaizenawardstatus','Kaizenawardyear','Botdesc','Mailrecipient','Mailnotes','Mailsend']
#    search_fields = ['Botno']


#LogEntry.objects.all().delete()
admin.site.unregister(Group)
admin.site.register(Bot,botadmin)
admin.site.register(Process,processadmin)
admin.site.register(Kaizenawardedyear)
admin.site.register(Developermail)
admin.site.register(Subprocess,subprocessadmin)
admin.site.register(Workstatus)
admin.site.register(Botstatus)
admin.site.register(TicketTrackingTable)
admin.site.register(Kaizenstatus)
admin.site.register(Mailreport_to)
admin.site.register(Mailreport_cc)
admin.site.register(dbMailrecipient)
admin.site.register(BotHist,bothistadmin)

admin.site.site_header = 'Redserv'                    # default: "Django Administration"
admin.site.index_title = 'Automation'                 # default: "Site administration"
admin.site.site_title = 'Automation'
#LogEntry.objects.all().delete()
