import datetime
import shutil
import smtplib
from datetime import date
import os
import zipfile
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import json
from urllib.parse import quote
import requests

import xlwt
from django.db import connection
from django.db.models import Q, IntegerField, Max
from django.db.models.functions import Cast
from openpyxl import Workbook
from tablib import Dataset


from .models import Bot, Process, Subprocess, Workstatus, Botstatus, Kaizenstatus, Kaizenawardedyear, Developermail, \
    Mailreport_to, Mailreport_cc, dbMailrecipient
from django.http import HttpResponse
from django.shortcuts import render, redirect
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from django.utils.dateparse import parse_date
import logging
logger = logging.getLogger(__name__)
from .tasks import mail_admin


# Create your views here.

#=============================Index================================================

def index(request):
    logger.info("testing log")
    process = Process.objects.all().values()
    subprocess = Subprocess.objects.all().values()
    workstatus = Workstatus.objects.all().values()
    botstatus = Botstatus.objects.all().values()
    kaizenstatus = Kaizenstatus.objects.all().values()
    kaizenawardedyear = Kaizenawardedyear.objects.all().values()
    developermail = Developermail.objects.all().values()
    newbotno = Bot.objects.all().aggregate(Max('Botno'))

    try:
        Bots = Bot.objects.filter(Botstatus=request.POST["dropdownselect"]).order_by('Botno')
        #Bots = Bot.objects.filter(Botstatus=request.POST["dropdownselect"]).annotate(num=Cast('Botno', IntegerField())).order_by('num')
    except:
        #Bots = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Under User Testing') | Q(Botstatus='Bot No.Generated')).order_by('-Botno').values()
        Bots = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Bot No.Generated')).order_by('-Botno').values()
    if "botcreateerror" in request.session:
        context = {'newbotno':newbotno['Botno__max']+1,'developermail':developermail,'Botcreateerror':request.session['botcreateerror'], 'Bots': Bots, 'process': process, 'subprocess': subprocess, 'workstatus': workstatus,'botstatus': botstatus, 'kaizenstatus': kaizenstatus, 'kaizenawardedyear': kaizenawardedyear, }
        del request.session['botcreateerror']
    else:
        context = {'newbotno':newbotno['Botno__max']+1,'developermail':developermail,'Bots': Bots, 'process': process, 'subprocess': subprocess, 'workstatus': workstatus,'botstatus': botstatus, 'kaizenstatus': kaizenstatus, 'kaizenawardedyear': kaizenawardedyear, }

    return render(request,'bot/index.html',context)



def editbot(request,param):
    Botno = param
    # Bots = Bot.objects.all().values()
    #Bots = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Under User Testing') | Q(Botstatus='Bot No.Generated')).order_by('-Botno').values()
    Bots = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Bot No.Generated')).order_by('-Botno').values()

    Botobj = Bot.objects.filter(Botno=Botno).values()
    # Botobj = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Under User Testing') | Q(Botstatus='Bot No.Generated')).order_by('Botno')

    if Botobj.exists():
        pass
    else:
        return redirect('bot:index')

    process = Process.objects.all().values()
    subprocess = Subprocess.objects.all().values()
    workstatus = Workstatus.objects.all().values()
    botstatus = Botstatus.objects.all().values()
    kaizenawardedyear = Kaizenawardedyear.objects.all().values()
    kaizenstatus = Kaizenstatus.objects.all().values()
    developermail = Developermail.objects.all().values()
    Context = {'Botdetails': Botobj, 'developermail': developermail, 'Bots': Bots, 'process': process,
               'subprocess': subprocess, 'workstatus': workstatus, 'botstatus': botstatus, 'kaizenstatus': kaizenstatus,
               'kaizenawardedyear': kaizenawardedyear, }
    return render(request, 'bot/index_search.html', Context)


def createbot(request):

    if 'submit' in request.POST:
        Botno = request.POST['botno']
        Botname = request.POST["botname"]
        Process_var = request.POST["process"]
        Subprocess_var = request.POST["subprocess"]
        Spocname = request.POST["spocname"]
        Requestormail = request.POST["requestormail"]
        Teamleadmail = request.POST["teamleadmail"]
        Managermail = request.POST["managermail"]
        Developermail_var = request.POST["developermail"]
        Technologyused = request.POST["technology"]

        try:
            Creationdate = datetime.datetime.strptime(request.POST["creationdate"], '%d/%m/%Y')
        except:
            Creationdate = ""

        try:
            Startdate = datetime.datetime.strptime(request.POST["startdate"], '%d/%m/%Y')
        except:
            Startdate = ""

        try:
            Enddate = datetime.datetime.strptime(request.POST["enddate"], '%d/%m/%Y')
        except:
            Enddate = ""

        try:
            enhancestartdate = datetime.datetime.strptime(request.POST["enhancestart"], '%d/%m/%Y')
        except:
            enhancestartdate = ""

        try:
            enhanceenddate = datetime.datetime.strptime(request.POST["enhanceend"], '%d/%m/%Y')
        except:
            enhanceenddate = ""

        Botstatus_var = request.POST["botstatus"]
        Manualtime = request.POST["manualtime"]
        Automationtime = request.POST["automationtime"]
        Totaltime = request.POST["totaltime"]
        Totalday = request.POST["totalday"]
        Kaizenstatus_var = request.POST["kaizenstatus"]
        Kaizenyear = request.POST["kaizenyear"]
        Botdesc = request.POST["botdesc"]
        Mailrecipient = request.POST["mailrecipient"]
        businessUnit = request.POST["businessunit"]

        bot = Bot()
        bot.Botno = Botno
        bot.Botname = Botname
        bot.Process = Process_var
        bot.Subprocess = Subprocess_var
        bot.Spocname = Spocname
        bot.Requestormail = Requestormail
        bot.Teamleadmail = Teamleadmail
        bot.Managermail = Managermail
        bot.Developermail = Developermail_var
        bot.Technologyused = Technologyused
        try:
            bot.Creationdate = Creationdate.strftime("%Y-%m-%d")
        except:
            bot.Creationdate = None

        try:
            bot.Startdate = Startdate.strftime("%Y-%m-%d")
        except:
            bot.Startdate = None

        try:
            bot.Enddate = Enddate.strftime("%Y-%m-%d")
        except:
            bot.Enddate = None

        try:
            bot.enhancestartdate = enhancestartdate.strftime("%Y-%m-%d")
        except:
            bot.enhancestartdate =None

        try:
            bot.enhanceenddate = enhanceenddate.strftime("%Y-%m-%d")
        except:
            bot.enhanceenddate = None

        bot.Botstatus = Botstatus_var
        bot.Manualtimespend = Manualtime
        bot.Automationtimespend = Automationtime
        bot.Totaltimesaved = Totaltime
        bot.Totaldaysaved = Totalday
        bot.Kaizenawardstatus = Kaizenstatus_var
        bot.Kaizenawardyear = Kaizenyear
        bot.Botdesc = Botdesc
        bot.Mailrecipient = Mailrecipient
        bot.businessunit=businessUnit

        Botobj = Bot.objects.filter(Botno=Botno).values()

        if Botobj.exists():
            request.session['botcreateerror'] = "Bot already exist"
            return redirect('bot:index')
        else:
            bot.save()
            try:
                #myfile = request.FILES['myfile']
                myfiles = request.FILES.getlist('myfile')

                for myfile in myfiles:
                    fs = FileSystemStorage()
                    folder_name = Botno
                    folder_path = os.path.join("E:\\Botomation\\data\\storage1\\", folder_name)
                    #folder_path = os.path.join("C:\\Automation\\Python\\Redservbot\\Redservbot\\media", folder_name)
                    if not os.path.exists(folder_path):
                        os.makedirs(folder_path)
                    file = fs.save(os.path.join(folder_name, myfile.name), myfile)
            except:
                pass

            msg = MIMEMultipart('related')
            name = str(str(Requestormail).split("@")[0]).replace(".","")
            MESSAGE_BODY = """\
                <html>
                  <head>
                    <style>
                    table, th, td {
  border: 1px solid black;
  border-collapse: collapse;  
}
th{
background-color:#C7CBC7;
text-align: left;
}
                    </style>
                  </head>
                    <body>
                    
                      Dear Sender <br><br>
                      Greetings for the day!!!<br><br>
                      We have generated Redbot number for the below activity for future reference<br><br>
                      <table>
                        <tr>
                           <th>Bot No.</th>
                           <td>"""+ Botno +"""</td>
                        </tr> 
                        <tr>
                           <th>Process</th>
                           <td>"""+ Process_var+"""</td>
                        </tr>
                        <tr>
                           <th>Sub Process</th>
                           <td>"""+Subprocess_var+"""</td>
                        </tr> 
                        <tr>
                           <th>Bot Name</th>
                           <td>"""+Botname+"""</td>
                        </tr> 
                        <tr>
                           <th>Business Unit</th>
                           <td>""" + str(businessUnit) + """</td>
                        </tr> 
                        <tr>
                           <th>Contact Person</th>
                           <td>"""+ Spocname + """</td>
                        </tr> 
                        <tr>
                           <th>Developer Name</th>
                           <td>"""+ str(str(Developermail_var).split("@")[0]).replace("."," ") + """</td>
                        </tr>  
                        <tr>
                           <th>Manual Time Spend</th>
                           <td>Yet to be calculated</td>
                        </tr> 
                        <tr>
                           <th>Automation Time Spend</th>
                           <td>Yet to be calculated</td>
                        </tr>                    
                      </table>
                      <br>
                      Thanks and Regards,<br>Team Botomation
                    </body>
                </html>
                """

            to_addr = []
            to_addr.append(Requestormail)
            cc_addr=[]
            cc_addr.append(Teamleadmail)
            cc_addr.append(Managermail)
            cc_addr.append(Developermail_var)
            Mailrecipientres_query = dbMailrecipient.objects.all().values()

            if ";" in str(Mailrecipient):
                Mailrecipient_list = str(Mailrecipient).split(";")
                for recipient in Mailrecipient_list:
                    cc_addr.append(recipient)
            else:
                cc_addr.append(Mailrecipient)

            for res in Mailrecipientres_query:
                if ";" in str(res["dbmailrecipient"]):
                    Mailrecipient_list = str(res["dbmailrecipient"]).split(";")
                    for recipient in Mailrecipient_list:
                        cc_addr.append(recipient)
                else:
                    cc_addr.append(res["dbmailrecipient"])


            body_part = MIMEText(MESSAGE_BODY, 'html')

            msg['Subject'] = Botstatus_var + "-" + Botno + "-" + Subprocess_var + "-" +Botname
            #msg['Subject'] = "Bot Status for " + date.today().strftime('%d/%m/%Y')
            msg['From'] = "botomation@redingtongroup.com"
            msg['To'] = ', '.join(to_addr)
            msg['Cc'] = ', '.join(cc_addr)
            recipients=to_addr+cc_addr
            msg.attach(body_part)

            server = smtplib.SMTP("smtp.office365.com", 587)
            server.starttls()
            try:
                server.login("botomation@redingtongroup.com", "!Redb0t23#")
                server.sendmail(msg['From'], recipients, msg.as_string())
                server.quit()
            except Exception as e:
                return HttpResponse(e)

            return redirect('bot:index')

    if 'search' in request.POST:
        Botno = request.POST['botno']
        #Bots = Bot.objects.all().values()
        Bots = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Under User Testing') | Q(Botstatus='Bot No.Generated')).order_by('-Botno').values()
        #Bots = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Under User Testing') | Q(Botstatus='Bot No.Generated')).order_by('Botno')

        Botobj = Bot.objects.filter(Botno=Botno).values()
        #Botobj = Bot.objects.filter(Q(Botstatus='Devolopment In Progress') | Q(Botstatus='Under User Testing') | Q(Botstatus='Bot No.Generated')).order_by('Botno')

        if Botobj.exists():
            pass
        else:
            return redirect('bot:index')

        process = Process.objects.all().values()
        subprocess = Subprocess.objects.all().values()
        workstatus = Workstatus.objects.all().values()
        botstatus = Botstatus.objects.all().values()
        kaizenawardedyear = Kaizenawardedyear.objects.all().values()
        kaizenstatus = Kaizenstatus.objects.all().values()
        developermail = Developermail.objects.all().values()
        Context = {'Botdetails':Botobj,'developermail':developermail,'Bots': Bots,'process':process,'subprocess':subprocess,'workstatus':workstatus,'botstatus':botstatus,'kaizenstatus':kaizenstatus,'kaizenawardedyear':kaizenawardedyear,}
        return render(request,'bot/index_search.html',Context)

    if 'update' in request.POST:

        botno=request.POST['botno']

        try:
            botobj = Bot.objects.get(Botno=botno)
        except:
            request.session["botcreateerror"]="Bot Not Found"
            return redirect('bot:index')

        Botname = request.POST["botname"]
        Process_var = request.POST["process"]
        Subprocess_var = request.POST["subprocess"]
        Spocname = request.POST["spocname"]
        Requestormail = request.POST["requestormail"]
        Teamleadmail = request.POST["teamleadmail"]
        Managermail = request.POST["managermail"]
        Developermail_var = request.POST["developermail"]
        Technologyused = request.POST["technology"]
        businessUnit = request.POST["businessunit"]


        if Botname=="" and Requestormail=="":
            request.session["botcreateerror"] = "Please click search button"
            return redirect('bot:index')

        try:
            Creationdate = datetime.datetime.strptime(request.POST["creationdate"], '%d/%m/%Y')
        except Exception as e:
            Creationdate = ""

        try:
            Startdate = datetime.datetime.strptime(request.POST["startdate"], '%d/%m/%Y')
        except:
            Startdate = ""

        try:
            Enddate = datetime.datetime.strptime(request.POST["enddate"], '%d/%m/%Y')
        except:
            Enddate = ""

        try:
            enhancestartdate = datetime.datetime.strptime(request.POST["enhancestart"], '%d/%m/%Y')
        except:
            enhancestartdate = ""

        try:
            enhanceenddate = datetime.datetime.strptime(request.POST["enhanceend"], '%d/%m/%Y')
        except:
            enhanceenddate = ""


        Botstatus_var = request.POST["botstatus"]
        Manualtime = request.POST["manualtime"]
        Automationtime = request.POST["automationtime"]
        Totaltime = request.POST["totaltime"]
        Totalday = request.POST["totalday"]
        Kaizenstatus_var = request.POST["kaizenstatus"]
        Kaizenyear = request.POST["kaizenyear"]
        Botdesc = request.POST["botdesc"]
        Mailrecipient = request.POST["mailrecipient"]
        #Mailnotes = request.POST["mailnotes"]
        Mailsend = request.POST.getlist('mailsend')

        botobj.Botno = request.POST['botno']
        botobj.Botname = Botname
        botobj.Process = Process_var
        botobj.Subprocess = Subprocess_var
        botobj.Spocname = Spocname
        botobj.Requestormail = Requestormail
        botobj.Teamleadmail = Teamleadmail
        botobj.Managermail = Managermail
        botobj.Developermail = Developermail_var
        botobj.Technologyused = Technologyused
        try:
            botobj.Creationdate = Creationdate.strftime("%Y-%m-%d")
        except:
            botobj.Creationdate = None

        try:
            botobj.Startdate = Startdate.strftime("%Y-%m-%d")
        except:
            botobj.Startdate = None

        try:
            botobj.Enddate = Enddate.strftime("%Y-%m-%d")
        except:
            botobj.Enddate = None

        try:
            botobj.enhancestartdate = enhancestartdate.strftime("%Y-%m-%d")
        except:
            botobj.enhancestartdate =None

        try:
            botobj.enhanceenddate = enhanceenddate.strftime("%Y-%m-%d")
        except:
            botobj.enhanceenddate = None

        botobj.Botstatus = Botstatus_var
        botobj.Manualtimespend = Manualtime
        botobj.Automationtimespend = Automationtime
        botobj.Totaltimesaved = Totaltime
        botobj.Totaldaysaved = Totalday
        botobj.Kaizenawardstatus = Kaizenstatus_var
        botobj.Kaizenawardyear = Kaizenyear
        botobj.Botdesc = Botdesc
        botobj.Mailrecipient = Mailrecipient
        botobj.Mailnotes = ""
        botobj.businessunit = businessUnit
        mailcontent = ""

        botdir = "E:\\Botomation\\data\\storage1\\"

        #zf = zipfile.ZipFile(botdir + botno + ".zip", "w")
        #for dirname, subdirs, files in os.walk(botdir):
        #    zf.write(dirname)
        #    for filename in files:
        #        zf.write(os.path.join(dirname, filename))
        #zf.close()

        shutil.make_archive(botdir +botno,'zip',botdir +botno)

        if Botstatus_var == "Bot No.Generated":
            mailcontent = "This is to inform that bot development is under progress."

        if Botstatus_var =="Completed":
            mailcontent = "This is to inform that the bot development has been completed. Please let us know if any changes need to be done."

        if Botstatus_var == "Under User Testing":
            mailcontent = "This is to inform that the bot development has been completed. Please do UAT and let us know if any changes need to be done."

        if Botstatus_var == "Cancelled":
            mailcontent = "We regret to inform that bot development has been cancelled. Please contact us for further clarification"

        if Mailsend:
            msg = MIMEMultipart('related')
            name = str(str(Requestormail).split("@")[0]).replace(".","")
            MESSAGE_BODY = """\
                            <html>
                              <head>
                                <style>
                                table, th, td {
              border: 1px solid black;
              border-collapse: collapse;  
            }
            th{
            background-color:#C7CBC7;
            }
                                </style>
                              </head>
                                <body>

                                  Dear Sender <br><br>
                                  Greetings for the day!!!<br><br>
                                  
                                  """ + mailcontent + """ <br><br>
                                 
                                  <table>
                                    <tr>
                                       <th>Bot No.</th>
                                       <td>""" + botno + """</td>
                                    </tr> 
                                    <tr>
                                       <th>Process</th>
                                       <td>""" + Process_var + """</td>
                                    </tr>
                                    <tr>
                                       <th>Sub Process</th>
                                       <td>""" + Subprocess_var + """</td>
                                    </tr> 
                                    <tr>
                                       <th>Bot Name</th>
                                       <td>""" + Botname + """</td>
                                    </tr> 
                                    <tr>
                                       <th>Requested By</th>
                                       <td>""" + str(str(Requestormail).split("@")[0]).replace(".","") + """</td>
                                    </tr> 
                                    <tr>
                                       <th>Contact Person</th>
                                       <td>""" + Spocname + """</td>
                                    </tr> 
                                    <tr>
                                       <th>Developer Name</th>
                                       <td>""" + str(str(Developermail_var).split("@")[0]).replace("."," ") + """</td>
                                    </tr>
                                    <tr>
                                    <th>Manual Time Spend</th>
                                       <td>""" + Manualtime + """</td>
                                    </tr>
                                    <tr>
                                    <th>Automation Time Spend</th>
                                       <td>""" + Automationtime + """</td>
                                    </tr>                      
                                  </table>
                                  <br>
                                  Thanks and Regards,<br>Team Botomation
                                </body>
                            </html>
                            """

            to_addr = []
            to_addr.append(Requestormail)
            cc_addr = []
            cc_addr.append(Teamleadmail)
            cc_addr.append(Managermail)
            cc_addr.append(Developermail_var)
            Mailrecipientres_query = dbMailrecipient.objects.all().values()

            if ";" in str(Mailrecipient):
                Mailrecipient_list = str(Mailrecipient).split(";")
                for recipient in Mailrecipient_list:
                    cc_addr.append(recipient)
            else:
                cc_addr.append(Mailrecipient)

            for res in Mailrecipientres_query:
                if ";" in str(res["dbmailrecipient"]):
                    Mailrecipient_list = str(res["dbmailrecipient"]).split(";")
                    for recipient in Mailrecipient_list:
                        cc_addr.append(recipient)
                else:
                    cc_addr.append(res["dbmailrecipient"])


            body_part = MIMEText(MESSAGE_BODY, 'html')

            #msg['Subject'] = "Bot Status for " + date.today().strftime('%d/%m/%Y')
            msg['Subject'] = Botstatus_var + "-" + botno + "-" + Subprocess_var + "-" + Botname
            msg['From'] = "botomation@redingtongroup.com"
            msg['To'] = ', '.join(to_addr)
            msg['Cc'] = ', '.join(cc_addr)
            recipients = to_addr + cc_addr
            msg.attach(body_part)

            try:
                with open(botdir + botno + ".zip", 'rb') as file:
                    # Attach the file with filename to the email
                    msg.attach(MIMEApplication(file.read(), Name=botno + ".zip"))
            except Exception as e:
                print(e)

            server = smtplib.SMTP("smtp.office365.com", 587)
            server.starttls()
            try:
                server.login("botomation@redingtongroup.com", "!Redb0t23#")
                server.sendmail(msg['From'], recipients, msg.as_string())
                server.quit()
            except Exception as e:
                return HttpResponse(e)

        else:
            botobj.Mailsend = 0

        botobj.save()

        try:
            myfile = request.FILES['myfile']
            fs = FileSystemStorage()
            folder_name = botno
            folder_path = os.path.join(botdir, folder_name)
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            file = fs.save(os.path.join(folder_name, myfile.name), myfile)
            fileurl = fs.url(file)
        except:
            pass

        return redirect('bot:index')


def file_download(request,param):
    Botobj = Bot.objects.filter(Botno=param).values()
    for bot in Botobj:
        process =bot["Process"]
        subprocess = bot["Subprocess"]
        botname = bot["Botname"]
    #folder_path = "D:/COE IMPROVEMENTS/" + process + "/" + subprocess +"/" + botname
    folder_path = "E:/Botomation/data/storage1/"+str(param)
    try:
        file_list = os.listdir(folder_path)
        response = HttpResponse(content_type='application/zip')
        zip_file = zipfile.ZipFile(response, 'w')
        for file_name in file_list:
            file_path = os.path.join(folder_path, file_name)
            zip_file.write(file_path, file_name)
        zip_file.close()
        response['Content-Disposition'] = f'attachment; filename=' + botname + '.zip'
        return response
    except:
        request.session['botcreateerror']="FileNotFound"
        return redirect('bot:index')

def file_getfilenamelist(request,param):

    process = Process.objects.all().values()
    subprocess = Subprocess.objects.all().values()
    workstatus = Workstatus.objects.all().values()
    botstatus = Botstatus.objects.all().values()
    kaizenstatus = Kaizenstatus.objects.all().values()
    kaizenawardedyear = Kaizenawardedyear.objects.all().values()
    Bots = Bot.objects.all().values()

    Botobj = Bot.objects.filter(Botno=param).values()
    for bot in Botobj:
        processurl = bot["Process"]
        subprocessurl = bot["Subprocess"]
        botname = bot["Botname"]
        botno=bot["Botno"]

    #path = "D:/COE IMPROVEMENTS/" + processurl + "/" + subprocessurl +"/" + botname
    path = "E:/Botomation/data/storage1/" + str(param)
    files = []
    folders = []

    try:
        for f in os.listdir(path):
            if os.path.isfile(os.path.join(path, f)):
                files.append(f)
            else:
                folders.append(f)
    except:
        pass

    context = {'botno': botno, 'botname': botname, 'files': files, 'folders': folders, 'show_modal': True,
               'Bots': Bots, 'process': process, 'subprocess': subprocess, 'workstatus': workstatus,
               'botstatus': botstatus, 'kaizenstatus': kaizenstatus, 'kaizenawardedyear': kaizenawardedyear, }

    return render(request, 'bot/index.html',context)

#============================Bot View========================================

def botviewquery(request):

    try:
        from_date = request.POST["fromdateview"]
        to_date = request.POST["todateview"]
    except:
        from_date=""
        to_date=""

    if from_date == "":
        pass
    else:
        from_date = (datetime.datetime.strptime(str(from_date), "%d/%m/%Y").strftime("%Y-%m-%d"))

    if to_date == "":
        pass
    else:
        to_date = (datetime.datetime.strptime(str(to_date), "%d/%m/%Y").strftime("%Y-%m-%d"))

    if from_date=="":
        selected_options = request.POST.getlist('dropdownselect')
        if selected_options:
            if "All" in selected_options:
                Bots = Bot.objects.all().values().order_by('Botno')
            else:
                Bots = Bot.objects.filter(Botstatus__in=selected_options).order_by('Botno')
        else:
            Bots = Bot.objects.all().values().order_by('Botno')
    else:
        selected_options = request.POST.getlist('dropdownselect')
        if selected_options:
            if "All" in selected_options:
                #Bots = Bot.objects.filter(Startdate__range=[from_date, to_date]).order_by('Botno')
                Bots = Bot.objects.filter(Q(Startdate__range=(from_date, to_date)) | Q(Enddate__range=(from_date, to_date)) | Q(enhancestartdate__range=(from_date, to_date)) | Q(enhanceenddate__range=(from_date, to_date))).order_by('Botno')
            else:
                #Bots = Bot.objects.filter(Startdate__range=[from_date, to_date],Botstatus__in=selected_options).order_by('Botno')
                Bots = Bot.objects.filter(Q(Startdate__range=(from_date, to_date)) | Q(Enddate__range=(from_date, to_date)) | Q(enhancestartdate__range=(from_date, to_date)) | Q(enhanceenddate__range=(from_date, to_date)),Botstatus__in=selected_options).order_by('Botno')
        else:
            Bots = Bot.objects.filter(Q(Startdate__range=(from_date, to_date)) | Q(Enddate__range=(from_date, to_date))  | Q(enhancestartdate__range=(from_date, to_date)) | Q(enhanceenddate__range=(from_date, to_date)),Botstatus__in=selected_options).order_by('Botno')
            #Bots = Bot.objects.filter(Startdate__range=[from_date, to_date]).order_by('Botno')

    if "mailsent" in request.session:
        context = {'mailsent':request.session['mailsent'], 'Bots': Bots, }
        del request.session['mailsent']
    else:
        context = {'Bots': Bots,}

    return render(request,'bot/botview.html',context)


def botview_file_download(request,param):
    Botobj = Bot.objects.filter(Botno=param).values()
    for bot in Botobj:
        process =bot["Process"]
        subprocess = bot["Subprocess"]
        botname = bot["Botname"]
    #folder_path = "D:/COE IMPROVEMENTS/" + process + "/" + subprocess +"/" + botname
    folder_path = "E:/Botomation/data/storage1/"+str(param)
    try:
        file_list = os.listdir(folder_path)
        response = HttpResponse(content_type='application/zip')
        zip_file = zipfile.ZipFile(response, 'w')
        for file_name in file_list:
            file_path = os.path.join(folder_path, file_name)
            zip_file.write(file_path, file_name)
        zip_file.close()
        response['Content-Disposition'] = f'attachment; filename=' + botname + '.zip'
        return response
    except:
        request.session['botcreateerror']="FileNotFound"
        return redirect('bot:botviewquery')

def botview_file_getfilenamelist(request,param):

    process = Process.objects.all().values()
    subprocess = Subprocess.objects.all().values()
    workstatus = Workstatus.objects.all().values()
    botstatus = Botstatus.objects.all().values()
    kaizenstatus = Kaizenstatus.objects.all().values()
    kaizenawardedyear = Kaizenawardedyear.objects.all().values()
    Bots = Bot.objects.all().values()

    Botobj = Bot.objects.filter(Botno=param).values()
    for bot in Botobj:
        processurl = bot["Process"]
        subprocessurl = bot["Subprocess"]
        botname = bot["Botname"]
        botno=bot["Botno"]

    #path = "D:/COE IMPROVEMENTS/" + processurl + "/" + subprocessurl +"/" + botname
    path = "E:/Botomation/data/storage1/" + str(param)
    files = []
    folders = []

    try:
        for f in os.listdir(path):
            if os.path.isfile(os.path.join(path, f)):
                files.append(f)
            else:
                folders.append(f)
    except:
        pass

    context = {'botno': botno, 'botname': botname, 'files': files, 'folders': folders, 'show_modal': True,
               'Bots': Bots, 'process': process, 'subprocess': subprocess, 'workstatus': workstatus,
               'botstatus': botstatus, 'kaizenstatus': kaizenstatus, 'kaizenawardedyear': kaizenawardedyear, }

    return render(request, 'bot/botview.html',context)


def downloadreport(request):

    from_date = request.POST["fromdatenew"]
    to_date = request.POST["todatenew"]

    if from_date=="":
        pass
    else:
        from_date=(datetime.datetime.strptime(str(from_date),"%d/%m/%Y").strftime("%Y-%m-%d"))

    if to_date == "":
        pass
    else:
        to_date=(datetime.datetime.strptime(str(to_date), "%d/%m/%Y").strftime("%Y-%m-%d"))

    selected_options = request.POST.getlist('hiddenInput')

    if from_date=="":

        if len(selected_options[0])>0:
            if "All" in selected_options:
                my_queryset = Bot.objects.all().order_by('Botno')
                #my_queryset = Bot.objects.filter(Q(Startdate__range=(from_date, to_date)) | Q(
                #    enhancestartdate__range=(from_date, to_date))).order_by('Botno')

            else:
                my_queryset = Bot.objects.filter(Botstatus__in=selected_options).order_by('Botno')
                #my_queryset = Bot.objects.filter(
                #    Q(Startdate__range=(from_date, to_date)) | Q(enhancestartdate__range=(from_date, to_date)),
                #    Botstatus__in=selected_options).order_by('Botno')
        else:
            my_queryset = Bot.objects.all().order_by('Botno')
            #my_queryset = Bot.objects.filter(Q(Startdate__range=(from_date, to_date)) | Q(
            #    enhancestartdate__range=(from_date, to_date))).order_by('Botno')
    else:

        if len(selected_options[0])>0:
            if "All" in selected_options:
                my_queryset = Bot.objects.filter(
                    Q(Startdate__range=(from_date, to_date)) | Q(Enddate__range=(from_date, to_date)) | Q(
                        enhancestartdate__range=(from_date, to_date)) | Q(
                        enhanceenddate__range=(from_date, to_date))).order_by('Botno')
            else:
                my_queryset = Bot.objects.filter(
                    Q(Startdate__range=(from_date, to_date)) | Q(Enddate__range=(from_date, to_date)) | Q(
                        enhancestartdate__range=(from_date, to_date)) | Q(enhanceenddate__range=(from_date, to_date)),
                    Botstatus__in=selected_options).order_by('Botno')
        else:
            my_queryset = Bot.objects.filter(
                Q(Startdate__range=(from_date, to_date)) | Q(Enddate__range=(from_date, to_date)) | Q(
                    enhancestartdate__range=(from_date, to_date)) | Q(
                    enhanceenddate__range=(from_date, to_date))).order_by('Botno')

    wb = Workbook()

    ws = wb.active

    today = date.today()

    #ws.append(['Botno-1', 'Botname-2','Botdesc-20','Process-3','Subprocess-4','Spocname-5','RequestBy-6','Teamlead-7','DevelopmentBy-8','Technologyused-9','Creationdate-10','Startdate-11','Enddate-12','Botstatus-19','Manualtimespend-13','Automationtimespend-14','Totaltimesaved-15','Totaldaysaved-16','Kaizenawardstatus-17','Kaizenawardedyear-18'])
    ws.append(['Botno','Botname','Process','Subprocess','Spocname','RequestBy','Teamlead','DevelopmentBy','Technologyused','Creationdate','Startdate','Enddate','EnhancementStartdate','EnhancementEnddate','Manualtimespend','Automationtimespend','Totaltimesaved','Totaldaysaved','Kaizenawardstatus','Kaizenawardedyear','Botstatus','Botdesc'])

    for obj in my_queryset:
        if str(obj.Startdate)!="None":
            startdate = datetime.datetime.strptime(str(obj.Startdate),"%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            startdate=""

        if str(obj.Creationdate)!="None":
            creationdate = datetime.datetime.strptime(str(obj.Creationdate),"%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            creationdate=""

        if str(obj.Enddate)!="None":
            enddate = datetime.datetime.strptime(str(obj.Enddate),"%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            enddate=""

        if str(obj.enhancestartdate)!="None":
            enhancestartdate = datetime.datetime.strptime(str(obj.enhancestartdate),"%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            enhancestartdate=""

        if str(obj.enhanceenddate)!="None":
            enhanceenddate = datetime.datetime.strptime(str(obj.enhanceenddate),"%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            enhanceenddate=""

        if "mins" in str(obj.Totaldaysaved) and "days" not in str(obj.Totaldaysaved):
            totaldaysaved = round(int(str(obj.Totaldaysaved).replace("mins", "")) / 60)
            if totaldaysaved <= 8:
                totaldaysaved = str(totaldaysaved) + " hrs"
            else:
                totaldaysaved = str(round(totaldaysaved / 8)) + " days"
        else:
            totaldaysaved = obj.Totaldaysaved

        if str(totaldaysaved) == "None":
            totaldaysaved = ""

        if str(obj.Manualtimespend) == "None":
            manualtimespend = ""
        else:
            manualtimespend = str(obj.Manualtimespend)

        if str(obj.Automationtimespend) == "None":
            automationtimespend = ""
        else:
            automationtimespend = str(obj.Automationtimespend)

        if str(obj.Totaltimesaved) == "None":
            Totaltimesaved = ""
        else:
            Totaltimesaved = str(obj.Totaltimesaved)

        ws.append([obj.Botno, obj.Botname, obj.Process, obj.Subprocess, obj.Spocname, obj.Requestormail, obj.Teamleadmail,
             obj.Developermail, obj.Technologyused, creationdate, startdate, enddate, enhancestartdate, enhanceenddate,
             manualtimespend, automationtimespend, Totaltimesaved, totaldaysaved, obj.Kaizenawardstatus,
             obj.Kaizenawardyear, obj.Botstatus, obj.Botdesc])

    # Generate the response
    '''
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 0.5
        ws.column_dimensions[column].width = adjusted_width
    '''

    response = HttpResponse(content_type='application/vnd.ms-excel')
    response['Content-Disposition'] = 'attachment; filename="BotDetails_{}.xlsx"'.format(today.strftime("%d-%m-%Y"))
    wb.save(response)

    return response


def mailreport(request):

    options = Options()
    options.headless = True  # Run Chrome in headless mode

    url = "http://172.26.1.19:85/bot/chart.html"
    #url = "http://172.24.3.13:85/bot/chart.html"

    driver = webdriver.Chrome(executable_path="C:\\chromedriver.exe", options=options)
    #driver = webdriver.Chrome(executable_path="D:\\chromedriver.exe", options=options)

    Totalcount = Bot.objects.all().count()

    driver.get(url)

    # Wait for the chart to be generated (replace "chart-element-id" with the actual ID of the chart element)
    chart_element = driver.find_element_by_id("chartContainer")

    time.sleep(4)

    chart_image = chart_element.screenshot_as_png

    with open(settings.MEDIA_ROOT + "\\chart.png", "wb") as f:
        f.write(chart_image)


    from_date = request.POST["fromdatemailnew"]
    to_date = request.POST["todatemailnew"]

    if from_date == "":
        pass
    else:
        from_date = (datetime.datetime.strptime(str(from_date), "%d/%m/%Y").strftime("%Y-%m-%d"))

    if to_date == "":
        pass
    else:
        to_date = (datetime.datetime.strptime(str(to_date), "%d/%m/%Y").strftime("%Y-%m-%d"))


    selected_options = request.POST.getlist('hiddenInputmail')
    for select in selected_options:
        temp = str(select)

    try:
        selected_options = temp.split(",")
    except:
        selected_options = ""


    if from_date=="":

        if len(selected_options[0])>0:
            if "All" in selected_options:
                my_queryset = Bot.objects.all().order_by('Botno')
            else:
                my_queryset = Bot.objects.filter(Botstatus__in=selected_options).order_by('Botno')
        else:
            my_queryset = Bot.objects.all().order_by('Botno')
    else:

        if len(selected_options[0])>0:
            if "All" in selected_options:
                my_queryset = Bot.objects.all().filter(Startdate__range=[from_date, to_date]).order_by('Botno')
            else:
                my_queryset = Bot.objects.filter(Startdate__range=[from_date, to_date],Botstatus__in=selected_options).order_by('Botno')
        else:
            my_queryset = Bot.objects.all().filter(Startdate__range=[from_date, to_date]).order_by('Botno')


    wb = Workbook()
    ws = wb.active

    headers = ['Botno','Botname','Process','Subprocess','Spocname','RequestBy','Teamlead','DevelopmentBy','Technologyused','Creationdate','Startdate','Enddate','Enhancementstartdate','Enhancementenddate','Manualtimespend','Automationtimespend','Totaltimesaved','Totaldaysaved','Kaizenawardstatus','Kaizenawardedyear','Botstatus','Botdesc']

    ws.append(headers)

    for obj in my_queryset:
        if str(obj.Creationdate) != "None":
            creationdate = datetime.datetime.strptime(str(obj.Creationdate), "%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            creationdate = ""

        if str(obj.Startdate) != "None":
            startdate = datetime.datetime.strptime(str(obj.Startdate), "%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            startdate = ""

        if str(obj.Enddate) != "None":
            enddate = datetime.datetime.strptime(str(obj.Enddate), "%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            enddate = ""

        if str(obj.enhancestartdate) != "None":
            enhancementstartdate = datetime.datetime.strptime(str(obj.enhancestartdate), "%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            enhancementstartdate = ""


        if str(obj.enhanceenddate) != "None":
            enhancementenddate = datetime.datetime.strptime(str(obj.enhanceenddate), "%Y-%m-%d").strftime("%d-%m-%Y")
        else:
            enhancementenddate = ""

        row = [obj.Botno,obj.Botname,obj.Process,obj.Subprocess,obj.Spocname,obj.Requestormail,obj.Teamleadmail,obj.Developermail,obj.Technologyused,creationdate,startdate,enddate,enhancementstartdate,enhancementenddate,obj.Manualtimespend,obj.Automationtimespend,obj.Totaltimesaved,obj.Totaldaysaved,obj.Kaizenawardstatus,obj.Kaizenawardyear,obj.Botstatus,obj.Botdesc]

        ws.append(row)

    wb.save(settings.MEDIA_ROOT + "\\RedBot.xlsx")

    msg = MIMEMultipart('related')
    db_name = "All"
    MESSAGE_BODY ="""\
    <html>
      <head></head>
        <body>
          Dear """ + db_name + """ <br><br>
          Greetings for the day!!!<br><br>
          Please find the Bot status for """ +date.today().strftime('%d/%m/%Y')+"""<br><br>
          <br> <h4>Total Bots: """+ str(Totalcount) +"""</h4><br><img src="cid:image1"> <br><br> Thanks,<br>Team Botomation 
        </body>
    </html>
    """

    body_part = MIMEText(MESSAGE_BODY, 'html')

    to_addr_queryset = Mailreport_to.objects.all().values()
    for to_addr in to_addr_queryset:
        to_addr = str(to_addr["To_address"]).split(";")

    cc_addr_queryset = Mailreport_cc.objects.all().values()
    for cc_addr in cc_addr_queryset:
        cc_addr = str(cc_addr["Cc_address"]).split(";")

    #to_addr = []
    #to_addr.append("rathina.moorthy@redingtongroup.com")
    #to_addr.append("thanis.albert@redingtongroup.com")
    #cc_addr = []
    #cc_addr.append("thanis.a@redingtongroup.com")

    recipients = to_addr + cc_addr
    #recipients = 'rathina.moorthy@redingtongroup.com,thanis.albert@redingtongroup.com'
    #recipients = 'thanis.albert@redingtongroup.com'

    msg['Subject'] = "Bot Status for "+date.today().strftime('%d/%m/%Y')
    msg['From'] = "botomation@redingtongroup.com"
    msg['To'] = ', '.join(to_addr)
    msg['Cc'] = ', '.join(cc_addr)
    msg.attach(body_part)

    with open(settings.MEDIA_ROOT +"\\chart.png", 'rb') as f:
        img_data = f.read()
    img = MIMEImage(img_data)
    img.add_header('Content-ID', '<image1>')
    msg.attach(img)

    with open(settings.MEDIA_ROOT +"\\RedBot.xlsx", 'rb') as file:
        msg.attach(MIMEApplication(file.read(), Name="RedBotReport.xlsx"))

    server = smtplib.SMTP("smtp.office365.com", 587)
    server.starttls()
    try:
        server.login("botomation@redingtongroup.com", "!Redb0t23#")
        server.sendmail(msg['From'], recipients, msg.as_string())
        server.quit()
        request.session["mailsent"]="sent"
        return redirect('bot:botviewquery')
    except Exception as e:
        return HttpResponse(e)

#===========================Chart View====================================

def sample_bar_chart(request):
    Totalcount = Bot.objects.all().count()
    Botnogenerated = Bot.objects.filter(Botstatus='Bot No.Generated').count()
    devlopment_in_progress = Bot.objects.filter(Botstatus='Devolopment In Progress').count()
    usertesting = Bot.objects.filter(Botstatus='Under User Testing').count()
    usertestanddev = Bot.objects.filter(Botstatus='User Testing & Dev Enhancement').count()
    tm = Bot.objects.filter(Botstatus='TM TO BE DONE').count()
    completed = Bot.objects.filter(Botstatus='Completed').count()
    cancelled = Bot.objects.filter(Botstatus='Cancelled').count()

    datapoints = [
        {"y": Botnogenerated, "label": "BotCreated","indexLabelPlacement": "outside"},
        {"y": devlopment_in_progress , "label": "Devolopment In Progress", "indexLabelPlacement": "outside"},
        {"y": usertesting, "label": "UserTesting","indexLabelPlacement": "outside"},
        {"y": usertestanddev, "label": "UserTesting_and_devenhancement", "indexLabelPlacement": "outside"},
        {"y": tm, "label": "Completed", "indexLabelPlacement": "outside"},
        {"y": completed, "label": "TimeStudy", "indexLabelPlacement": "outside"},
        {"y": cancelled, "label": "Cancelled", "indexLabelPlacement": "outside"},
    ]

    bot_data = [
        {"label": "BotNoCreated", "y": Botnogenerated},
        {"label": "Devolopment In Progress", "y": devlopment_in_progress },
        {"label": "UserTesting", "y": usertesting},
        {"label": "UserTesting_and_devenhancement", "y": usertestanddev},
        {"label": "Completed", "y": tm},
        {"label": "TimeStudy", "y": completed},
        {"label": "Cancelled", "y": cancelled}
    ]

    return render(request, 'bot/chart.html', {"datapoints": datapoints,"bot_data" : bot_data,"chartview":"bar", "Totalcount":Totalcount})


def changechartview(request):

    fromdate = request.POST["fromdatenew"]
    todate = request.POST["todatenew"]

    if fromdate == "":
        pass
    else:
        fromdate = (datetime.datetime.strptime(str(fromdate), "%d/%m/%Y").strftime("%Y-%m-%d"))

    if todate == "":
        pass
    else:
        todate = (datetime.datetime.strptime(str(todate), "%d/%m/%Y").strftime("%Y-%m-%d"))

    selected_options = request.POST.getlist('hiddenInput')
    for select in selected_options:
        temp = str(select)

    try:
        selected_options = temp.split(",")
    except:
        selected_options = ""

    Totalcount = Bot.objects.all().count()
    charttype = request.POST["chartselect"]

    if fromdate == "":
        if len(selected_options[0]) > 0:
            if "All" in selected_options:
                Botnogenerated = Bot.objects.filter(Botstatus='Bot No.Generated').count()
                devlopment_in_progress = Bot.objects.filter(Botstatus='Devolopment In Progress').count()
                usertesting = Bot.objects.filter(Botstatus='Under User Testing').count()
                usertestanddev = Bot.objects.filter(Botstatus='User Testing & Dev Enhancement').count()
                tm = Bot.objects.filter(Botstatus='TM TO BE DONE').count()
                completed = Bot.objects.filter(Botstatus='Completed').count()
                cancelled = Bot.objects.filter(Botstatus='Cancelled').count()
            else:
                Botnogenerated = Bot.objects.filter(Botstatus='Bot No.Generated', Botstatus__in=selected_options).count()
                devlopment_in_progress = Bot.objects.filter(Botstatus='Devolopment In Progress', Botstatus__in=selected_options).count()
                usertesting = Bot.objects.filter(Botstatus='Under User Testing', Botstatus__in=selected_options).count()
                usertestanddev = Bot.objects.filter(Botstatus='User Testing & Dev Enhancement', Botstatus__in=selected_options).count()
                tm = Bot.objects.filter(Botstatus='TM TO BE DONE', Botstatus__in=selected_options).count()
                completed = Bot.objects.filter(Botstatus='Completed', Botstatus__in=selected_options).count()
                cancelled = Bot.objects.filter(Botstatus='Cancelled', Botstatus__in=selected_options).count()
        else:
            Botnogenerated = Bot.objects.filter(Botstatus='Bot No.Generated').count()
            devlopment_in_progress = Bot.objects.filter(Botstatus='Devolopment In Progress').count()
            usertesting = Bot.objects.filter(Botstatus='Under User Testing').count()
            usertestanddev = Bot.objects.filter(Botstatus='User Testing & Dev Enhancement').count()
            tm = Bot.objects.filter(Botstatus='TM TO BE DONE').count()
            completed = Bot.objects.filter(Botstatus='Completed').count()
            cancelled = Bot.objects.filter(Botstatus='Cancelled').count()

    else:
        if len(selected_options[0]) > 0:
            if "All" in selected_options:
                Botnogenerated = Bot.objects.filter(Botstatus='Bot No.Generated',Startdate__range=[fromdate, todate]).count()
                devlopment_in_progress = Bot.objects.filter(Botstatus='Devolopment In Progress',Startdate__range=[fromdate, todate]).count()
                usertesting = Bot.objects.filter(Botstatus='Under User Testing',Startdate__range=[fromdate, todate]).count()
                usertestanddev = Bot.objects.filter(Botstatus='User Testing & Dev Enhancement',Startdate__range=[fromdate, todate]).count()
                tm = Bot.objects.filter(Botstatus='TM TO BE DONE',Startdate__range=[fromdate, todate]).count()
                completed = Bot.objects.filter(Botstatus='Completed',Startdate__range=[fromdate, todate]).count()
                cancelled = Bot.objects.filter(Botstatus='Cancelled',Startdate__range=[fromdate, todate]).count()
            else:
                Botnogenerated = Bot.objects.filter(Botstatus='Bot No.Generated',Startdate__range=[fromdate, todate],Botstatus__in=selected_options).count()
                devlopment_in_progress = Bot.objects.filter(Botstatus='Devolopment In Progress',Startdate__range=[fromdate, todate],Botstatus__in=selected_options).count()
                usertesting = Bot.objects.filter(Botstatus='Under User Testing',Startdate__range=[fromdate, todate],Botstatus__in=selected_options).count()
                usertestanddev = Bot.objects.filter(Botstatus='User Testing & Dev Enhancement',Startdate__range=[fromdate, todate],Botstatus__in=selected_options).count()
                tm = Bot.objects.filter(Botstatus='TM TO BE DONE',Startdate__range=[fromdate, todate],Botstatus__in=selected_options).count()
                completed = Bot.objects.filter(Botstatus='Completed',Startdate__range=[fromdate, todate],Botstatus__in=selected_options).count()
                cancelled = Bot.objects.filter(Botstatus='Cancelled',Startdate__range=[fromdate, todate],Botstatus__in=selected_options).count()
        else:
            Botnogenerated = Bot.objects.filter(Botstatus='Bot No.Generated',Startdate__range=[fromdate, todate]).count()
            devlopment_in_progress = Bot.objects.filter(Botstatus='Devolopment In Progress',Startdate__range=[fromdate, todate]).count()
            usertesting = Bot.objects.filter(Botstatus='Under User Testing',Startdate__range=[fromdate, todate]).count()
            usertestanddev = Bot.objects.filter(Botstatus='User Testing & Dev Enhancement',Startdate__range=[fromdate, todate]).count()
            tm = Bot.objects.filter(Botstatus='TM TO BE DONE',Startdate__range=[fromdate, todate]).count()
            completed = Bot.objects.filter(Botstatus='Completed',Startdate__range=[fromdate, todate]).count()
            cancelled = Bot.objects.filter(Botstatus='Cancelled',Startdate__range=[fromdate, todate]).count()


    datapoints=[]
    bot_data = []

    Botnogenerated_var={"y": Botnogenerated, "label": "BotNoGenerated", "indexLabelPlacement": "outside"}
    developmentinprogress_var = {"y": devlopment_in_progress, "label": "Devolopment In Progress", "indexLabelPlacement": "outside"}
    usertesting_var = {"y": usertesting, "label": "UserTesting", "indexLabelPlacement": "outside"}
    usertestanddev_var = {"y": usertestanddev, "label": "UserTesting_and_devenhancement", "indexLabelPlacement": "outside"}
    tm_var = {"y": tm, "label": "TimeStudy", "indexLabelPlacement": "outside"}
    completed_var = {"y": completed, "label": "Completed", "indexLabelPlacement": "outside"}
    cancelled_var = {"y": cancelled, "label": "Cancelled", "indexLabelPlacement": "outside"}

    botnogen_label = {"label": "BotNoGenerated", "y": Botnogenerated}
    dev_label = {"label": "Devolopment In Progress", "y": devlopment_in_progress}
    usertest_label = {"y": usertesting, "label": "UserTesting", "indexLabelPlacement": "outside"}
    usertestanddev_label = {"y": usertestanddev, "label": "UserTestandDev", "indexLabelPlacement": "outside"}
    tm_label = {"y": tm, "label": "TimeStudy", "indexLabelPlacement": "outside"}
    comp_label = {"y": completed, "label": "Completed", "indexLabelPlacement": "outside"}
    canc_label = {"y": cancelled, "label": "Cancelled", "indexLabelPlacement": "outside"}

    if "All" in selected_options:

        if Botnogenerated > 0:
            datapoints.append(Botnogenerated_var)
            bot_data.append(botnogen_label)

        if devlopment_in_progress > 0:
            datapoints.append(developmentinprogress_var)
            bot_data.append(dev_label)

        if usertesting > 0:
            datapoints.append(usertesting_var)
            bot_data.append(usertest_label)

        if usertestanddev > 0:
            datapoints.append(usertestanddev_var)
            bot_data.append(usertestanddev_label)

        if tm > 0:
            datapoints.append(tm_var)
            bot_data.append(tm_label)

        if completed > 0:
            datapoints.append(completed_var)
            bot_data.append(comp_label)

        if cancelled > 0:
            datapoints.append(cancelled_var)
            bot_data.append(canc_label)

    else:

        if Botnogenerated > 0 and "Bot No.Generated" in selected_options:
            datapoints.append(Botnogenerated_var)
            bot_data.append(botnogen_label)

        if devlopment_in_progress > 0  and "Devolopment In Progress" in selected_options:
            datapoints.append(developmentinprogress_var)
            bot_data.append(dev_label)

        if usertesting > 0  and "Under User Testing" in selected_options:
            datapoints.append(usertesting_var)
            bot_data.append(usertest_label)

        if usertestanddev > 0  and "User Testing & Dev Enhancement" in selected_options:
            datapoints.append(usertestanddev_var)
            bot_data.append(usertestanddev_label)

        if tm > 0  and "TM TO BE DONE" in selected_options:
            datapoints.append(tm_var)
            bot_data.append(tm_label)

        if completed > 0  and "Completed" in selected_options:
            datapoints.append(completed_var)
            bot_data.append(comp_label)

        if cancelled > 0  and "Cancelled" in selected_options:
            datapoints.append(cancelled_var)
            bot_data.append(canc_label)

    '''    
    datapoints = [
        {"y": Botnogenerated, "label": "BotNoGenerated", "indexLabelPlacement": "outside"},
        {"y": devlopment_in_progress, "label": "Devolopment In Progress", "indexLabelPlacement": "outside"},
        {"y": usertesting, "label": "UserTesting", "indexLabelPlacement": "outside"},
        {"y": usertestanddev, "label": "UserTesting_and_devenhancement", "indexLabelPlacement": "outside"},
        {"y": tm, "label": "Completed", "indexLabelPlacement": "outside"},
        {"y": completed, "label": "TimeStudy", "indexLabelPlacement": "outside"},
        {"y": cancelled, "label": "Cancelled", "indexLabelPlacement": "outside"},
    ] 
    bot_data = [
        {"label": "BotNoGenerated", "y": Botnogenerated},
        {"label": "Devolopment In Progress", "y": devlopment_in_progress},
        {"label": "UserTesting", "y": usertesting},
        {"label": "UserTesting_and_devenhancement", "y": usertestanddev},
        {"label": "Completed", "y": tm},
        {"label": "TimeStudy", "y": completed},
        {"label": "Cancelled", "y": cancelled}
    ]
    '''

    if charttype=="pie":
        return render(request, 'bot/chart.html', {"datapoints": datapoints, "bot_data": bot_data,"chartview":"pie","Totalcount":Totalcount})
    if charttype=="bar":
        return render(request, 'bot/chart.html', {"datapoints": datapoints, "bot_data": bot_data,"chartview":"bar","Totalcount":Totalcount})
    if charttype=="select":
        return HttpResponse("select")

#========================Test View=======================

def test(request):
    return render(request,'bot/test.html')
















