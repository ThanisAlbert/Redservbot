from django.db import models
from django.db.models import Func, IntegerField


class BotHist(models.Model):
    botno = models.IntegerField(null=True,blank=True)
    botname = models.CharField(max_length=250,null=True, blank=True)
    Developermail = models.CharField(max_length=250,null=True, blank=True)
    botstatus = models.CharField(max_length=250, null=True,blank=True)
    creationdate = models.DateField(null=True, blank=True)
    startdate = models.DateField(null=True,blank=True)
    enddate = models.DateField(null=True,blank=True)
    enhancestartdate = models.DateField(null=True,blank=True)
    enhanceenddate = models.DateField(null=True,blank=True)
    livestatus = models.CharField(null=True,max_length=250, blank=True)
    remarks = models.CharField(null=True,max_length=1000,blank=True)
    last_updated_datetime = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.botname

    class Meta:
        verbose_name = 'BotHistory'
        verbose_name_plural = 'BotHistory'

class CastAsInteger(Func):
    function = 'CAST'
    template = '%(function)s(%(expressions)s AS INTEGER)'
    output_field = IntegerField()


class TicketTrackingTable(models.Model):
    projectno = models.AutoField(primary_key=True)
    businessunit = models.CharField(max_length=24, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    activity = models.CharField(max_length=24, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    projectname = models.CharField(max_length=150, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    projecttype = models.CharField(max_length=24, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    projectdesc = models.TextField(db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    process = models.CharField(max_length=24, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    subprocess = models.CharField(max_length=60, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    spocname = models.CharField(max_length=80, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    requestormailid = models.CharField(max_length=150, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    teamleadmailid = models.CharField(max_length=150, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    managermailid = models.CharField(max_length=150, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    developermailid = models.CharField(max_length=150, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    technologyused = models.CharField(max_length=100, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    startdate = models.DateTimeField(blank=True, null=True)
    enddate = models.DateTimeField(blank=True, null=True)
    status = models.CharField(max_length=24, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    livestatus = models.CharField(max_length=16, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    devcomments = models.TextField(db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    manualtime = models.CharField(max_length=50, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    automationtime = models.CharField(max_length=50, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    totaltime = models.CharField(max_length=50, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    totalday = models.CharField(max_length=50, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    kaizenstatus = models.CharField(max_length=50, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    mailrecipient = models.CharField(max_length=700, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    enhancestartdate = models.DateTimeField(blank=True, null=True)
    enhanceenddate = models.DateTimeField(blank=True, null=True)
    priority = models.CharField(max_length=12, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    categorization = models.CharField(max_length=80, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    remarks = models.TextField(db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    idea_contributed_by = models.CharField(max_length=80, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    team_mates_name = models.CharField(max_length=500, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    idea_category = models.CharField(max_length=80, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    problem_statement = models.TextField(db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    solvable_problem = models.TextField(db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    action_plan = models.TextField(db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    help_required = models.TextField(db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    creationdate = models.DateTimeField()
    lastmodifieddate = models.DateTimeField()
    request_type = models.CharField(max_length=16, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    expected_benefit = models.CharField(max_length=1000, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    actual_benefit = models.CharField(max_length=1000, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)
    reason = models.CharField(max_length=80, db_collation='SQL_Latin1_General_CP1_CS_AS', blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'ticket_tracking_table'



# Create your models here.
class Bot(models.Model):

    Botno = models.IntegerField(primary_key=True,db_index=True)
    Botname =models.CharField(max_length=250,null=True, blank=True)
    Process = models.CharField(max_length=500,null=True, blank=True)
    Subprocess = models.CharField(max_length=500,null=True, blank=True)
    Spocname = models.CharField(max_length=250,null=True, blank=True)
    Requestormail = models.CharField(max_length=250,null=True, blank=True)
    Teamleadmail = models.CharField(max_length=250,null=True, blank=True)
    Managermail = models.CharField(max_length=250,null=True, blank=True)
    Developermail = models.CharField(max_length=250,null=True, blank=True)
    Technologyused = models.CharField(max_length=250,null=True, blank=True)
    Creationdate =models.DateField(null=True, blank=True)
    Startdate = models.DateField(null=True, blank=True)
    Enddate = models.DateField(null=True, blank=True)
    Workstatus = models.CharField(max_length=250,null=True, blank=True)
    Botstatus = models.CharField(max_length=250,null=True, blank=True)
    Manualtimespend = models.CharField(max_length=250,null=True, blank=True)
    Automationtimespend =models.CharField(max_length=250,null=True, blank=True)
    Totaltimesaved = models.CharField(max_length=250,null=True, blank=True)
    Totaldaysaved = models.CharField(max_length=250,null=True, blank=True)
    Kaizenawardstatus = models.CharField(max_length=250,null=True, blank=True)
    Kaizenawardyear = models.CharField(max_length=250,null=True,blank=True)
    Botdesc = models.CharField(max_length=500,null=True, blank=True)
    Mailrecipient = models.CharField(max_length=250,null=True, blank=True)
    Mailnotes = models.CharField(max_length=250,null=True, blank=True)
    Mailsend = models.BooleanField(null=True)
    enhancestartdate =models.DateField(null=True, blank=True)
    enhanceenddate = models.DateField(null=True, blank=True)
    businessunit = models.CharField(max_length=250,null=True,blank=True)
    priority = models.CharField(max_length=250,null=True,blank=True)
    livestatus = models.CharField(max_length=250,null=True,blank=True)
    remarks = models.CharField(max_length=1000, null=True, blank=True)
    categorization = models.CharField(max_length=250,null=True,blank=True)



    def __str__(self):
        return self.Botname

class Process(models.Model):
    processname = models.CharField(max_length=250)

    def __str__(self):
        return self.processname

    class Meta:
        verbose_name = 'Process'
        verbose_name_plural = 'Process'

class Subprocess(models.Model):
    subprocessname = models.CharField(max_length=250)

    def __str__(self):
        return self.subprocessname

    class Meta:
        verbose_name = 'SubProcess'
        verbose_name_plural = 'SubProcess'

class Workstatus(models.Model):
    workstatus = models.CharField(max_length=250)

    def __str__(self):
        return  self.workstatus

    class Meta:
        verbose_name = 'Workstatus'
        verbose_name_plural = 'Workstatus'

class Botstatus(models.Model):
    botstatus = models.CharField(max_length=250)

    def __str__(self):
        return  self.botstatus

    class Meta:
        verbose_name = 'Botstatus'
        verbose_name_plural = 'Botstatus'

class Kaizenstatus(models.Model):
    Kaizenstatus = models.CharField(max_length=250)

    class Meta:
        verbose_name = 'Kaizenstatus'
        verbose_name_plural = 'Kaizenstatus'

    def __str__(self):
        return  self.Kaizenstatus


class Kaizenawardedyear(models.Model):
    Kaizenawardedyear = models.CharField(max_length=250)

    class Meta:
        verbose_name = 'Kaizenawardedyear'
        verbose_name_plural = 'Kaizenawardedyear'

    def __str__(self):
        return  self.Kaizenawardedyear


class Developermail(models.Model):
    Developermail = models.CharField(max_length=250)

    def __str__(self):
        return self.Developermail

    class Meta:
        verbose_name = 'DeveloperMail'
        verbose_name_plural = 'DeveloperMail'


class dbMailrecipient(models.Model):
    dbmailrecipient = models.CharField(max_length=1000)

    def __str__(self):
        return self.dbmailrecipient

    class Meta:
        verbose_name = 'MailRecipient'
        verbose_name_plural = 'MailRecipient'


class Mailreport_to(models.Model):
    To_address = models.CharField(max_length=1000)

    def __str__(self):
        return self.To_address

    class Meta:
        verbose_name = 'Mailreport_to'
        verbose_name_plural = 'Mailreport_to'


class Mailreport_cc(models.Model):
    Cc_address = models.CharField(max_length=1000)

    def __str__(self):
        return self.Cc_address

    class Meta:
        verbose_name = 'Mailreport_cc'
        verbose_name_plural = 'Mailreport_cc'