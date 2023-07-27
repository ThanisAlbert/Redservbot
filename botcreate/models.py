from django.db import models

# Create your models here.
class Bot(models.Model):

    Botno = models.IntegerField(primary_key=True)
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