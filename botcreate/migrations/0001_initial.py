# Generated by Django 4.1.4 on 2023-02-22 01:03

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Bot',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('Botno', models.IntegerField()),
                ('Botname', models.CharField(max_length=250)),
                ('Process', models.CharField(max_length=250)),
                ('Subprocess', models.CharField(max_length=250)),
                ('Spocname', models.CharField(max_length=250)),
                ('Requestormail', models.CharField(max_length=250)),
                ('Teamleadmail', models.CharField(max_length=250)),
                ('Managermail', models.CharField(max_length=250)),
                ('Developermail', models.CharField(max_length=250)),
                ('Technologyused', models.CharField(max_length=250)),
                ('Creationdate', models.DateField()),
                ('Startdate', models.DateField()),
                ('Enddate', models.DateField()),
                ('Workstatus', models.CharField(max_length=250)),
                ('Botstatus', models.CharField(max_length=250)),
                ('Manualtimespend', models.CharField(max_length=250)),
                ('Automationtimespend', models.CharField(max_length=250)),
                ('Totaltimesaved', models.CharField(max_length=250)),
                ('Totaldaysaved', models.CharField(max_length=250)),
                ('Kaizenawardstatus', models.CharField(max_length=250)),
                ('Kaizenawardyear', models.IntegerField()),
                ('Botdesc', models.CharField(max_length=500)),
                ('Mailrecipient', models.CharField(max_length=250)),
                ('Mailnotes', models.CharField(max_length=250)),
                ('Mailsend', models.BooleanField()),
            ],
        ),
    ]
