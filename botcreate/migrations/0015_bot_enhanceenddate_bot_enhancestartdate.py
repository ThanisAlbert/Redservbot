# Generated by Django 4.1.4 on 2023-05-09 09:47

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('botcreate', '0014_dbmailrecipient_delete_mailrecipient'),
    ]

    operations = [
        migrations.AddField(
            model_name='bot',
            name='enhanceenddate',
            field=models.DateField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name='bot',
            name='enhancestartdate',
            field=models.DateField(blank=True, null=True),
        ),
    ]