# Generated by Django 4.1.4 on 2023-12-23 12:57

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('botcreate', '0019_bothist'),
    ]

    operations = [
        migrations.AddField(
            model_name='bothist',
            name='creationdate',
            field=models.DateField(blank=True, null=True),
        ),
    ]
