# Generated by Django 4.1.4 on 2023-03-29 03:12

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('botcreate', '0009_alter_bot_process_alter_bot_subprocess'),
    ]

    operations = [
        migrations.AlterField(
            model_name='bot',
            name='Botstatus',
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AlterField(
            model_name='bot',
            name='Kaizenawardyear',
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
    ]
