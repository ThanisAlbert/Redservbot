# Generated by Django 4.1.4 on 2023-03-23 11:03

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('botcreate', '0007_alter_bot_botstatus_alter_bot_kaizenawardyear_and_more'),
    ]

    operations = [
        migrations.RemoveField(
            model_name='bot',
            name='id',
        ),
        migrations.AlterField(
            model_name='bot',
            name='Botno',
            field=models.IntegerField(primary_key=True, serialize=False),
        ),
    ]
