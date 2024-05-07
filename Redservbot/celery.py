from __future__ import absolute_import, unicode_literals
import os

from celery import Celery
from celery.schedules import crontab
from django.conf import settings

import botcreate.tasks

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Redservbot.settings')

app= Celery('Redservbot')
app.conf.enable_utc=False
app.conf.update(timezone='Asia/Kolkata')
app.config_from_object(settings,namespace='CELERY')

app.conf.beat_schedule={
    'mailsend':{
        'task':'botcreate.tasks.mail_admin',
        'schedule': crontab(hour=22,minute=50)
    }
}


#app.conf.beat_schedule={
#    'mailtest':{
#        'task':'botcreate.tasks.mailsupport',
#        'schedule': crontab(hour=22,minute=52)
#    }
#}

app.autodiscover_tasks()

@app.task(bind=True)
def debug_task(self):
    print(f'Request: {self.request!r}')