import sqlite3

from botcreate.models import TicketTrackingTable


class sqlite:

    def __init__(self):
        self.db_path = 'E:\\Botomation\\Redsevbotvenv\\Redservbot\\db.sqlite3'
        self.connection = sqlite3.connect(self.db_path)
        self.cursor = self.connection.cursor()

    def process(self):
        process = []
        table_name = 'botcreate_process'
        self.cursor.execute(f'SELECT * FROM {table_name}')
        rows = self.cursor.fetchall()
        for row in rows:
            process.append(row[1])
        return process

    def subprocess(self):
        subprocess = []
        table_name = 'botcreate_subprocess'
        self.cursor.execute(f'SELECT * FROM {table_name}')
        rows = self.cursor.fetchall()
        for row in rows:
            subprocess.append(row[1])
        return subprocess

    def botstatus(self):
        status = []
        table_name = 'botcreate_botstatus'
        self.cursor.execute(f'SELECT * FROM {table_name}')
        rows = self.cursor.fetchall()
        for row in rows:
            print(row)
            status.append(row[1])
        return status


    def kaizenstatus(self):
        kaizenstatus = []
        table_name = 'botcreate_kaizenstatus'
        self.cursor.execute(f'SELECT * FROM {table_name}')
        rows = self.cursor.fetchall()
        for row in rows:
            kaizenstatus.append(row[1])
        return kaizenstatus


    def livestatus(self):
        livestatus = []
        table_name = 'botcreate_kaizenawardedyear'
        self.cursor.execute(f'SELECT * FROM {table_name}')
        rows = self.cursor.fetchall()
        for row in rows:
            livestatus.append(row[1])
        return livestatus

    def developermail(self):
        developermail = []
        table_name = 'botcreate_developermail'
        self.cursor.execute(f'SELECT * FROM {table_name}')
        rows = self.cursor.fetchall()
        for row in rows:
            developermail.append(row[1])
        return developermail

    def newbotno(self):
        tickettable = TicketTrackingTable.objects.all().values()
        bot_set = set()
        for bot in tickettable:
            bot_set.add(bot["projectno"])
        return max(bot_set)+1





