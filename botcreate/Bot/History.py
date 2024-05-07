from botcreate.models import BotHist, Bot


class BotHistory:

    def __init__(self, Bot):
        self.Bot = Bot

    def save_history(self):

        bot_obj_list = Bot.objects.filter(Botno=self.Bot.Botno).values()

        for botobj in bot_obj_list:
            bot = BotHist()
            bot.botno = botobj['Botno']
            bot.botname = botobj["Botname"]
            bot.botstatus = botobj["Botstatus"]
            bot.creationdate=botobj["Creationdate"]
            bot.startdate = botobj["Startdate"]
            bot.enddate = botobj["Enddate"]
            bot.enhancestartdate = botobj["enhancestartdate"]
            bot.enhanceenddate = botobj["enhanceenddate"]
            bot.livestatus = botobj["livestatus"]
            bot.remarks = botobj["remarks"]
            bot.Developermail = botobj["Developermail"]
            bot.save()

