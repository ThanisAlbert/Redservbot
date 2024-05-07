from django import template

from botcreate.models import TicketTrackingTable

register = template.Library()

@register.filter(name='replacenhancestart')
def replacenhancestart(value):
    if "1900" in str(value):
        return ""
    else:
        return value

@register.filter(name='replacetimesave')
def totaltimesave(value):
    if str(value) == "None":
        Totaltimesaved = ""
    elif ":" in str(value):
        print(str(value))
        Totaltimesaved = str(value)
        hour = Totaltimesaved.split(":")[0]
        minute = Totaltimesaved.split(":")[1]
        hourinmins = int(hour) * 60
        Totaltimesaved = str(hourinmins + int(minute)) + str(" mins")
        return Totaltimesaved
    else:
        Totaltimesaved = str(value)
        return Totaltimesaved


@register.filter(name='replacenhanceend')
def replacenhanceend(value):
    if "1900" in str(value):
        return ""
    else:
        return value


@register.filter(name='replac')
def replac(value):

    if "mins" in str(value) and "days" not in str(value):
        totaldaysaved = round(int(str(value).replace("mins", "")) / 60)
        if totaldaysaved <= 8:
            totaldaysaved = str(totaldaysaved) + " hrs"
        else:
            totaldaysaved = str(round(totaldaysaved / 8)) + " Business days"
    else:
        totaldaysaved = value

    return totaldaysaved

@register.filter(name='replaceslash')
def replaceslash(value):

    if value == "/":
        return " "
    else:
        return value

