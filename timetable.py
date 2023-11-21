import os
import excel2img
import win32com.client
import datetime as dt
import pandas as pd
from PIL import Image

day = str(dt.date.today()).split("-")[2]
month = str(dt.date.today()).split("-")[1]
year = str(dt.date.today()).split("-")[0][2:]

outlook = win32com.client.Dispatch("outlook.application").GetNamespace("MAPI")

inbox = outlook.Folders("jjeffery17@heathfieldcc.co.uk").Folders("Inbox")
messages = inbox.Items

for msg in messages:
    try:
        if ("Room Changes" in msg.Subject) and ("{0}.{1}.{2}".format(day, month, year) in msg.Attachments[0].FileName):
            attachment = msg.Attachments[0]
            print(msg.Subject)
            break
    except IndexError:
        pass

if attachment == None:
    raise("Room Changes Not Found")

attachment.SaveAsFile(os.getcwd()+"\\"+attachment.FileName)

changes = pd.read_excel(os.getcwd()+"\\"+attachment.FileName)
changes = changes.values

relevant = []

for row in changes:
    if row[2] == "13DUD/Tu":
        relevant.append(row.tolist())
    elif row[2] == "13D/Dm":
        relevant.append(row.tolist())
    elif row[2] == "13E/Fm":
        relevant.append(row.tolist())
    elif row[2] == "13A/Cp":
        relevant.append(row.tolist())
    elif row[2] == "13C/Se1":
        relevant.append(row.tolist())

print(relevant)

periods = []
data = []
for row in relevant:
    if row[0][-2] != "M":
        periods.append(row[0][-2])
    else:
        periods.append("AM")
    data.append([row[2], row[1], row[4], row[3]])
print(periods)
print(data)


data = pd.DataFrame(data,
                    periods,
                    ["Lesson", "Initial", "Replace", "Staff"])
data.to_excel("out.xlsx")

excel2img.export_img("out.xlsx", "changes.png")

i1 = Image.open(r"TimetableBG.png")
i2 = Image.open(r"changes.png")
i2Size = i2.size
i2 = i2.resize((250, i2Size[1]))

i1.paste(i2, (0, 0))

i1.save("Final.png")

os.system("git add Final.png")
os.system('git commit -m "Room Changes {0} {1} {2}"'.format(day, month, year))
os.system("git push")
