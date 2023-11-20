
import pandas as pd
import os
import matplotlib.pyplot as plt
import math
import datetime
import matplotlib.dates as mdates
import imaplib
import email
import numpy as np
import subprocess

DATA_DIRECTORY = "./data"
PLOTS_DIRECTORY = "./plots"

if not os.path.exists("config.py"):
    usernameInput = input("Enter your iCloud email address: ")
    passwordInput = input("Enter your iCloud password: ")
    receiverInput = input("Enter the email address you receive the data at: ")
    with open("config.py", "w") as f:
        f.write("email = \"" + usernameInput + "\"\n")
        f.write("password = \"" + passwordInput + "\"\n")
        f.write("receiver = \"" + receiverInput + "\"\n")

from config import email, password, receiver

M = imaplib.IMAP4_SSL('imap.mail.me.com', 993)
M.login(email, password)
print("Logged into iCloud")
M.select()
typ, data = M.search(None, 'SUBJECT', '"Session Data to Open in Excel"', "TO", receiver)
if len(data) > 0 and data[0] != None:
  for mail_id in data[0].split():
      typ, data = M.fetch(mail_id, '(BODY[])')
      msg = email.message_from_bytes(data[0][1])
      for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
          continue
        if part.get('Content-Disposition') is None:
          continue

        filename = part.get_filename()
        print(filename)
        att_path = os.path.join(DATA_DIRECTORY, filename)
        print("Processing... " + att_path + " from email.")
        if not os.path.isfile(att_path):
          fp = open(att_path, 'wb')
          fp.write(part.get_payload(decode=True))
          fp.close()

      M.store(mail_id, '+FLAGS', '\\Deleted')

M.expunge()
M.close()
M.logout()

processed = []
if not os.path.exists("processeed"):
    open("processed", 'a').close()

with open("processed", "r") as f:
    processed = f.read().splitlines()

collection = pd.DataFrame()
if os.path.exists("collection.csv"):
    collection = pd.read_csv("collection.csv")
    collection["Datetime"] = pd.to_datetime(collection["Timestamp"])
else:
    processed = []

for filename in os.listdir(DATA_DIRECTORY):
    if filename.endswith('.xls') and filename not in processed:
        path = os.path.join(DATA_DIRECTORY, filename)
        print("Processing... " + path)
        session = pd.DataFrame()
        session["Timestamp"] = pd.read_xml(path, xpath="//ss:Worksheet[@ss:Name='Jumps']/ss:Table/ss:Row/*[@ss:StyleID='s2083']/*", namespaces={"ss": "urn:schemas-microsoft-com:office:spreadsheet"}).iloc[:, 1]
        session["Datetime"] = pd.to_datetime(session["Timestamp"])
        session["Jump Height (cm)"] = pd.read_xml(path, xpath="//ss:Worksheet[@ss:Name='Jumps']/ss:Table/ss:Row/*/*[@ss:Type='Number']", namespaces={"ss": "urn:schemas-microsoft-com:office:spreadsheet"}).iloc[:, 1]
        
        # plot session
        date = session["Datetime"][0].date()
        averageHigh = session['Jump Height (cm)'].nlargest(math.ceil(len(session) / 4)).median()
        highest = session["Jump Height (cm)"].max()
        x=session["Datetime"].loc[session["Jump Height (cm)"] == highest]
        lastDatePlot = session.plot.scatter(x="Datetime", y="Jump Height (cm)")
        lastDatePlot.set_title("Jump Height (cm) vs. Time on " + str(date))
        lastDatePlot.set_xlabel("Time")
        lastDatePlot.set_ylabel("Jump Height (cm)")
        lastDatePlot.xaxis.set_major_locator(mdates.MinuteLocator(interval=15))
        lastDatePlot.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        # highlight the highest value in red and add a label with the number
        lastDatePlot.scatter(x=session["Datetime"].loc[session["Jump Height (cm)"] == highest].head(1), y=highest, color="red")
        lastDatePlot.annotate(str(highest), (session["Datetime"].loc[session["Jump Height (cm)"] == highest].head(1), highest))
        # add a horizontal line at the average height and label it
        lastDatePlot.axhline(y=averageHigh, color='r', linestyle='-')
        # add a label with the count and average
        lastDatePlot.annotate("n = " + str(len(session)) + " x̄ = " + str(round(averageHigh, 1)), (session["Datetime"].head(1), highest - 1))
        #  add a polynomial trendline with datetime as the x-axis and jump height as the y-axis
        z = np.polyfit(session["Datetime"].astype(np.int64) // 10**9, session["Jump Height (cm)"], 2)
        p = np.poly1d(z)
        lastDatePlot.plot(session["Datetime"],p(session["Datetime"].astype(np.int64) // 10**9),"r--")

        # save the plot to daily_plots directory
        plotfile = str(date) + "_AvgHigh_" + str(averageHigh) + "_Count_" + str(len(session)) + ".png"
        plt.savefig(os.path.join(PLOTS_DIRECTORY, plotfile))
        plt.close()
        subprocess.call(['open', os.path.join(PLOTS_DIRECTORY, plotfile)])

        collection = pd.concat([collection, session], ignore_index = True, axis = 0)
        processed.append(filename)
        with open("processed", "w") as f:
            f.write("\n".join(processed))
        print("... done.")

collection.to_csv("collection.csv", index=False)
print("Saved data to collection.csv and generating graphs...")


# filter collection to only include rows within latest calendar week and then plot
today = datetime.date.today()
startOfWeek = today - datetime.timedelta(days=today.weekday())
endOfWeek = startOfWeek + datetime.timedelta(days=6)
lastWeekCollection = collection[(collection["Datetime"].dt.date >= startOfWeek) & (collection["Datetime"].dt.date <= endOfWeek)]
lastWeekCollection["Date"] = lastWeekCollection["Datetime"].dt.date
lastWeekPlot = lastWeekCollection.plot.scatter(x="Date", y="Jump Height (cm)")
lastWeekPlot.set_title("Jump Height (cm) vs. Date for Previous Week")
lastWeekPlot.set_xlabel("Weekday")
lastWeekPlot.set_ylabel("Jump Height (cm)")
lastWeekPlot.set_xlim(startOfWeek, endOfWeek)
weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
lastWeekPlot.set_xticks(pd.date_range(start=startOfWeek, end=endOfWeek, freq='D'))
lastWeekPlot.set_xticklabels(weekdays)

# add a line for the weekly average high
averageHigh = lastWeekCollection['Jump Height (cm)'].nlargest(math.ceil(len(lastWeekCollection) / 4)).median()
highest = lastWeekCollection["Jump Height (cm)"].max()
lastWeekPlot.axhline(y=averageHigh, color='r', linestyle='-')
lastWeekPlot.annotate("n = " + str(len(lastWeekCollection)) + " x̄ = " + str(round(averageHigh, 1)), (lastWeekCollection["Datetime"].head(1), highest - 1))

for i in range(7):
    date = startOfWeek + datetime.timedelta(days=i)
    dateCollection = lastWeekCollection[lastWeekCollection["Datetime"].dt.date == date]
    averageHigh = dateCollection['Jump Height (cm)'].nlargest(math.ceil(len(dateCollection) / 4)).median()
    lastWeekPlot.scatter(x=date, y=averageHigh, color='r', linestyle='-')
    lastWeekPlot.annotate(str(round(averageHigh, 1)), (date, averageHigh))
    highest = dateCollection["Jump Height (cm)"].max()
    lastWeekPlot.scatter(x=date, y=highest, color="red")
    lastWeekPlot.annotate(str(highest), (date, highest))



# filter collection to only include rows within latest 6 calendar weeks and then plot
sixWeeksAgo = startOfWeek - datetime.timedelta(weeks=5)
lastSixWeeksCollection = collection[(collection["Datetime"].dt.date >= sixWeeksAgo) & (collection["Datetime"].dt.date <= endOfWeek)]
lastSixWeeksCollection["Date"] = lastSixWeeksCollection["Datetime"].dt.date
lastSixWeeksPlot = lastSixWeeksCollection.plot.scatter(x="Date", y="Jump Height (cm)")
lastSixWeeksPlot.set_title("Jump Height (cm) vs. Date for Previous 6 Weeks")
lastSixWeeksPlot.set_xlabel("Day of Week")
lastSixWeeksPlot.set_ylabel("Jump Height (cm)")
lastSixWeeksPlot.set_xlim(sixWeeksAgo, endOfWeek)
lastSixWeeksPlot.set_xticks(pd.date_range(start=sixWeeksAgo, end=endOfWeek, freq='W'))
prevLabels = ['-5', '-4', '-3', '-1', 'Mon', 'EOW']
lastSixWeeksPlot.set_xticklabels(prevLabels)

averageHigh = lastSixWeeksCollection['Jump Height (cm)'].nlargest(math.ceil(len(lastSixWeeksCollection) / 4)).median()
highest = lastSixWeeksCollection["Jump Height (cm)"].max()
lastSixWeeksPlot.axhline(y=averageHigh, color='r', linestyle='-')
lastSixWeeksPlot.annotate("n = " + str(len(lastSixWeeksCollection)) + " x̄ = " + str(round(averageHigh, 1)), (lastSixWeeksCollection["Datetime"].head(1), highest - 1))

for i in range(42):
    date = sixWeeksAgo + datetime.timedelta(days=i)
    dateCollection = lastSixWeeksCollection[lastSixWeeksCollection["Datetime"].dt.date == date]
    averageHigh = dateCollection['Jump Height (cm)'].nlargest(math.ceil(len(dateCollection) / 4)).median()
    lastSixWeeksPlot.scatter(x=date, y=averageHigh, color='r', linestyle='-')
    highest = dateCollection["Jump Height (cm)"].max()
    lastSixWeeksPlot.scatter(x=date, y=highest, color="red")

# filter collection to only include rows from the current season (since 9 aug 2023 to 30 apr 2024) and then plotä
startOfSeason = datetime.date(2023, 8, 9)
endOfSeason = datetime.date(2024, 4, 30)
seasonCollection = collection[(collection["Datetime"].dt.date >= startOfSeason) & (collection["Datetime"].dt.date <= endOfSeason)]
seasonCollection["Date"] = seasonCollection["Datetime"].dt.date
seasonPlot = seasonCollection.plot.scatter(x="Date", y="Jump Height (cm)")
seasonPlot.set_title("Jump Height (cm) vs. Date for SVW Season")
seasonPlot.set_xlabel("Date")
seasonPlot.set_ylabel("Jump Height (cm)")
seasonPlot.set_xlim(startOfSeason, endOfSeason)
seasonPlot.set_xticks(pd.date_range(start=startOfSeason, end=endOfSeason, freq='M'))
seasonLabels = ['Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May']
seasonPlot.set_xticklabels(seasonLabels)

averageHigh = seasonCollection['Jump Height (cm)'].nlargest(math.ceil(len(seasonCollection) / 4)).median()
highest = seasonCollection["Jump Height (cm)"].max()
seasonPlot.axhline(y=averageHigh, color='r', linestyle='-')
seasonPlot.annotate("n = " + str(len(seasonCollection)) + " x̄ = " + str(round(averageHigh, 1)), (seasonCollection["Datetime"].head(1), highest - 1))

for i in range((endOfSeason - startOfSeason).days):
    date = startOfSeason + datetime.timedelta(days=i)
    dateCollection = seasonCollection[seasonCollection["Datetime"].dt.date == date]
    averageHigh = dateCollection['Jump Height (cm)'].nlargest(math.ceil(len(dateCollection) / 4)).median()
    seasonPlot.scatter(x=date, y=averageHigh, color='r', linestyle='-')
    highest = dateCollection["Jump Height (cm)"].max()
    seasonPlot.scatter(x=date, y=highest, color="red")
    
plt.show()