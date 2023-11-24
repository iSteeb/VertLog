
import pandas as pd
import os
import matplotlib.pyplot as plt
import math
import datetime
import matplotlib.dates as mdates
import imaplib
import email
import numpy as np
import pick
import subprocess

DATA_DIRECTORY = "./data"
PLOTS_DIRECTORY = "./plots"

if not os.path.exists("config.py"):
    usernameInput = input("Enter your IMAP email address: ")
    passwordInput = input("Enter your IMAP password: ")
    receiverInput = input("Enter the email address you receive the data at: ")
    serverInput = input("Enter the IMAP server address: ")
    portInput = input("Enter the server port: ")
    SSLInput = input("Enter whether to use SSL (y/n): ")
    with open("config.py", "w") as f:
        f.write("email = \"" + usernameInput + "\"\n")
        f.write("password = \"" + passwordInput + "\"\n")
        f.write("receiver = \"" + receiverInput + "\"\n")
        f.write("server = \"" + serverInput + "\"\n")
        f.write("port = \"" + portInput + "\"\n")
        f.write("SSL = " + SSLInput + "\n")

import config

if config.SSL == "y" or config.SSL == "Y":
    M = imaplib.IMAP4_SSL(config.server, int(config.port))
else:
    M = imaplib.IMAP4(config.server, config.port)
M.login(config.email, config.password)

print("Logged into iCloud")
M.select()
typ, data = M.search(None, 'SUBJECT', '"Session Data to Open in Excel"', "TO", config.receiver)
if len(data) > 0 and data[0] != None:
    print("Found " + str(len(data[0].split())) + " emails with session data.")
    for mail_id in data[0].split():
        typ, data = M.fetch(mail_id, 'BODY[]')
        msg = email.message_from_bytes(data[0][1])   
        for part in msg.walk():
            if part.get_content_type() == "application/vnd.ms-excel":
                filename = part.get_filename()
                att_path = os.path.join(DATA_DIRECTORY, filename)
                print("Processing " + att_path + " from email.")
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
        dt = session["Datetime"][0]
        averageHigh = session['Jump Height (cm)'].nlargest(math.ceil(len(session) / 4)).median()
        highest = session["Jump Height (cm)"].max()
        x=session["Datetime"].loc[session["Jump Height (cm)"] == highest]
        lastDatePlot = session.plot.scatter(x="Datetime", y="Jump Height (cm)")
        lastDatePlot.set_title("Jump Height (cm) vs. Time on " + str(dt.date()) + " at " + str(dt.time().strftime("%H:%M")))
        lastDatePlot.set_xlabel("Time")
        lastDatePlot.set_ylabel("Jump Height (cm)" + " n = " + str(len(session)) + " x̄ = " + str(round(averageHigh, 1)))
        lastDatePlot.xaxis.set_major_locator(mdates.MinuteLocator(interval=15))
        lastDatePlot.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        # highlight the highest value in red and add a label with the number
        lastDatePlot.scatter(x=session["Datetime"].loc[session["Jump Height (cm)"] == highest].head(1), y=highest, color="red")
        lastDatePlot.annotate(str(highest), (session["Datetime"].loc[session["Jump Height (cm)"] == highest].head(1), highest))
        # add a horizontal line at the average high height
        lastDatePlot.axhline(y=averageHigh, color='r', linestyle='-')
        #  add a polynomial trendline with datetime as the x-axis and jump height as the y-axis
        z = np.polyfit(session["Datetime"].astype(np.int64) // 10**9, session["Jump Height (cm)"], 2)
        p = np.poly1d(z)
        lastDatePlot.plot(session["Datetime"],p(session["Datetime"].astype(np.int64) // 10**9),"r--")

        # save the plot to plots directory
        plotfile = str(dt.date()) + "T" + str(dt.time().strftime("%H-%M")) + ".png"
        plt.savefig(os.path.join(PLOTS_DIRECTORY, plotfile))

        collection = pd.concat([collection, session], ignore_index = True, axis = 0)
        
        processed.append(filename)
        with open("processed", "w") as f:
            f.write("\n".join(processed))
        
        print("... done. Saved plot to " + plotfile + " and added data to collection dataframe.")

collection.to_csv("collection.csv", index=False)
print("Saving dataframe and generating graphs...")

today = datetime.date.today()
startOfWeek = today - datetime.timedelta(days=today.weekday())
endOfWeek = today + datetime.timedelta(days=(6 - today.weekday()))
options = ["[THIS WEEK]", "[LAST WEEK]", "[LAST SIX WEEKS]", "[THIS SEASON]", "[CUSTOM DATE]"]
option, index = pick.pick(options, "Select a Period to Graph", indicator='=>', default_index=0)
startDate = ""
if option == "[CUSTOM DATE]":
    startDate = input("Enter a start date (YYYY-MM-DD): ")
    startDate = datetime.datetime.strptime(startDate, "%Y-%m-%d").date()
elif option == "[LAST WEEK]":
    startDate = startOfWeek - datetime.timedelta(weeks=1)
    endOfWeek = endOfWeek - datetime.timedelta(weeks=1)
elif option == "[LAST SIX WEEKS]":
    startDate = startOfWeek - datetime.timedelta(weeks=5)
elif option == "[THIS SEASON]":
    startDate = datetime.date(2023, 6, 1)
elif option == "[THIS WEEK]":
    startDate = startOfWeek

subCollection = collection[(collection["Datetime"].dt.date >= startDate) & (collection["Datetime"].dt.date <= endOfWeek)]
subCollection["Date"] = subCollection["Datetime"].dt.date
subAverageHigh = subCollection['Jump Height (cm)'].nlargest(math.ceil(len(subCollection) / 4)).median()
highest = subCollection["Jump Height (cm)"].max()

seasonPlot = subCollection.plot.scatter(x="Date", y="Jump Height (cm)")
seasonPlot.set_title("Jump Height (cm) vs. Date from " + str(startDate) + " to " + str(endOfWeek))
seasonPlot.set_xlabel("Date")
seasonPlot.set_ylabel("Jump Height (cm)" + " n = " + str(len(subCollection)) + " x̄ = " + str(round(subAverageHigh, 1)))
seasonPlot.set_xlim(startDate - datetime.timedelta(days=1), endOfWeek + datetime.timedelta(days=1))
seasonPlot.set_xticks(pd.date_range(start=startDate - datetime.timedelta(days=1), end=endOfWeek + datetime.timedelta(days=1), freq='M'))
seasonPlot.axhline(y=subAverageHigh, color='r', linestyle='-')

# for i in range((endOfWeek - startDate).days + 2):
#     date = startDate + datetime.timedelta(days=i)
#     dateCollection = subCollection[subCollection["Datetime"].dt.date == date]
#     highest = dateCollection["Jump Height (cm)"].max()
    
plt.show()

print("... done. Terminating.")


