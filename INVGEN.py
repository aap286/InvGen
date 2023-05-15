from fundamentals import (
    dataTypeCor,
    createFrame,
    allowed_file,
    money,
    getInterest,
    dateFormat,
    isValid,
)
from flask import Flask, request, render_template
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
from email.message import EmailMessage
import os, win32com.client, winshell, win32com
import pandas as pd
import numpy as np
import warnings
import smtplib
import webview
import pdfkit
import math
import ssl


# suppresses warnings
warnings.simplefilter(action="ignore", category=FutureWarning)

# location of html to pdf converter
sysmConfig = pdfkit.configuration(
    wkhtmltopdf="style\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
)

# import configuration file
configFrame = pd.read_csv(
    "0_Configuration\\configuration.csv", header=None, names=["Name", "Value"]
).T

# renaming columns
configFrame.columns = configFrame.iloc[0]  # column names matches to the first row
configFrame = configFrame.drop(configFrame.index[0])  # removing the first row

# correcting dataType
dataTypeCor(configFrame, 0, 11, float)

# config data types
""" 
0 Maintenance Rate                    float64
1 Weighted Interest                   float64
2 Reserve Fund                        float64
3 Recovery Interest rate              float64
4 Non occupancy charges               float64
5 Payment Discount                    float64
6 Payment Discount Period             float64
7 Delayed Payment charges             float64
8 CGST Rate                           float64
9 SGST Rate                           float64
10 Delayed Payment charges Interest    float64
# NA Tax arrears                       object
# Deemed Conveyance arrears            object
# Painting Project arrears             object
# Lift Project arrears                 object
# Delay Payment Charges                object
11 emailPasscode                       object
12 email Address                       object
13 row 3                               object
"""

# converts vertical data frame into list which each element is a dictionary
configFrame = list(configFrame.to_dict().values())

# empty array to store configFrame values
config = []
for value in configFrame:
    config.append(value["Value"])

# # memory management
del configFrame

# # store row names
# itemName = [""] * 35
# itemNamesA = [""] * 2

# # send mail
email_password = config[11]

# # sender email address
sender = config[12]

# aptArea = config[13]

# flask web service
def create_app():
    app = Flask(__name__, template_folder="style\\templates", static_folder="style")
    window = webview.create_window("Invoice", app, width=850, height=750)

    # home page
    @app.route("/", methods=["GET", "POST"])
    def home():

        # create frame of fixed size
        # df = createFrame(49, 14, 47, float)

        # after submitting the form
        if request.method == "POST":
            file = request.files["file"]  # input excel file
            invoiceDate = request.form.get("invoiceDate")  # invoice creation date
            year = request.form.get("year")  # invoice year range
            period = request.form.get("period")  # invoice period name
            subject = request.form["subject"]
            body = request.form["message"]

            # TRUE when no inputs are empty
            if (
                invoiceDate != ""
                and year != ""
                and period != ""
                and file != ""
                and subject != ""
                and body != ""
                and allowed_file(file.filename)
            ):
                # accepts excel file
                filename = secure_filename(file.filename)  # removes / from file name
                filename = os.path.join("Input", filename)  # directory of input files
                file.save(filename)  # saves file in input folder

                # # directory of saved invoice pdfs
                # dirName = "{}\\0_Invoices\\{}".format(os.getcwd(), year)

                # # creates directory file for invoice pdf
                # try:
                #     os.makedirs(dirName)
                #     # creates shortcut
                #     desktop = winshell.desktop()
                #     path = os.path.join(desktop, "Invoices.lnk".format(os.getcwd()))
                #     target = "{}\\0_Invoices".format(os.getcwd())
                #     icon = "{}\\style\\icon\\invoice.ico".format(os.getcwd())
                #     shell = win32com.client.Dispatch("WScript.Shell")
                #     shortcut = shell.CreateShortCut(path)
                #     shortcut.Targetpath = target
                #     shortcut.IconLocation = icon
                #     shortcut.save()
                # except FileExistsError:
                #     None

                # import form excel file
                apartments = pd.read_excel(filename, "A5")

                # convert format of date
                # invoiceDate = datetime.strptime(invoiceDate, "%Y-%m-%d").date()

                # reinstate due date of invoice
                # invoiceDateDue = invoiceDate + timedelta(days=config[6])

                # # adding columns to summary excel file
                # itemNamesA[0] = "Current Period (FY{})".format(year)
                # itemName[1] = "Maintenance @ Rs.{}/sqft/mth".format(config[0])
                # itemName[2] = "Less: Interest credit on Corpus @ {}% p.a.".format(
                #     config[1]
                # )

                # if config[3] > 0:
                #     itemName[3] = "Add: {}".format(aptArea)
                # elif config[3] < 0:
                #     itemName[3] = "Less: {}".format(aptArea)
                # itemName[4] = "Net Maintenance Payable"
                # itemName[5] = "Reserve Fund @ Rs.{}/sqft/mth (excl GST)".format(
                #     config[2]
                # )
                # itemName[
                #     6
                # ] = "Non-Occupancy Charges @ {}% of item 1 (if rented out)".format(
                #     config[4]
                # )
                # itemName[7] = "CGST @ {}% (on items 5,6)".format(config[8])
                # itemName[8] = "CGST @ {}% (on items 5,6)".format(config[9])
                # itemName[9] = "Total Current Period Dues (A1)"
                # itemNamesA[1] = "Arrears from previous periods/invoices (if any)"
                # itemName[10] = "{}".format(list(apartments.columns)[7])
                # itemName[11] = "{}".format(list(apartments.columns)[8])
                # itemName[12] = "{}".format(list(apartments.columns)[9])
                # itemName[13] = "{}".format(list(apartments.columns)[10])
                # itemName[14] = "{}".format(list(apartments.columns)[11])
                # itemName[15] = "{}".format(list(apartments.columns)[12])
               
                # itemName[16] = "CGST @ {}% (on items 15)".format(config[8])
                # itemName[17] = "CGST @ {}% (on items 15)".format(config[9])
                # itemName[18] = "Total Arrears (A2)"
                # itemName[19] = "Grand Total Due (Payable upto {}) (A1+A2)".format(
                #     invoiceDateDue.strftime("%d-%b-%Y")
                # )
                # itemName[20] = "Full Year Amount"
                # itemName[21] = "Less: Discount @ {}% (on item 4)".format(config[5])
                # itemName[22] = "Net payable on or before {}".format(
                #     invoiceDateDue.strftime("%d-%b-%Y")
                # )

                # ? counts total aprtments send
                counter = 0
                #  ? track apartments that didnt recieve email 
                noEmail = []
                emailSent = []

                # read line by line for each user
                for j in range(len(apartments)):
                

                #     # extract first row from apartments
                    userOne = apartments.loc[j]

                    # output
                    """ 
                    0 S.NO                                                                       2
                    1 B.NO                                                                      A1
                    2 Flat No.                                                                 103
                    3 Name                                                                  MR TYC
                    4 Area, Sq.fit                                                            2045
                    5 Self-occupied                                                              N
                    6 Actual Corpus Deposit                                                 300006
                    7 Mntce Arrears from previous periods invoices (if any)                      0
                    8 NA Tax arrears                                                             0
                    9 Deemed Conveyance arrears                                                  0
                    10 Painting Project arrears                                                 NaN
                    11 Lift Project arrears                                                       0
                    12 Email ID                                                          @gmail.com
                    13 Phone No.                                                                  0
                    14 Invoice No.                                                       191/2023-24
                        """
               
                #     # store amount
                #     item = [0] * 36

                #     # store delayed payment for arears
                #     delayedItem = [0] * 11

                #     # store due payment grand total
                #     delayedTotal = [0] * 11

                #     # print(userOne)
                # name of user
                    name = userOne[3]

                    # apartment and building number
                    aptNo = "{a}/{b}".format(
                        a=userOne[1], b=userOne[2]
                    )  # apartment number concanted

                #     # status of occupancy
                #     status = userOne[5]

                #     # corpus value
                #     corpus = "{:,}".format(userOne[6])

                #     # Section A
                #     # maintenance @ 1.80/sqft/mth
                #     item[1] = round(config[0] * 12 * userOne[4], 2)

                #     # Interest credit
                #     item[2] = round(config[1] * userOne[6] / 100, 2)

                #     # Recovery excess interest rate
                #     item[3] = round(config[3] * userOne[6] / 100, 2)

                #     # Net Maintenance Payable
                #     item[4] = round(item[1] - item[2] + item[3], 2)

                #     # Reserve Fund
                #     item[5] = round(config[2] * 12 * userOne[4], 2)

                #     # Non-Occupancy Charges
                #     try:
                #         if status == "Y":
                #             status = "Rented"
                #             item[6] = round(config[4] * item[1] / 100, 2)
                #         else:
                #             status = "Self-Occupied"
                #     except:
                #         None

                #     item[7] = config[8] / 100 * (item[5] + item[6])
                #     item[8] = config[9] / 100 * (item[5] + item[6])
                #     item[9] = item[4] + item[5] + item[6] + item[7] + item[8]


                #     # sets till item 15
                #     for i in range(1, 7):
                #         if userOne[i + 6] != 0:
                #             item[i + 9] = userOne[i + 6]
                  

                #     item[16] = item[15] * config[8] / 100
                #     item[17] = item[15] * config[9] / 100
                #     item[18] = (
                #         item[10]
                #         + item[11]
                #         + item[12]
                #         + item[13]
                #         + item[14]
                #         + item[15]
                #         + item[16]
                #         + item[17]
                #     )
                #     item[19] = item[18] + item[9]

                    
                #     item[20] = item[19]

                #     item[21] = round(config[5] * item[4] / 100, 2)
                #     item[22] = round(item[20] - item[21], 2)
                #     item[23] = item[20]
                #     schpay = item[20]
                #     interest = config[7]

                #     # dates array
                #     dates = [0] * 12

                #     #  generates payment plan for due dates entire year
                #     getInterest(item, invoiceDateDue, schpay, interest, dates)

                #     #  reformat numbers to currency
                #     money(item)

                #     if config[3] < 0:
                #         item[3] = "({})".format(item[3])

                #     money(delayedItem)

                    #     # reformat date
                    # dateFormat(dates)
                    # itemName[23:35] = dates

                    #     # change format of invoice  date
                    # invoiceDateStr = invoiceDate.strftime("%d/%m/%Y")
                    # invoiceDateDueStr = invoiceDateDue.strftime("%d-%b-%Y")

                    #     # indexing invoice docs
                    # index = "{}/{}".format(str(101 + j), year)

                    # ! invoice No.
                    # invoiceNo = userOne["Invoice Number"]

                    # html = render_template(
                    #     "output.html",
                    #     item=item,
                    #     config=config,
                    #     itemName=itemName,
                    #     itemNamesA=itemNamesA,
                    #     delayedItem=delayedItem,
                    #     delayedTotal=delayedTotal,
                    #     name=name,
                    #     aptNo=aptNo,
                    #     aptArea=userOne[4],
                    #     status=status,
                    #     corpus=corpus,
                    #     index=index,
                    #     invoiceDate=invoiceDateStr,
                    #     invoiceDateDue=invoiceDateDueStr,
                    #     year=year,
                    #     period=period,
                    #     invoiceNo = invoiceNo
                    # )

                    #     # PDF options
                    options = {
                            "orientation": "portrait",
                            "page-size": "A4",
                            "margin-top": "1.0cm",
                            "margin-right": "1cm",
                            "margin-bottom": "1.0cm",
                            "margin-left": "1cm",
                            "encoding": "UTF-8",
                            "enable-local-file-access": "",
                        }

                        #     # pdf file name
                    aptNoSave = "{a} {b}".format(
                            # a=userOne["B.NO"], b=userOne["Flat No."]
                            a=userOne[1], b=userOne[2]
                        )
                # pdfkit.from_string(
                #         # html,
                #         "0_Invoices\\{}\\{}.pdf".format(year, aptNoSave),
                #         options=options,
                #         configuration=sysmConfig,
                #         css=["style\\css\\outputstyle.css"],
                #     )

                    #     # convert int to float
                    # for t in range(len(item)):
                    #     try:
                    #         item[t] = float(item[t])
                    #     except:
                    #         None

                    #     # inserting row to summary data Frame
                    # row = np.concatenate((userOne[1:15], item[1 : len(item)]))
                    # df = df.append(
                    #     pd.Series(row, index=df.columns[: len(row)]), ignore_index=True
                    # )

                    # email Object
                    msg = EmailMessage()
                    msg["From"] = sender
                        # msg["To"] = userOne[13]
                    # msg["To"] = userOne[4]
                    reciever = userOne[4]
                    msg["To"] = reciever
                    msg["Subject"] = subject

                   
                        # forming body of the email
                    salution = "Dear {},\n".format(name)
                    msg.set_content(salution + body)

                    try:
                        if reciever != "":
                                #print("Inside if condition")
                                # opens pdf file and attaches to draft
                                with open(
                                    "0_Invoices\\{}\\{}.pdf".format(year, aptNoSave), "rb"
                                ) as content_file:
                                    content = content_file.read()
                                    msg.add_attachment(
                                        content,
                                        maintype="application",
                                        subtype="pdf",
                                        filename="{}.pdf".format(aptNoSave),
                                    )

                                # creates a safe network
                                context = ssl.create_default_context()
                                with smtplib.SMTP_SSL(
                                    "smtp.gmail.com", 465, context=context
                                ) as smtp:
                                    smtp.login(sender, email_password)
                                    # smtp.sendmail(sender, userOne[13], msg.as_string())
                                    smtp.sendmail(sender, reciever, msg.as_string())

                                counter = counter +1
                                emailSent.append(aptNoSave)

                                print("Aprt No: {} sent to {}".format(aptNoSave, userOne[4]))
                    except:
                        
                        None

                    # user has email
                # try:
                #         # if userOne[13] != "" and isValid(userOne[13]):

                #             # opens pdf file and attaches to draft
                #             with open(
                #                 "0_Invoices\\{}\\{}.pdf".format(year, aptNoSave), "rb"
                #             ) as content_file:
                #                 content = content_file.read()
                #                 msg.add_attachment(
                #                     content,
                #                     maintype="application",
                #                     subtype="pdf",
                #                     filename="{}.pdf".format(aptNoSave),
                #                 )

                #             # creates a safe network
                #             context = ssl.create_default_context()
                #             with smtplib.SMTP_SSL(
                #                 "smtp.gmail.com", 465, context=context
                #             ) as smtp:
                #                 smtp.login(sender, email_password)
                #                 smtp.sendmail(sender, userOne[13], msg.as_string())
                # except:
                #         # None

                # drop first row
                # df.drop(index=df.index[0], axis=0, inplace=True)

                # # renaming columns
                # for i in range(14):
                #     df.rename(
                #         columns={df.columns[i]: apartments.columns[i + 1]}, inplace=True
                #     )
                # for k in range(34):
                #     df.rename(
                #         columns={df.columns[k + 14]: itemName[k + 1]}, inplace=True
                #     )
                # exporting summary data Frame
                # df.to_csv("0_Invoices\\{}\\1.csv".format(year))
                print(counter)
                #print(noEmail)
                print(emailSent)
                return render_template("progress.html")

        # template before running the file
        return render_template("home.html")

    if __name__ == "__main__":
        # app.run(debug=False)
        webview.start()


# runs service
create_app()
