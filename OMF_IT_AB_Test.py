import os
import re
from validate_email import validate_email
import pandas as pd
from pandas import DataFrame, read_csv, merge
import numpy as np
import phonenumbers
import datetime
from datetime import date
from dateutil.relativedelta import relativedelta
import win32com.client as win32
import shutil, sys
import urllib
import sqlalchemy
import random
#import numpy as np
#from numpy import arange

# Email Notification
def vba_email(recipients, subject, message, cc=[], attachments=[]):
    """ list of str, str, str, list of str, list of str -> None
    Sends an email using the existing oulook profile with the following params
    > Email is sent to the list of recipeints
    > Subject is the subject line of the message
    > Message is the HTML formatted body of the message
    > CC is the list of emails to include on the CC line
    > Attachments are the file paths to the attchments for the message
    """
    # Initialize the outlook object within python
    outlook = win32.GetObject(Class="Outlook.Application")
    # Create the mail, populate the mail item
    mailer = outlook.CreateItem(0)
    mailer.To = ";".join(recipients)
    if cc:
        mailer.CC = ";".join(cc)
    mailer.Subject = subject
    mailer.HTMLBody = message
    # Add attachments if they exist, exit if attachment can't be found
    for item in attachments:
        if os.path.exists(item):
            mailer.Attachments.Add(item)
        else:
            raise FileNotFoundError(
                "The following attachment could not be found: {}".format(item)
            )
    # Send and collect garbage
    mailer.Send()
    del mailer
    del outlook


run_date = str(date.today())


try:
    # define city categories
    twocity = [
        "Bedford, OH",
        "Charlotte, NC",
        "Hickory, NC",
        "Chicago, IL",
        "Houston, TX",
        "Little Rock, AR",
        "Philadelphia, PA",
        "Pittsburgh, PA",
        "Waukesha, WI",
        "Harlingen, TX",
        "Milwaukee, WI",
        "Saint Louis, MO",
        "Shreveport, LA",
        "Syracuse, NY",
        "Toledo, OH"
    ]
    len(twocity)
    frenchcity = ["Boisbriand, QC", "Gatineau, QC", "Laval, QC"]
    len(frenchcity)

    removelist = twocity + frenchcity
    len(removelist)

    # Getting file path

    file = "\\\\adfs01.uhi.amerco\\interdepartment\\storagegroup\\Digital Marketing\\Emails\\One Month Free Email Campaign\\Documentation\\OMF_In_Town.csv"

    # read csv file
    df1 = pd.read_csv(file)
    N_file=len(df1.index)

    # remove create date 3 months ago
    three_months = date.today() + relativedelta(months=-3)
    three_months
    df1["Create Date"] = pd.to_datetime(df1["Create Date"])
    df1["Create Date"] = df1["Create Date"].dt.date
    df1 = df1[df1["Create Date"] > three_months]
    N_after_removed_createdate=len(df1.index)
    N_createdate_removed=N_file-len(df1.index)

    # remove duplicate based on email address
    df1.sort_values("Email Address", inplace=True)
    df1["Email Address"] = df1["Email Address"].str.title()

    df1.drop_duplicates(subset="Email Address", keep="first", inplace=True)
    N_Duplicates_removed=len(df1.index)
    N_Duplicates=N_after_removed_createdate-N_Duplicates_removed


    # remove wired or invalid email address
    df1 = df1.astype({"Email Address": str})
    df1["is_valid_email"] = df1["Email Address"].apply(lambda x: validate_email(x))
    df1 = df1[df1["is_valid_email"]]
    N_Invalidemail_removed=N_Duplicates_removed-len(df1.index)

    # proper formate customers' first name and last name
    df1["Customer’s First Name"] = df1["Customer’s First Name"].str.title()
    df1["Customer’s Last Name"] = df1["Customer’s Last Name"].str.title()
    len(df1.index)

    # vlook up to match up with equipment type
    df2 = pd.read_csv(
        "\\\\adfs01.uhi.amerco\\interdepartment\\storagegroup\\Digital Marketing\\Emails\\One Month Free Email Campaign\\Documentation\\EquipmentCode.csv"
    )
    len(df2.index)

    df3 = df1.merge(df2, on="Equipment Model", how="left")
    len(df3.index)
    df3.to_csv = "it.csv"

    # remove rows with NA Equipment Type
    df3 = df3.dropna(subset=["Equipment Type"])
    N_all_equipment=len(df3.index)

    # remove In-Town U-Box customers
    df3 = df3[~df3["Equipment Type"].isin(["U-Box"])]
    N_UBox=N_all_equipment-len(df3.index)

    # create city.st for vlookup
    df3["City of Destination"] = df3["City of Destination"].str.title()
    df3["City, ST"] = df3["City of Destination"] + ", " + df3["State of Destination"]
    len(df3.index)

    df4 = pd.read_excel(
        "\\\\adfs01.uhi.amerco\\interdepartment\\storagegroup\\Digital Marketing\\Emails\\One Month Free Email Campaign\\Entity Information Database - IT.xlsx",
        sheet_name="In-Town",
    )

    df5 = df3.merge(df4, on="City, ST", how="left")
    df5.drop_duplicates(subset="Email Address", keep="first", inplace=True)
    len(df5.index)
    df5.columns
    df5["City, ST_Entity"] = df5["City"] + ", " + df5["State"]

    list(df5.columns)
    df6 = df5.dropna(subset=["Entity"])
    len(df6.index)
    list(df6.columns)

    # format the Zipcode
    df6 = df6.astype({"Zip": str})

    # format phone number

    df6 = df6.astype({"Phone#": str})

    df6["Phone#"] = df6["Phone#"].apply(
        lambda x: "(" + x[:3] + ") " + x[3:6] + "-" + x[6:10]
    )
    df6.columns

    # drop extra columns
    cols_to_drop = [
        "Create Date",
        "Pickup Date",
        "Expected Arrival Date",
        "is_valid_email",
        "State of Origin",
        "City of Origin",
        "State of Destination",
        "Equipment Model",
        "In Town or One Way",
        "Source of Reservation",
        "City",
        "State",
        "Entity",
    ]

    df6.drop(cols_to_drop, axis=1, inplace=True)
    # df6 = df6.reindex(
    #     columns=[
    #         "Email Address",
    #         "Customer’s First Name",
    #         "Customer’s Last Name",
    #         "City of Destination",
    #         "Equipment Type",
    #         "Name",
    #         "Address",
    #         "City, ST",
    #         "Zip",
    #         "Phone#",
    #         "GM Email",
    #         "Website",
    #     ]
    # )
    df6.sort_values("City of Destination", inplace=True)

    # seperate two city list
    df2city = df6[df6["City, ST"].isin(twocity)]
    len(df2city.index)

    # Vlookup for twoecity list
    df2city_database = pd.read_excel(
        "\\\\adfs01.uhi.amerco\\interdepartment\\storagegroup\\Digital Marketing\\Emails\\One Month Free Email Campaign\\Entity Information Database - IT - 2 Cities.xlsx"
    )
    df2city_final = df2city.merge(df2city_database, on="City, ST", how="left")
    df2city_final.columns
    df2city_final = df2city_final.drop(
        ["Name", "Address", "City, ST", "Zip", "Phone#", "GM Email", "Website", "City, ST_Entity"], axis=1
    )
    df2city_final.columns = (
        "Email",
        "FirstName",
        "LastName",
        "City_of_Destination",
        "Equipment",
        "Entity_Name",
        "Entity_Address",
        "City_ST",
        "Zipcode",
        "entityphonenumber",
        "GM_Email",
        "Entityurl",
        "Entity_Name2",
        "Entity_Address2",
        "City_ST2",
        "Zipcode2",
        "entityphonenumber2",
        "GM_Email2",
        "Entityurl2",
        "resultslink",
    )

    # Rename the column title
    df6.drop(["City, ST"], axis=1, inplace=True)
    df6 = df6.reindex(
        columns=[
            "Email Address",
            "Customer’s First Name",
            "Customer’s Last Name",
            "City of Destination",
            "Equipment Type",
            "Name",
            "Address",
            "City, ST_Entity",
            "Zip",
            "Phone#",
            "GM Email",
            "Website",
        ]
    )

    df6.columns = [
        "Email",
        "FirstName",
        "LastName",
        "City_of_Destination",
        "Equipment",
        "Entity_Name",
        "Entity_Address",
        "City_ST",
        "Zipcode",
        "entityphonenumber",
        "GM_Email",
        "url",
    ]

    # seperate frenchcity list
    dffrenchcity = df6[df6["City_ST"].isin(frenchcity)]
    len(dffrenchcity.index)

    # remove all cities from IT list
    df_final = df6[~df6["City_ST"].isin(removelist)]

    #seperate lists for OMF-IT for test purpose
    df_final_control=df_final.sample(frac=1/2)
    df_final_testB=df_final.drop(df_final_control.index)
    # df_final_testA=df_final_test.sample(frac=1/2)
    # df_final_testB=df_final_test.drop(df_final_testA.index)
    # df_final_testB=df_final.sample(frac=1/3)

    #Add URL tracking for IT_Control
    book_online_link="?utm_campaign=storage&utm_source=omfIT&utm_medium=book_online_link&utm_content="

    entity_name_link="?utm_campaign=storage&utm_source=omfIT&utm_medium=entity_name_link&utm_content="

    book_storage_button="?utm_campaign=storage&utm_source=omfIT&utm_medium=book_storage_button&utm_content="

    hero="?utm_campaign=storage&utm_source=omfIT&utm_medium=hero&utm_content="

    today=date.today().strftime("%Y%m%d")

    for i, row in df_final_control.iterrows():
        df_final_control.at[i,"senddate"]=today

    for i, row in df_final_control.iterrows():
        df_final_control.at[i,"url2"]=df_final_control.at[i,"url"]+book_online_link+today

    for i, row in df_final_control.iterrows():
        df_final_control.at[i,"url3"]=df_final_control.at[i,"url"]+entity_name_link+today

    for i, row in df_final_control.iterrows():
        df_final_control.at[i,"url4"]=df_final_control.at[i,"url"]+book_storage_button+today

    for i, row in df_final_control.iterrows():
        df_final_control.at[i,"url5"]=df_final_control.at[i,"url"]+hero+today

    df_final_control.drop(["url"], axis=1, inplace=True)
    df_final_control.columns
    df_final_control.columns = ['Email', 'FirstName', 'LastName','City_of_Destination', 'Equipment','Entity_Name', 'Entity_Address', 'City_ST', 'Zipcode', 'entityphonenumber', 'GM_Email', 'senddate', 'url', 'url2','url3', 'url4']

    # #Add URL tracking for IT-df_final_testA
    # book_online_link="?utm_campaign=storage&utm_source=omfIT&utm_medium=book_online_link_A&utm_content="
    #
    # entity_name_link="?utm_campaign=storage&utm_source=omfIT&utm_medium=entity_name_link_A&utm_content="
    #
    # book_storage_button="?utm_campaign=storage&utm_source=omfIT&utm_medium=book_storage_button_A&utm_content="
    #
    # hero="?utm_campaign=storage&utm_source=omfIT&utm_medium=hero_A&utm_content="
    #
    # today=date.today().strftime("%Y%m%d")
    #
    # for i, row in df_final_testA.iterrows():
    #     df_final_testA.at[i,"senddate"]=today
    #
    # for i, row in df_final_testA.iterrows():
    #     df_final_testA.at[i,"url2"]=df_final_testA.at[i,"url"]+book_online_link+today
    #
    # for i, row in df_final_testA.iterrows():
    #     df_final_testA.at[i,"url3"]=df_final_testA.at[i,"url"]+entity_name_link+today
    #
    # for i, row in df_final_testA.iterrows():
    #     df_final_testA.at[i,"url4"]=df_final_testA.at[i,"url"]+book_storage_button+today
    #
    # for i, row in df_final_testA.iterrows():
    #     df_final_testA.at[i,"url5"]=df_final_testA.at[i,"url"]+hero+today
    #
    # df_final_testA.drop(["url"], axis=1, inplace=True)
    # df_final_testA.columns
    # df_final_testA.columns = ['Email', 'FirstName', 'LastName','City_of_Destination', 'Equipment','Entity_Name', 'Entity_Address', 'City_ST', 'Zipcode', 'entityphonenumber', 'GM_Email', 'senddate', 'url', 'url2','url3', 'url4']

    #Add URL tracking for IT - TestB

    book_online_link="?utm_campaign=storage&utm_source=omfIT&utm_medium=book_online_link_B&utm_content="

    entity_name_link="?utm_campaign=storage&utm_source=omfIT&utm_medium=entity_name_link_B&utm_content="

    book_storage_button="?utm_campaign=storage&utm_source=omfIT&utm_medium=book_storage_button_B&utm_content="

    hero="?utm_campaign=storage&utm_source=omfIT&utm_medium=hero_B&utm_content="

    today=date.today().strftime("%Y%m%d")

    for i, row in df_final_testB.iterrows():
        df_final_testB.at[i,"senddate"]=today

    for i, row in df_final_testB.iterrows():
        df_final_testB.at[i,"url2"]=df_final_testB.at[i,"url"]+book_online_link+today

    for i, row in df_final_testB.iterrows():
        df_final_testB.at[i,"url3"]=df_final_testB.at[i,"url"]+entity_name_link+today

    for i, row in df_final_testB.iterrows():
        df_final_testB.at[i,"url4"]=df_final_testB.at[i,"url"]+book_storage_button+today

    for i, row in df_final_testB.iterrows():
        df_final_testB.at[i,"url5"]=df_final_testB.at[i,"url"]+hero+today

    df_final_testB.drop(["url"], axis=1, inplace=True)
    df_final_testB.columns
    df_final_testB.columns = ['Email', 'FirstName', 'LastName','City_of_Destination', 'Equipment','Entity_Name', 'Entity_Address', 'City_ST', 'Zipcode', 'entityphonenumber', 'GM_Email', 'senddate', 'url', 'url2','url3', 'url4']


    #Add URL tracking for IT - French

    book_online_link_French="?utm_campaign=storage&utm_source=omfITFrench&utm_medium=book_online_link&utm_content="

    entity_name_link_French="?utm_campaign=storage&utm_source=omfITFrench&utm_medium=entity_name_link&utm_content="

    book_storage_button_French="?utm_campaign=storage&utm_source=omfITFrench&utm_medium=book_storage_button&utm_content="

    hero_French="?utm_campaign=storage&utm_source=omfITFrench&utm_medium=hero&utm_content="

    for i, row in dffrenchcity.iterrows():
        dffrenchcity.at[i,"senddate"]=today

    for i, row in dffrenchcity.iterrows():
        dffrenchcity.at[i,"url2"]=dffrenchcity.at[i,"url"]+book_online_link_French+today

    for i, row in dffrenchcity.iterrows():
        dffrenchcity.at[i,"url3"]=dffrenchcity.at[i,"url"]+entity_name_link_French+today

    for i, row in dffrenchcity.iterrows():
        dffrenchcity.at[i,"url4"]=dffrenchcity.at[i,"url"]+book_storage_button_French+today

    for i, row in dffrenchcity.iterrows():
        dffrenchcity.at[i,"url5"]=dffrenchcity.at[i,"url"]+hero_French+today

    dffrenchcity.drop(["url"], axis=1, inplace=True)
    dffrenchcity.columns
    dffrenchcity.columns = ['Email', 'FirstName', 'LastName','City_of_Destination', 'Equipment','Entity_Name', 'Entity_Address', 'City_ST', 'Zipcode', 'entityphonenumber', 'GM_Email', 'senddate', 'url', 'url2','url3', 'url4']

    #Add URL tracking for IT - 2 cities
    for i, row in df2city_final.iterrows():
        df2city_final.at[i,"Entityurl_NEW"]=df2city_final.at[i,"Entityurl"]+entity_name_link+today

    for i, row in df2city_final.iterrows():
        df2city_final.at[i,"Entityurl2_NEW"]=df2city_final.at[i,"Entityurl2"]+entity_name_link+today

    for i, row in df2city_final.iterrows():
        df2city_final.at[i,"senddate"]=today

    for i, row in df2city_final.iterrows():
        df2city_final.at[i,"url"]=df2city_final.at[i,"resultslink"]+book_online_link+today

    for i, row in df2city_final.iterrows():
        df2city_final.at[i,"url2"]=df2city_final.at[i,"resultslink"]+book_storage_button+today

    for i, row in df2city_final.iterrows():
        df2city_final.at[i,"url3"]=df2city_final.at[i,"resultslink"]+hero+today

    df2city_final.drop(["Entityurl","Entityurl2"], axis=1, inplace=True)
    df2city_final.columns

    df2city_final = df2city_final.reindex( columns=['Email', 'FirstName', 'LastName', 'City_of_Destination', 'Equipment',
       'Entity_Name', 'Entity_Address', 'City_ST', 'Zipcode',
       'entityphonenumber', 'GM_Email', 'Entityurl_NEW', 'Entity_Name2', 'Entity_Address2', 'City_ST2', 'Zipcode2', 'entityphonenumber2', 'GM_Email2', 'Entityurl2_NEW', 'resultslink' , 'senddate', 'url',
       'url2', 'url3'])

    df2city_final.columns=['Email', 'FirstName', 'LastName', 'City_of_Destination', 'Equipment',
       'Entity_Name', 'Entity_Address', 'City_ST', 'Zipcode',
       'entityphonenumber', 'GM_Email', 'Entityurl', 'Entity_Name2', 'Entity_Address2', 'City_ST2', 'Zipcode2', 'entityphonenumber2', 'GM_Email2', 'Entityurl2', 'resultslink' , 'senddate', 'url',
       'url2', 'url3']


    # drop the excel files to OMF folder
    #Control
    storage_dir_IT_control = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/storagegroup/Digital Marketing/Emails/One Month Free Email Campaign/Email Lists/{} - OMF - IT - Control.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    DMWA_dir_IT_control = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/DMWA/Storage/Emails/One Month Free Email Campaign/{} - OMF - IT - Control.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    df_final_control.to_excel(storage_dir_IT_control, index=False)
    shutil.copy2(storage_dir_IT_control, DMWA_dir_IT_control)

    # #testA
    # storage_dir_IT_testA = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/storagegroup/Digital Marketing/Emails/One Month Free Email Campaign/Email Lists/{} - OMF - IT - TestA.xlsx".format(
    #     datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    # )
    #
    # DMWA_dir_IT_testA = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/DMWA/Storage/Emails/One Month Free Email Campaign/{} - OMF - IT - TestA.xlsx".format(
    #     datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    # )
    #
    # df_final_testA.to_excel(storage_dir_IT_testA, index=False)
    # shutil.copy2(storage_dir_IT_testA, DMWA_dir_IT_testA)

    #testB
    storage_dir_IT_testB = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/storagegroup/Digital Marketing/Emails/One Month Free Email Campaign/Email Lists/{} - OMF - IT - TestB.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    DMWA_dir_IT_testB = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/DMWA/Storage/Emails/One Month Free Email Campaign/{} - OMF - IT - TestB.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    df_final_testB.to_excel(storage_dir_IT_testB, index=False)
    shutil.copy2(storage_dir_IT_testB, DMWA_dir_IT_testB)

    #Copy to the DWMA folder
    #shutil.copy2(storage_dir_IT, DMWA_dir_IT)

    N_IT_final=len(df_final_control.index)+len(df_final_testB.index)

    storage_dir_IT_2city = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/storagegroup/Digital Marketing/Emails/One Month Free Email Campaign/Email Lists/{} - OMF - IT - 2 Cities.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    DMWA_dir_IT_2city = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/DMWA/Storage/Emails/One Month Free Email Campaign/{} - OMF - IT - 2 Cities.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    df2city_final.to_excel(storage_dir_IT_2city, index=False)

    #Copy to the DWMA folder
    shutil.copy2(storage_dir_IT_2city, DMWA_dir_IT_2city)
    N_IT_2city=len(df2city_final.index)

    storage_dir_IT_french = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/storagegroup/Digital Marketing/Emails/One Month Free Email Campaign/Email Lists/{} - OMF - IT - French.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    DMWA_dir_IT_french = "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/DMWA/Storage/Emails/One Month Free Email Campaign/{} - OMF - IT - french.xlsx".format(
        datetime.datetime.strftime(datetime.date.today(), "%m.%d.%y")
    )

    dffrenchcity.to_excel(storage_dir_IT_french, index=False)

    #Copy to the DWMA folder
    shutil.copy2(storage_dir_IT_french, DMWA_dir_IT_french)
    N_IT_French=len(dffrenchcity.index)

    with open(
        "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/storagegroup/Digital Marketing/Emails/One Month Free Email Campaign/Documentation/OMF Code Run Log_IT.txt",
        "a",
    ) as f:
        f.write("\n")  # add a new line
        f.write(run_date + "|" + "Run Successfully" + "|" + "N/A")

    # sending email
    recipients = ["fengyao_luo@uhaul.com", "gaurang_makharia@uhaul.com", "haris_heldic@uhaul.com"]
    message = "Hello<br><br>The OMF Email Campaign code run was successful today :D <br><br>" + str(f'{N_file:,}')+ " Records Originally <br>" + str(f'{N_createdate_removed:,}')+ " Records Have Create Date Older Than 3 Months Ago<br>" + str(f'{N_Duplicates:,}')+ " Records are Duplicates<br>" + str(f'{N_Invalidemail_removed:,}') + " Records Have Invalid Emails<br>" + str(f'{N_UBox:,}')+ " Records are U-Box<br><br>" + str(f'{N_IT_final:,}') + " Records in IT<br>" + str(f'{N_IT_2city:,}') + " Records in IT - 2 Cities<br>" + str(f'{N_IT_French:,}') + " Records in IT - French<br><br> Thank you,<br>Fengyao"
    subject = "OMF Email Campaign Code Run Status - IT"
    vba_email(recipients, subject, message)

    #put numbers into SQL table
    current_Date = str(date.today())
    now = datetime.datetime.now()
    current_Time = now.strftime("%H:%M:%S")

    data= [[current_Date,N_file,N_createdate_removed,N_Duplicates,N_Invalidemail_removed,N_UBox,N_IT_final, N_IT_2city, N_IT_French,current_Time]]

    df = pd.DataFrame(data, columns = ['Date','N_file', 'N_createdate_removed', 'N_Duplicates', 'N_Invalidemail_removed', 'N_UBox' ,'N_IT_final', 'N_IT_2city', 'N_IT_French','Time'])

    base_con = (
            "Driver={ODBC Driver 13 for SQL Server};"
            "Server=OPSReport02.uhaul.amerco.org;"
            "Database=StorageReporting;"
            "UID=1248505;"
            "PWD=Fengyao505L;"
            )


    print("Uploading Data to SQL...")
    # URLLib finds the important information from our base connection
    params = urllib.parse.quote_plus(base_con)
    # SQLAlchemy takes all this info to create the engine
    engine = sqlalchemy.create_engine("mssql+pyodbc:///?odbc_connect=%s" % params)

    df.to_sql("OMF_Email_Counts_IT", engine, if_exists="append", index=False)

except Exception as e:
    with open(
        "\\\\adfs01.uhi.amerco/INTERDEPARTMENT/storagegroup/Digital Marketing/Emails/One Month Free Email Campaign/Documentation/OMF Code Run Log_IT.txt",
        "a",
    ) as f:
        f.write("\n")  # add a new line
        f.write(run_date + "|" + "Run Unsuccessfully" + "|" + str(e))

    # sending email
    recipients = ["fengyao_luo@uhaul.com","gaurang_makharia@uhaul.com", "haris_heldic@uhaul.com"]
    message = "Hello<br><br>The OMF Email Campaign code run was NOT successful today :( <br><br> Please Check <br><br>Thank you,<br>Fengyao"
    subject = "OMF Email Campaign Code Run Status - IT"
    vba_email(recipients, subject, message)
