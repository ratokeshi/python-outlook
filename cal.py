import datetime as dt
import pandas as pd
import win32com.client
import fnmatch, sys
import os.path
from os import path 
from shutil import copyfile

def get_calendar(begin,end):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
    calendar = calendar.Restrict(restriction)
    return calendar

def get_appointments(calendar,subject_kw = None,exclude_subject_kw = None, body_kw = None):
    if subject_kw == None:
        appointments = [app for app in calendar]    
    else:
        appointments = [app for app in calendar if subject_kw in app.subject]
    if exclude_subject_kw != None:
        appointments = [app for app in appointments if exclude_subject_kw not in app.subject]
    cal_subject = [app.subject for app in appointments]
    cal_start = [app.start for app in appointments]
    cal_end = [app.end for app in appointments]
    cal_body = [app.body for app in appointments]
    cal_recipients = [app.recipients for app in appointments]

    df = pd.DataFrame({'subject': cal_subject,
                       'start': cal_start,
                       'end': cal_end,
                       'body': cal_body,
                       'recipients': cal_recipients})
    df.to_csv('results.txt', index=False)                   
    return df

def make_cpd(appointments):
    appointments['Date'] = appointments['start']
    appointments['Hours'] = (appointments['end'] - appointments['start']).dt.seconds/3600
    print(appointments)
    appointments.rename(columns={'subject':'Meeting Description','recipients':'Recipients','body':'Body'}, inplace = True)
    appointments.drop(['start','end'], axis = 1, inplace = True)
    print(appointments)
    summary = appointments.groupby('Meeting Description')['Hours'].sum()#.get_group('Recipients')
    return summary

def findfile(lookfor):
    lookfor="*" + lookfor + "*"
    #print("...Searching the following files for pattern match: " + str(os.listdir(os.path.dirname(__file__))))
    #print("...Looking at this list of files:")
    #print(os.listdir(os.path.dirname(__file__)))
    for file in os.listdir(os.path.dirname(__file__)):
        print("...Compare: " + file + " to pattern: " + lookfor)
        if fnmatch.fnmatch(file, lookfor):
            print(lookfor + ". is in directory")
            break
            #return True
        else:
            print(lookfor + ". is NOT a match.")
    return(file)        
     

# Set beginning and ending meeting search range
begin = dt.datetime(2021,1,17)
end = dt.datetime(2021,1,23)
filename="Meeting_Hours"
filetype="xlsx"
fileoutput=filename + "." + filetype


#get calander items to place in excel
cal = get_calendar(begin, end)
appointments = get_appointments(cal, subject_kw = ' ', exclude_subject_kw = 'Canceled')
result = make_cpd(appointments)
print(result)


# Create a timestamp for backup files 
now=dt.datetime.now()
timestamp="-" + str(now.year) + "-" + str("%02d"%now.month) + "-" + str("%02d"%now.day) + "--" + str("%02d"%now.hour) + str("%02d"%now.minute) + str("%02d"%now.second)
print("...Timestame is: " + timestamp)

#look for backup files if none create 
looking = findfile("backup")
print("...Searching the following files for pattern match: " + str(os.listdir(os.path.dirname(__file__))))
if findfile("backup"):
    print("...Yes, there is a backup here. ")
else:
    copyfile(fileoutput, filename + timestamp + "-backup." + filetype)
    print("...Creating first backup called: " + fileoutput, filename + timestamp + "-backup." + filetype)      
if path.exists(fileoutput):
    print("..." + fileoutput + " file Exists.")
    os.rename(fileoutput, fileoutput + timestamp + "-backup." + filetype)
    print("...Renamed file to: " + fileoutput + timestamp + "-backup." + filetype)
result.to_excel(fileoutput)
print("...New output file created: " + fileoutput)