import pandas as pd
import datetime
import smtplib

# enter credentials
GMAIL_ID  = 'sirjansk@gmail.com'
GMAIL_PW = 'thisisafakepassword'
def sendEmail(to,sub,msg):
    print(f"Email to {to} sent with the subject {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PW)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    
    s.quit()
    
if __name__ == "__main__":

    df = pd.read_excel("data.xlsx")
    today = datetime.datetime.now().strftime("%d-%m")
    currentYear = datetime.datetime.now().strftime("%Y")
    writeInd = []

    for index, item in df.iterrows():
        holiday = item['Festival'].strftime("%d-%m")
        if(today == holiday and currentYear not in str(item['Year'])):
            sendEmail(item['Email ID'],item['Subject'], item['Message'])
            writeInd.append(index)
                 

    print (writeInd)    
    for i in writeInd:
        yr = df.loc[i, 'Year']
        print(yr)
        df.loc[i, 'Year'] = str(yr) + ',' + str(currentYear)

    df.to_excel('data.xlsx', index = False)
