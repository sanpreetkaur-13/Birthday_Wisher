import pandas as pd
import datetime
import smtplib
import os

Mail_id=-""
Mail_password=""
def sendEmail(to, sub,msg):
    print(f"Email to {to} sent with sub {sub} and msg {msg}")
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(Mail_id,Mail_password)
    s.sendmail(Mail_id,to,f"Subject:{sub}\n\n{msg}")
    s.quit()
if __name__=="__main__":
    df=pd.read_excel("Birthday_list.xlsx")
    today=datetime.datetime.now().strftime("%d-%m")
    #print(type(today))
    yearNow=datetime.datetime.now().strftime("%Y")
    writeInd=[]
    for index,item in df.iterrows():
        #print(index,item['Birthday'])
        bdy=item['Birthday'].strftime("%d-%m")
        if(today==bdy) and yearNow not in str(item['Year']):
            sendEmail(item['Email'],"Birthday Wishes",item['Dialogue'])
            writeInd.append(index)
    #print(writeInd)
    for i in writeInd:
        #yr=df.loc[i,'Year']
        df.loc[i,'Year']=str(yr)+','+str(yearNow)

    df.to_excel('Birthday_list.xlsx', index=False)