import pandas as pd
import datetime
import smtplib

GAMI_ID = "aseemg786@gmail.com"
GMAIL_PSWD = "
def sendEmail(to, sub, msg):
    print(f"Email to {to} send with subject: {sub} and message {msg} ")
    s = smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GAMI_ID,GMAIL_PSWD)
    s.sendmail(GAMI_ID,to,f"Sublect: {sub}\n\n{msg}")
    s.quit()


if __name__ == "__main__":
    df = pd.read_excel("data.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearmow = datetime.datetime.now().strftime("%Y")
    writeInd = []
    # print(type(today))
    for index, item in df.iterrows():
        # print(index,item["Birthday"])
        bday = item['Birthday'].strftime("%d-%m")
        # print(bday)
        if (today == bday) and yearmow not in str(item['Year']) :
            sendEmail(item['Email'], "Happy Bithday", item['Dialgue'])
            writeInd.append(index)

    # print(writeInd)   
    for i in writeInd:
        yr = df.loc[i, 'Year']     
        df.loc[i,'Year'] = str(yr) + ',' +  str(yearmow)
        # print(df.loc[i,'Year'])
         # print(df)
    df.to_excel('data.xlsx',index=False)
