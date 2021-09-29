import pandas as pd
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


#READ THE EMAILS FROM EXCEL FILE

data=pd.read_excel("Employees_Details.xlsx")
#print(type(data))
email_col=data.get("Employee_Email")
list_of_emails=list(email_col)
print(list_of_emails)


try:
    
    server=sm.SMTP("smtp.gmail.com",587)
    server.starttls()
        
    server.login("nainageorge04@gmail.com","pass)
    from_="nainageorge04@gmail.com"
    to_=list_of_emails
    
    message=MIMEMultipart("alternative")
    message['Subject']="Hello, this is Ayushi George! Have a nice Day!"
    message['from']="nainageorge04@gmail.com"
    
    html='''
    <html>
    <head>
    <style>
    table {
      font-family: arial, sans-serif;
      border-collapse: collapse;
      width: 100%;
    }

    td, th {
      border: 1px solid #dddddd;
      text-align: left;
      padding: 8px;
    }

    tr:nth-child(even) {
      background-color: #dddddd;
    }
    </style>
    </head>
    <body>
    <p>These are the MY ACHIEVEMENTS :</p>
    <h2>ACHIEVEMENTS</h2>
    
    <table>
      <tr>
        <th>DATE</th>
        <th>YEAR</th>
        <th>ACTIVITY</th>
      </tr>
      <tr>
        <td>17 January 2021</td>
        <td>FIRST YEAR</td>
        <td>HACKATHON TEAM LEADER</td>
      </tr>
      <tr>
        <td>8 Feburary 2021</td>
        <td>FIRST YEAR </td>
        <td>BAGGED 4TH POSITION IN DEBATE COMPETITION</td>
      </tr>
      <tr>
         <td>23 January 2021</td>
        <td>FIRST YEAR</td>
        <td>E-MEET WITH MANEKA GANDHI</td>
      </tr>
      <tr>
        <td>25 January 2021  </td>
        <td>FIRST YEAR</td>
        <td>PAID CONTENT WRITING INTERNSHIP</td>
      </tr>
        <tr>
        <td>26 January 2021  </td>
        <td>FIRST YEAR</td>
        <td>SOCIAL MEDIA  MARKETING INTERNSHIP</td>
      </tr>
        <tr>
        <td>31 January 2021  </td>
        <td>SECOND YEAR</td>
        <td>CLASS REPRESENTATIVE</td>
      </tr>
    </table>

    </body>
    </html>

 
    '''
    
    text=MIMEText(html,"html")
    
    message.attach(text)
    
    server.sendmail(from_,to_,message.as_string())
    print("Message has been send to the emails .")
    
except exeption as e :
    print(e)
