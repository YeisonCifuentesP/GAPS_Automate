import pandas as pd
import pyodbc

import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.application import MIMEApplication

from datetime import datetime
from datetime import timedelta



dateHourFinish = datetime.now()
dateHourBegin = dateHourFinish - timedelta(days=1)

dateHourFinishFormat = (dateHourFinish.strftime("%b %d %Y %H:%M"))
dateHourBeginFormat = (dateHourBegin.strftime("%b %d %Y %H:%M"))


dateHourFinishSQL = (dateHourFinish.strftime("%Y-%m-%d %H:%M:%S"))
dateHourBeginSQL = (dateHourBegin.strftime("%Y-%m-%d %H:%M:%S"))

#Consultation to obtain and process GAPS
def getGaps(deviceId):
 # Conexion a la BD
    server = 'siindependenceserver.database.windows.net'
    database = 'OxyDB'
    username = 'indsii'
    password = '1nd3p3nd3nc32019*'
    driver = '{ODBC Driver 17 for SQL Server}'
    cnxn = pyodbc.connect('DRIVER='+driver+';SERVER='+server +
                          ';PORT=1433;DATABASE='+database+';UID='+username+';PWD=' + password)

    query = '''SELECT [Id]
                    ,[carga_gancho]
                    ,[fecha_hora]
                    ,[tstm]
                    ,[deviceId]
                FROM Oxy.Oxy_Operational_data_ALL
                WHERE  deviceId = '{}' and fecha_hora BETWEEN  '{}' AND '{}' '''.format(deviceId, dateHourBeginSQL, dateHourFinishSQL)

    df = pd.read_sql(query, cnxn)
    listGaps = []

    for i, row in df.iterrows():

        if(i != (len(df)-1)):
            valueGap = df.tstm[i+1]-row.tstm
        else:
            valueGap = 0
        listGaps.append(valueGap)
    df["gap"] = listGaps

    dfGaps = (df[(df.gap >= 30)])
    lowIndicator = (len(dfGaps[(dfGaps.gap >= 30) &
                               (dfGaps.gap < 60)]))
    mediumIndicator = (len(dfGaps[(dfGaps.gap >= 60) &
                                  (dfGaps.gap < 90)]))
    highIndicator = (len(dfGaps[(dfGaps.gap >= 90)]))

    if(lowIndicator > 0):
        lowIndicator = ((len(dfGaps[(dfGaps.gap >= 30) &
                                    (dfGaps.gap < 60)]))*100)/len(dfGaps)
        lowIndicator = str(round(lowIndicator, 2))+"%"
    else:
        lowIndicator = "0.00%"

    if(mediumIndicator > 0):
        mediumIndicator = ((len(dfGaps[(dfGaps.gap >= 60) &
                                       (dfGaps.gap < 90)]))*100)/len(dfGaps)
        mediumIndicator = str(round(mediumIndicator, 2))+"%"

    else:
        mediumIndicator = "0.00%"

    if(highIndicator > 0):
        highIndicator = ((len(dfGaps[(dfGaps.gap >= 90)]))*100)/len(dfGaps)
        highIndicator = str(round(highIndicator, 2))+"%"
    else:
        highIndicator = "0.00%"

    dfIndicator = pd.DataFrame({
        'Indicador': ["Bajo", "Medio", "Alto", "Total"],
        'Porcentaje': [(lowIndicator), (mediumIndicator), (highIndicator), str((len(dfGaps))) + " Registros"],
    })

    return dfIndicator

#Format of the tables containing the indicators
def create_html_table(df):

    row_data = ''

    for i in range(df.shape[0]):

        for j in range(df.shape[1]):
            if((i == 3)):
                if ((i % 2) != 0) & (j == 0):  # not an even row and the start of a new row
                    row_data += '\n<tr style="background-color:#f2f2f2"> \n <td class = "text_column">' + \
                        str(df.iloc[i, j])+'</td>'

                elif j == 0:  # The first column
                    row_data += '\n<tr> \n <td class = "text_column">' + \
                        str(df.iloc[i, j])+'</td>'

                elif j == 1:  # second column
                    row_data += '\n <td class = "number_column">' + \
                        str(df.iloc[i, j])+'</td>'
            else:
                if ((i % 2) != 0) & (j == 0):  # not an even row and the start of a new row
                    row_data += '\n<tr style="background-color:##ffffff"> \n <td class = "text_column">' + \
                        str(df.iloc[i, j])+'</td>'

                elif j == 0:  # The first column
                    row_data += '\n<tr> \n <td class = "text_column">' + \
                        str(df.iloc[i, j])+'</td>'

                elif j == 1:  # second column
                    row_data += '\n <td class = "number_column">' + \
                        str(df.iloc[i, j])+'</td>'

    return row_data

#html code containing the message body
def body_email(gapsRig23, gapsRig30, gapsRig43):

    htmlTableRig23 = create_html_table(gapsRig23)
    htmlTableRig30 = create_html_table(gapsRig30)
    htmlTableRig43 = create_html_table(gapsRig43)

    email_content = """<html>

    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
   
    <title>Indicador Ausencia de Datos</title>
    <style type="text/css">
        a {
        color: #2E86C1
        }

        body,
        #header h1,
        #header h2,
        p {
        margin: 0;
        padding: 0;
        }

        #main {
        border: 1px solid #cfcece;
        margin-top: 50px;
        }

        img {
        display: block;
        }

        #top-message p,
        #bottom p {
        color: #3f4042;
        font-size: 12px;
        font-family: Arial, Helvetica, sans-serif;
        }

        #header h1 {
        color: #ffffff !important;     
        font-family: Dax;
        font-size: 20px;
        margin-bottom: 0 !important;
        padding-bottom: 0;
        }

        #header p {
        color: #ffffff !important;
        
        font-family: "Lucida Grande", "Lucida Sans", "Lucida Sans Unicode", sans-serif;
        font-size: 12px;
        }

        h5 {
        margin: 0 0 0.8em 0;
        }

        h5 {
        font-size: 18px;
        color: #444444 !important;
        font-family: Arial, Helvetica, sans-serif;
        }

        p {
        font-size: 12px;
        color: #ffffff !important;
        font-family: "Lucida Grande", "Lucida Sans", "Lucida Sans Unicode", sans-serif;
        line-height: 1.5;
        }

        #links p {
        color: #2E86C1  !important;
        font-family: "Lucida Grande", "Lucida Sans", "Lucida Sans Unicode", sans-serif;
        font-size: 12px;
        }

        #content-3 h5 {
        color: #000000 !important;     
        font-family: Dax;
        font-size: 15px;
        margin-bottom: 0 !important;
        padding-bottom: 0;
        }

    </style>
    </head>

    <body>

    <table id="grad" width="100%"  background="http://www.skanhawk.com/wp-content/uploads/2021/10/slide_background_4.png" cellpadding="0" cellspacing="0" bgcolor="e4e4e4">
        <tr>
        <td>
            <table id="main" width="50%" align="center" cellpadding="0" cellspacing="15" bgcolor="ffffff">
            <tr>
                <td>
                <table id="header" cellpadding="10" cellspacing="0" align="center" bgcolor="8fb3e9">
                    <tr>
                    <td width="400" align="center" bgcolor="#00809D">
                        <h1>Indicador Ausencia de Datos</h1>
                        <p align="left">"""+dateHourBeginFormat+"""</p>
                        <p align="left">"""+dateHourFinishFormat+"""</p>
                    </td>
                    <td width="100" align="center" bgcolor="#00809D">
                        <img src="http://www.skanhawk.com/wp-content/uploads/2021/10/SkanHawk_logo.png" width="120" height="70" />
                    </td>
                    </tr>

                </table>
                </td>
            </tr>
            <tr>
                <td>
                <table id="content-3" cellpadding="0" cellspacing="0" align="center">
                    <tr>
                    <td width="200" valign="middle">
                        <h5 align="center">RIG 23</h5>                    
                    </td>
                    </tr>
                </table>
                <table id="content-3" cellpadding="0" cellspacing="0" align="center">
                    <tr>
                    <td width="200" valign="middle">
                        <p> """ + htmlTableRig23 + """</p>
                    </td>
                    </tr>
                </table>
                </td>
            </tr>
            <tr>
                <td>
                <table id="content-3" cellpadding="0" cellspacing="0" align="center">
                    <tr>
                    <td width="200" valign="top">
                        <h5 align="center">RIG 30</h5>                    
                    </td>
                    </tr>
                </table>
                <table id="content-3" cellpadding="0" cellspacing="0" align="center">
                    <tr>
                    <td width="200" valign="middle">
                        <p> """ + htmlTableRig30 + """</p>
                    </td>
                    </tr>
                </table>
                </td>
            </tr>
            <tr>
                <td>
                <table id="content-3" cellpadding="0" cellspacing="0" align="center">
                    <tr>
                    <td width="200" valign="top">
                        <h5 align="center">RIG 43</h5>                    
                    </td>
                    </tr>
                </table>
                <table id="content-3" cellpadding="0" cellspacing="0" align="center">
                    <tr>
                    <td width="200" valign="top">
                        <p> """ + htmlTableRig43 + """</p>
                    </td>
                    </tr>
                </table>
                </td>
            </tr>
            <tr id="links">
                <td align="center">
                    <p><a
                        href="https://app.powerbi.com/groups/me/apps/12695a14-582b-4d23-8042-6279b413c8bc/dashboards/6c55759c-2dac-4b63-afc0-8cca4f5b3585?language=en-US">Rig
                        23</a> |
                    <a
                        href="https://app.powerbi.com/groups/me/apps/179b2f4f-5639-4bab-881d-1d90a6432dbe/dashboards/3a50495c-b517-4263-bb55-440514660d66?language=en-US">Rig
                        30</a> |
                    <a
                        href="https://app.powerbi.com/groups/me/apps/ebcd7c98-6175-4283-bd55-6b6fb22deb0a/dashboards/41c77b01-d580-487a-9547-08d821f57543?language=en-US">Rig
                        43</a>
                    </p>
                </td>
                </tr>
            </table>
            <table id="bottom" cellpadding="20" cellspacing="0" width="600" align="center">
            <tr>
                <td align="center">
                <p>Copyright &copy; 2021 by SkanHawk Enterprise | Designed by Analytic Area | Developed by Yeison Cifuentes</p>
                </td>
            </tr>
            </table>
        </td>
        </tr>
        
    </table><!-- wrapper -->

    </body>

    </html>
    """

    return email_content

#Protocol send email
def send_email(gapsRig23, gapsRig30, gapsRig43):

    email_content = body_email(gapsRig23, gapsRig30, gapsRig43)
    dateHourFinishEmail = (dateHourFinish.strftime("%d/%m/%Y %H:%M"))
    dateHourBeginEmail = (dateHourBegin.strftime("%d/%m/%Y %H:%M"))

    msg = MIMEMultipart()
    msg['Subject'] = 'Estado Transmisión Plataforma SkanHawk Zona Arauca ' + \
        dateHourBeginEmail + " - " + dateHourFinishEmail

    msg['From'] = 'informacion@skanhawk.com'
    msg['To'] = 'lmrincon@skanhawk.com,yfcifuentes@skanhawk.com'
    # msg["Cc"] = "serenity@example.com,inara@example.com"
    password = "Inde3030*"
    

       #msg.attach(email_content)
    pdf = MIMEApplication(open("D:\Independence\Sistema\Desktop\SKAN_HAWK\DOCUMENTACION_SKAN\DOCUMENTOS\Test.xlsx", 'rb').read())
    pdf.add_header('Content-Disposition', 'attachment', filename= "example.xlsx")
    msg.attach(pdf)
   
    html=MIMEApplication(email_content)
    html.add_header('Content-Type', 'text/html')
    msg.attach(html)

    try:
        s = smtplib.SMTP('smtp-mail.outlook.com: 587')
        s.starttls()

        # Login Credentials for sending the mail
        s.login(msg['From'], password)

        # s.sendmail(msg['From'], msg["To"].split(",") + msg["Cc"].split(","), msg.as_string())
        s.sendmail(msg['From'], msg["To"].split(","), msg.as_string())
        s.quit()
    except:
        print('Ocurrio una excepción al enviar el correo')
