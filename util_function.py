from oauth2client.service_account import ServiceAccountCredentials
import gspread
from datetime import date
import pandas as pd
import smtplib
import ssl

def connect_to_sheet(key_file):
    # Authorize the API
    scope = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file'
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(key_file, scope)
    client = gspread.authorize(creds)

    return(client)

def open_sheet(client, sheet_name):
    sheet = client.open(sheet_name)

    return(sheet)

def open_worksheet(sheet, name):

    sheet_instance = sheet.worksheet(name)

    return(sheet_instance)


#Introducir checkbox
def pasar_lista(sheet, sheet_instance, presencial, next_row, num_alumnos):

    sheetId = sheet_instance._properties['sheetId']

    type_alumno = create_type_alumno(presencial)

    requests = create_requests(sheetId, next_row, num_alumnos, type_alumno)

    sheet.batch_update(requests)


#Introducir fecha
def crear_fechas(sheet_instance, next_row, num_alumnos):

    today = date.today()
    today = [today.strftime("%d/%m/%Y")]

    range_date = 'A' + str(next_row + 1) + ':A' + str(next_row + num_alumnos + 1)

    sheet_instance.batch_update([{
        'range': range_date,
        'values': [today for i in range(num_alumnos)]}])


#Introducir nombres
def crear_nombres(sheet_instance, original_names, next_row, num_alumnos):

    names = sheet_instance.range(original_names)
    names = [[name.value] for name in names]

    range_names = 'B' + str(next_row + 1) + ':B' + str(next_row + num_alumnos + 1)
    sheet_instance.batch_update([{
        'range': range_names,
        'values': names}])


def create_type_alumno(presencial):

    alumno_clase = {"values": [{"userEnteredValue": {"boolValue": False}}, {"userEnteredValue": {
                        "boolValue": True}}, {"userEnteredValue": {"boolValue": False}}]}
    alumno_casa = {"values": [{"userEnteredValue": {"boolValue": True}}, {"userEnteredValue": {
                        "boolValue": False}}, {"userEnteredValue": {"boolValue": False}}]}
    alumno_falta = {"values": [{"userEnteredValue": {"boolValue": False}}, {"userEnteredValue": {
                        "boolValue": False}}, {"userEnteredValue": {"boolValue": True}}]}

    tipos = [alumno_clase, alumno_casa, alumno_falta]
    type_alumno = [tipos[i] for i in list(map(int, presencial))]

    return(type_alumno)


def create_requests(sheetId, next_row, num_alumnos, type_alumno):

    requests = {"requests": [
        {
            "repeatCell": {
                "cell": {"dataValidation": {"condition": {"type": "BOOLEAN"}}},
                "range": {"sheetId": sheetId, "startRowIndex": next_row, "endRowIndex": (next_row + num_alumnos),
                          "startColumnIndex": 2, "endColumnIndex": 5},
                "fields": "dataValidation"
            }
        },
        {
            "updateCells": {
                "rows": [
                    type_alumno
                ],
                "start": {"rowIndex": next_row, "columnIndex": 2, "sheetId": sheetId},
                "fields": "userEnteredValue"
            }
        }
    ]}

    return(requests)


def check_new_day(sheet_instance, original_names):

    names = sheet_instance.range(original_names)
    petardeo = False

    if len(names) != 18:
        petardeo = True

    if names[3].value != 'Aberto Lara':
        petardeo = True

    if names[14].value != 'Javier Aparicio':
        petardeo = True

    if petardeo:
        send_mail("La estamos liando petarda, hecha el freno macareno!", "arana.ieszizur@gmail.com")

    return(petardeo)


def send_mail(mail_text, send_to):

    mail_gmail = 'arana.ieszizur@gmail.com'
    pass_gmail = ''

    # Enviamos un mail con los eventos disponibles
    smtp_server = "smtp.gmail.com"
    port = 587  # For starttls

    # Create a secure SSL context
    context = ssl.create_default_context()

    # Try to log in to server and send email
    try:
        server = smtplib.SMTP(smtp_server, port)
        server.ehlo()  # Can be omitted
        server.starttls(context=context)  # Secure the connection
        server.ehlo()  # Can be omitted
        server.login(mail_gmail, pass_gmail)

        subject = "PASANDO LISTA"

        message = 'Subject: {}\n\n{}'.format(subject, mail_text)

        server.sendmail(mail_gmail, send_to, message.encode('utf-8'))

    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.quit()


