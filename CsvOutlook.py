import win32com.client as win32
import datetime
from openpyxl import load_workbook
t=0

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
account= win32.Dispatch("Outlook.Application").Session.Accounts



def forwardFormatMailError(mail):
    reply=email.Forward()
    sender=email.Sender
    newBody = "Formato Incorrecto"
    reply.HTMLBody = newBody + reply.HTMLBody
    reply.To=sender
    reply.Send()
    print('Mandado')

def forwardDateError(mail):
    reply=email.Forward()
    sender=email.Sender
    print(sender)
    newBody='Archivo expirado'
    reply.HTMLBody = newBody + reply.HTMLBody
    reply.To=sender
    reply.Send()
    print('Mandado')

def parseXlsx(filename,mail):
    formato=[]
    formatFlag=0
    dateFlag=0
    wb = load_workbook(filename,read_only=True, data_only=True)       #Determina el numero de filas
    ws = wb.worksheets[0]
    workbook = xlrd.open_workbook(filename)        #Determina el numero de filas
    sheet=workbook.sheet_by_index(0)
    row_count=sheet.nrows
    fecha=str(ws.cell(7,5).value).split()[0]
    if fecha==datetime.date.today():
        dateFlag=1
    for i in range(2,13):
        formato.append(ws.cell(8,i).value.upper())
    if re.match('NO.?',formato[0]) and re.match('.*?ID.*?OTM', formato[1]) and re.match('.*?L[IÍ]NEA', formato[2]) and re.match('.*?PREDIO', formato[3]) and re.match('.*?UNIDAD', formato[4]) and re.match('C[P]?([ÓO]DIGO POSTAL)?', formato[5]) and re.match('POBLACI[ÓO]N', formato[6]) and re.match('C[V]?(ONTROL VEH[ÍI]CULAR)?', formato[7]) and re.match('.*?TARIFA', formato[8]) and re.match('AUTORIZA', formato[9]) and re.match('.*?IMPORTE', formato[10]):
        formatFlag=1
    if formatFlag==0:
        forwardFormatMailError(mail)
        return
    if dateFlag==0:
        forwardDateError(mail)
    filas=[]
    noProcessedCV=[]
    for i in range(9,row_count+1):
        fila=[]
        for a in range(2,13):
            if ws.cell(i,a).value != None:
                if ws.cell(i,a).value !='':
                    if a==9:
                        cv=str(ws.cell(i,a).value)
                        while len(cv)<8:
                            cv='0'+cv
                        fila.append(cv)
                        continue
                    fila.append(ws.cell(i,a).value)
            else:
                fila.append('')
        if '' in fila:
            noProcessedCV.append(fila[7])
        else:
            filas.append(fila)


def createReply(email,num):
    reply=email.Forward()
    sender=email.Sender
    print(sender)
    if num>1:
        newBody = "Debe haber únicamente un archivo adjunto"
    else:
        newBody=   'No existe archivo adjunto'
    reply.HTMLBody = newBody + reply.HTMLBody
    reply.To=sender
    reply.Send()
    print('Mandado')


def retrieval():
    flag=0
    inbox = outlook.GetDefaultFolder(6)                         # "6" refers to the index of a folder - in this case the inbox.                                      
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    for message in messages:
        if message.Unread==True:
            messageSubject=message.Subject
            messageSubjectUpper=messageSubject.upper()
            print(messageSubjectUpper)
            subjectSplit=messageSubjectUpper.split()
            if subjectSplit[0]=='ALTA' and subjectSplit[1]=='DE' and subjectSplit[2]=='TARIFAS' and subjectSplit[3]=='RPA':
                numberOfAttachments=len(message.Attachments)
                if numberOfAttachments==1:
                    print('Guardando')
                    flag=1
                    file_name=r'C:\Users\aperalda\Documents' + '\\'+ message.Attachments[0].FileName
                    message.Attachments[0].SaveAsFile(file_name)
                    mailLocated=message
                    break
                else:
                    print('Te va a caer prro')
                    createReply(message,numberOfAttachments)
    if flag==1:
        parseXlsx(file_name, mailLocated)
retrieval()


