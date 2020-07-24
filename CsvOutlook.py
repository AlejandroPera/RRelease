import win32com.client as win32
import datetime,xlrd,re,os,shutil,time
from openpyxl import load_workbook

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
account= win32.Dispatch("Outlook.Application").Session.Accounts
flagTime=0
formato=[]
tfares=r'C:\Users\aperalda\Documents\AltaDeTarifas\TarifMaster.xlsx'
 
def txt_to_str(route):
    f = open(route, mode="r", encoding="utf-8")
    content = f.read()
    f.close()
    return str(content)

def noProcessed(noProcessedCV):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    destinatarios=''
    mail.To = destinatarios
    mail.Subject='CV no procesados'
    f = open(r'C:\Users\aperalda\Documents\adicionales\mail.txt', 'w', encoding='UTF-8')
    f.write('<!DOCTYPE html> <html> <head> <title>FORMATO TEXT</title> </head> <style> table, th, td {border: 1px solid black;border-collapse: collapse; text-align: center;}</style><meta charset="UTF-8"> <body style="background-color:#FFE406"> <img src="llpc.png" alt="DHL LOGO.png"> <h2 style="font-family:verdana;text-align:center;">CV no procesados</h2> <p style="font-family:verdana;">Esta es una alerta informativa sobre CV no capturados correctamente:</p><table style="width:100%"><tr><th>CV/th></tr>')
    for i in range(len(noProcessedCV)):
        f.write("<tr> <td>%s</td> </tr>\n" %(noProcessedCV[i]))
    f.write('</table></body></html>')
    f.close()
    body=txt_to_str(r'C:\Users\aperalda\Documents\adicionales\mail.txt')
    mail.HTMLBody = body
    images_path = "C:\\Users\\aperalda\\Documents\\RAudit\\RateAudit\\img\\"
    mail.Attachments.Add(Source= images_path+"llpc.PNG")
    mail.Send()
    print('No procesados enviado')



def tariifario(filas,filename):
    trf=load_workbook(tfares)
    ws = trf.worksheets[0]
    origin_Nomenclature_row = list(ws.rows)[1]
    destination_site_col=list(ws.columns)[2]
    unity_row=list(ws.rows)[2]
    dest=[cell.value for cell in destination_site_col]
    unities=[cell.value for cell in unity_row]
    tarifas=[]

    for f in filas:
        print(f)
        if type(f[10]) == str:
            special_case_flag=0
            site=f[3].split()[0]
            state=f[6].split('/')[0].strip()
            destination=f[6].split('/')[1].strip()
            unity_type=f[4]
            # print(unity_type,'Unity type')
            # print(site,'o')
            # print(state,'S')
            # print(destination,'d')
            if destination in ['LA PAZ', 'BENITO JUAREZ','CALERA']:
                special_case_flag=1
            if special_case_flag==0:
                destination_index=dest.index(destination)+1
            else:
                if destination=='LA PAZ':
                    if state=='BCS':
                        destination_index=102
                    else:
                        destination_index=23
                elif destination=='CALERA':
                    if state=='ZAC':
                        destination_index=502
                    else:
                        destination_index=514
                else:
                    if state=='QTR':
                        destination_index=119
                    else:
                        destination_index=133

            # print(destination_index,'d')


            if site=='015':
                indexUnity=unities.index(unity_type,4,11)+1
            elif site=='009':
                indexUnity=unities.index(unity_type,12,26)+1
            elif site=='037':
                indexUnity=unities.index(unity_type,27,34)+1
            elif site=='140':
                indexUnity=unities.index(unity_type,35,42)+1
            elif site=='130':
                indexUnity=unities.index(unity_type,43,50)+1
            elif site=='139':
                indexUnity=unities.index(unity_type,51,58)+1
            elif site=='151':
                indexUnity=unities.index(unity_type,59,66)+1
            elif site=='187':
                indexUnity=unities.index(unity_type,67,74)+1
            elif site=='004':
                indexUnity=unities.index(unity_type,75,90)+1
            elif site=='051':
                indexUnity=unities.index(unity_type,91,106)+1
            elif site=='100':
                indexUnity=unities.index(unity_type,107,114)+1
            elif site=='108':
                indexUnity=unities.index(unity_type,115,122)+1
            elif site=='116':
                indexUnity=unities.index(unity_type,123,138)+1
            elif site=='065':
                indexUnity=unities.index(unity_type,139,146)+1
            elif site=='016':
                indexUnity=unities.index(unity_type,147,155)+1
            elif site=='083':
                indexUnity=unities.index(unity_type,156,159)+1
            elif site=='186':
                indexUnity=unities.index(unity_type,160,167)+1
            elif site=='002':
                indexUnity=unities.index(unity_type,168,176)+1
            elif site=='014':
                indexUnity=unities.index(unity_type,177,184)+1
            elif site=='019':
                indexUnity=unities.index(unity_type,185,193)+1
            elif site=='035':
                indexUnity=unities.index(unity_type,194,202)+1
            elif site=='024':
                indexUnity=unities.index(unity_type,203,211)+1
            elif site=='146':
                indexUnity=unities.index(unity_type,212,219)+1
            elif site=='027':
                indexUnity=unities.index(unity_type,220,227)+1
            elif site=='031':
                indexUnity=unities.index(unity_type,228,236)+1
            elif site=='132':
                indexUnity=unities.index(unity_type,237,244)+1
            elif site=='115':
                indexUnity=unities.index(unity_type,245,252)+1
            elif site=='145':
                indexUnity=unities.index(unity_type,253,262)+1
            elif site=='182':
                indexUnity=unities.index(unity_type,263,270)+1
            elif site=='185':
                indexUnity=unities.index(unity_type,271,275)+1
            print(destination_index,indexUnity,'VectorMatricial')
            fare_CV=ws.cell(destination_index,indexUnity).value
            # print(fare_CV,'fare')
            tarifas.append([fare_CV, f[7]])
    print(tarifas)
    dest = shutil.copy(filename, r'C:\Users\aperalda\Documents')
    return tarifas

def validation(filename,mail):
    workbook = xlrd.open_workbook(filename)        #Determina el numero de filas
    sheet=workbook.sheet_by_index(0)
    row_count=sheet.nrows 
    print('cantidad de tarifas a evaluar: ',row_count)
    wb = load_workbook(filename,read_only=True, data_only=True)       #Determina el numero de filas
    ws = wb.worksheets[0]
    for i in range(2,13):
        formato.append(ws.cell(8,i).value.upper())
    # fecha=str(ws.cell(7,5).value).split()[0]
    # if fecha==str(datetime.date.today()):
    #     pass
    if re.match('NO.?',formato[0]) and re.match('.*?ID.*?OTM', formato[1]) and re.match('.*?L[IÍ]NEA', formato[2]) and re.match('.*?PREDIO', formato[3]) and re.match('.*?UNIDAD', formato[4]) and re.match('C[P]?([ÓO]DIGO POSTAL)?', formato[5]) and re.match('POBLACI[ÓO]N', formato[6]) and re.match('C[V]?(ONTROL VEH[ÍI]CULAR)?', formato[7]) and re.match('.*?TARIFA', formato[8]) and re.match('AUTORIZA', formato[9]) and re.match('.*?IMPORTE', formato[10]):
        pass
    else:
        forwardFormatMailError(mail)
    # if type(None) in formato:
    #     pass
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
                    else:
                        fila.append(ws.cell(i,a).value)
                else:
                    fila.append('')
            else:
                fila.append('')
        if '' in fila:
            noProcessedCV.append(fila[7])
        else:
            filas.append(fila)

    print('tarifas a procesar: ', len(filas))
    print('tarifas no procesadas: ', len(noProcessedCV))
    wb.close()
    tarifas_CV=tariifario(filas,filename)
    tarifas_CV.append(noProcessedCV)
    return tarifas_CV


def forwardFormatMailError(mail):
    reply=mail.Forward()
    sender=mail.Sender
    newBody = "Formato Incorrecto"
    reply.HTMLBody = newBody + reply.HTMLBody
    reply.To=sender
    reply.Send()
    print('Mandado error de formato en xlsx')

# def forwardDateError(mail):
#     reply=mail.Forward()
#     sender=mail.Sender
#     print(sender)
#     newBody='Archivo expirado'
#     reply.HTMLBody = newBody + reply.HTMLBody
#     reply.To=sender
#     reply.Send()
#     print('Mandado')


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
    print('Mandado error de archivo adjunto')

def retrieval():
    flag=0

    global flagTime
    global nowT
    print('actualizando')


    if flagTime==1:
        timeElapsed=datetime.datetime.now()-nowT
        timeElapsed=timeElapsed.seconds
        timeToWait=60-timeElapsed
        print(timeToWait)
        time.sleep(timeToWait)
        flagTime=0

    nowT=datetime.datetime.now()
    inbox = outlook.GetDefaultFolder(6)                         # "6" refers to the index of a folder - in this case the inbox.                                      
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    if nowT+datetime.timedelta(seconds=60)>datetime.datetime.now() and flagTime==0:
        flagTime=0
        print('leyendo')
        for message in messages:
            if message.Unread==True and message.Subject.upper() == 'ALTA DE TARIFAS RPA':
                numberOfAttachments=len(message.Attachments)
                if numberOfAttachments==1:
                    print('Guardando')
                    flag=1
                    file_name=r'C:\Users\aperalda\Downloads' + '\\'+ message.Attachments[0].FileName
                    print(file_name)
                    message.Attachments[0].SaveAsFile(file_name)
                    message.Unread = False
                    print('Mensaje leído')
                    break
                else:
                    message.Unread=False
                    time.sleep(1)
                    print('Error en attachment')
                    createReply(message,numberOfAttachments)
                    flagTime=1
    if flag==1:
        flag=0
        masterArray=validation(file_name, message)
        time.sleep(2)
        os.remove(file_name)
        flagTime=1
        return masterArray
    else:
        print('No se encontró alta de tarifas')
        flagTime=1


while 1:
    print(retrieval())


