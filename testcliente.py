from openpyxl import load_workbook
import xlrd, datetime, re
import pandas as pd

formato=[]

filename=r'D:\Descargas\TestFormatooo.xlsx' 
tfares=r'D:\Descargas\TarifMaster.xlsx'
 
def tariifario(filas):
    trf=load_workbook(tfares)
    ws = trf.worksheets[0]
    origin_Nomenclature_row = list(ws.rows)[1]
    destination_site_col=list(ws.columns)[2]
    unity_row=list(ws.rows)[2]
    dest=[cell.value for cell in destination_site_col]
    unities=[cell.value for cell in unity_row]
    for f in filas:
        special_case_flag=0
        site=f[3].split()[0]
        state=f[6].split('/')[0].strip()
        destination=f[6].split('/')[1].strip()
        unity_type=f[4]
        print(unity_type,'Unity type')
        print(site,'o')
        print(state,'S')
        print(destination,'d')
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
        print(destination_index,'d')
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
        print(fare_CV,'fare')

def validation():
    workbook = xlrd.open_workbook(filename)        #Determina el numero de filas
    sheet=workbook.sheet_by_index(0)
    row_count=sheet.nrows 
    print(row_count)
    wb = load_workbook(filename,read_only=True, data_only=True)       #Determina el numero de filas
    ws = wb.worksheets[0]
    for i in range(2,13):
        formato.append(ws.cell(8,i).value.upper())
    fecha=str(ws.cell(7,5).value).split()[0]
    if fecha==str(datetime.date.today()):
        pass
    if re.match('NO.?',formato[0]) and re.match('.*?ID.*?OTM', formato[1]) and re.match('.*?L[IÍ]NEA', formato[2]) and re.match('.*?PREDIO', formato[3]) and re.match('.*?UNIDAD', formato[4]) and re.match('C[P]?([ÓO]DIGO POSTAL)?', formato[5]) and re.match('POBLACI[ÓO]N', formato[6]) and re.match('C[V]?(ONTROL VEH[ÍI]CULAR)?', formato[7]) and re.match('.*?TARIFA', formato[8]) and re.match('AUTORIZA', formato[9]) and re.match('.*?IMPORTE', formato[10]):
        pass
    if type(None) in formato:
        pass
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

    print(filas)
    print(noProcessedCV)
    tariifario(filas)
    # noprocesados(noProcessedCV)
validation()

