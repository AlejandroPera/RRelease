from openpyxl import load_workbook
import xlrd, datetime, re
import pandas as pd

formato=[]

filename=r'C:\Users\aperalda\Desktop\TestFormato.xlsx' 
tfares=r'C:\Users\aperalda\Documents\AltaDeTarifas\TarifMaster.xlsx'

def tariifario(filas):
    # tarifario=pd.read_excel(tfares)
    # print(tarifario)
    trf=load_workbook(tfares)
    ws = trf.worksheets[0]
    origin_Nomenclature_row = list(ws.rows)[1]
    destination_site_col=list(ws.columns)[2]
    unity_row=list(ws.rows)[2]
    dest=[cell.value for cell in destination_site_col]
    # origin=[cell.value for cell in origin_Nomenclature_row]
    unities=[cell.value for cell in unity_row]
    # print(origin)
    print(dest)
    print(unities)
    for f in filas:
        special_case_flag=0
        site=f[3].split()[0]
        state=f[6].split('/')[0].strip()
        destination=f[6].split('/')[1].strip()
        unity_type=f[4]
        print(site,'o')
        print(state,'S')
        print(destination,'d')
        if destination in ['LA PAZ', 'BENITO JUAREZ','CALERA']:
            special_case_flag=1
        # origin_index=origin.index(site)+1
        # print(origin_index,'o')
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
            indexUnity=unities.index(unity_type,4,10)
        elif site=='009':
            indexUnity=unities.index(unity_type,12,25)
        elif



def validation():
    workbook = xlrd.open_workbook(filename)        #Determina el numero de filas
    sheet=workbook.sheet_by_index(0)
    row_count=sheet.nrows 
    print(row_count)
    wb = load_workbook(filename,read_only=True, data_only=True)       #Determina el numero de filas
    ws = wb.worksheets[0]
    for i in range(2,13):
        formato.append(ws.cell(8,i).value.upper())
    print(formato)
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
                    fila.append(ws.cell(i,a).value)
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


