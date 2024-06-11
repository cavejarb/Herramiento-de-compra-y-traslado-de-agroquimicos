import pandas as pd
from time import time
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import warnings
from openpyxl.utils.dataframe import dataframe_to_rows
pd.options.mode.chained_assignment = None 
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore', category=RuntimeWarning)
from shareplum import Office365
from shareplum.site import Version
from shareplum import Site
import numpy as np
from datetime import date
import datetime as dt
from datetime import datetime
from datetime import timedelta
import tkinter as tk
from tkcalendar import Calendar
today = date.today()
start_time = time()
row_limit = 5000

authcookie = Office365('https://sunshinebouquet1.sharepoint.com/', username='scastro@sunshinebouquet.com', password='CCyl3uwWUK6ZD6sf').GetCookies()
site = Site('https://sunshinebouquet1.sharepoint.com/sites/aprovisionamiento',version=Version.v2019, authcookie=authcookie)
siteDBLogistics = Site('https://sunshinebouquet1.sharepoint.com/sites/CosteodeTransporte',version=Version.v2019, authcookie=authcookie)
siteMatEmpaque = Site('https://sunshinebouquet1.sharepoint.com/sites/MatEmpaque',version=Version.v2019, authcookie=authcookie)

def readExcel(path,sheet):
    a = pd.read_excel(os.path.dirname(__file__)+f'\\{path}.xlsx',sheet)
    return a

def get_excel_sh(site, folder:str, namefile:str, sheetname:str,typeFolder:int,folder2:str):
#'Función para leer Excel Online Sharepoint'
    if typeFolder == 1:
        folder = site.Folder('Shared%20Documents/Indicadores/'+folder)
    elif typeFolder == 2:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Ordenes de compra/{folder}/{folder2}')
    elif typeFolder == 3:
        folder = site.Folder('Documentos%20compartidos/'+folder)
    elif typeFolder == 4:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Consumos/{folder}/{folder2}')
    elif typeFolder == 5:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Inventario Siesa/{folder}/{folder2}')
    elif typeFolder==6:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos')
    elif typeFolder==7:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Cotizaciones/{folder}/{folder2}')
    elif typeFolder==8:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Traslados internos/{folder}/{folder2}')
    elif typeFolder==9:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Productos vencidos/{folder}/{folder2}')
    elif typeFolder==10:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Inventario en almacenes/{folder}/{folder2}')
    else:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Cierre fincas/{folder}/{folder2}')
    df= pd.read_excel(folder.get_file(namefile), sheet_name=sheetname)
    return df 

def getListSP(nameListSP,fields,row_limit,idx):
    listSP = pd.DataFrame(columns=fields)
    more = True
    while more:
        query = {'Where': [('Gt', 'ID', str(idx))]}
        data = nameListSP.GetListItems(fields=fields, query=query, row_limit=row_limit)
        data_df = pd.DataFrame(data[0:])
        if not data_df.empty:
            listSP = listSP._append(data_df)
            ids = pd.to_numeric(data_df['ID'])
            idx = ids.max()
        else:
            more = False
    return listSP

# Split string with chars ";""#"
def trySplit(x):
    try:
        return x.split(";#")[1] 
    except: 
        return ''

def getList(site,nameList:str,fields,columnasAjuste = []): 
    print(f'Descargando {nameList}')
    listAux = site.List(nameList)
    Lista = getListSP(listAux,fields,row_limit,0)
    for columnaAjuste in columnasAjuste:
        Lista[columnaAjuste] = Lista[columnaAjuste].apply(lambda x : trySplit(x) )
    print(f'{nameList} descargada')
    return Lista

def ajustarNombresFincas(df,columnaNombre):
    try:
        df[columnaNombre] = df[columnaNombre].replace('BUENA VISTA','BUENAVISTA')
        df[columnaNombre] = df[columnaNombre].replace('LAS CUADRAS','CUADRAS')
        df[columnaNombre] = df[columnaNombre].replace('FLORES LA ESMERALDA MEDELLIN','ESMERALDA MED')
        df[columnaNombre] = df[columnaNombre].replace('FLORES DE TENJO','FLORES TENJO')
        df[columnaNombre] = df[columnaNombre].replace('LA FUENTE','FUENTE')
        df[columnaNombre] = df[columnaNombre].replace('EL JARDIN','JARDIN')
        df[columnaNombre] = df[columnaNombre].replace('LAURELES','LOS LAURELES')
        df[columnaNombre] = df[columnaNombre].replace('LA MARAVILLA','MARAVILLA')
        df[columnaNombre] = df[columnaNombre].replace('EL ROCIO','ROCIO')
        df[columnaNombre] = df[columnaNombre].replace('BUENA VISTA','BUENAVISTA')
        df[columnaNombre] = df[columnaNombre].replace('LAS CUADRAS','CUADRAS')
        df[columnaNombre] = df[columnaNombre].replace('FLORES LA ESMERALDA MEDELLIN','ESMERALDA MED')
        df[columnaNombre] = df[columnaNombre].replace('FLORES DE TENJO','FLORES TENJO')
        df[columnaNombre] = df[columnaNombre].replace('LA FUENTE','FUENTE')
        df[columnaNombre] = df[columnaNombre].replace('EL JARDIN','JARDIN')
        df[columnaNombre] = df[columnaNombre].replace('LAURELES','LOS LAURELES')
        df[columnaNombre] = df[columnaNombre].replace('LA MARAVILLA','MARAVILLA')
        df[columnaNombre] = df[columnaNombre].replace('EL ROCIO','ROCIO')
    except:
        pass
    return df

def obtenerConsumos(siteF,añoF,semanaF,semanaInventarioF,diccionarioFincasF,columnaMerge,adicionalesF):
    if adicionalesF==1:
        consumosF = get_excel_sh(siteF,añoF,f'Semana{semanaF}.xlsx','Sheet',4,f'Semana{semanaInventarioF}')
    else:
        consumosF = get_excel_sh(siteF,añoF,f'Semana{semanaF}.xlsx','Sheet',4,f'Semana{semanaInventarioF}/Adicionales')
    consumosF = eliminarPrimeraFila(consumosF,3)
    consumosF = pd.merge(consumosF, diccionarioFincasF, how='left' ,left_on= [columnaMerge], right_on= [columnaMerge])
    consumosF.drop(['Pesado'], inplace=True, axis=1)
    return consumosF

def create_excel(df,namewb:str,namesh:str): #'Función para crear Excel Local'
    excel_st = Workbook()
    ws = excel_st.active
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    filename = os.path.dirname(__file__)+f'\\{namewb}.xlsx' 
    ws.title = namesh
    excel = excel_st.save(filename=filename)
    print('------------------ Excel Creado ------------------\n')
    return excel

def file_upload_to_sharepoint(siteExport,folderAño:str,folderSemana:str,fileName:str,typeFolder:str):
    fileNamePath= os.path.dirname(__file__)+f'\\{fileName}.xlsx'
    if typeFolder==1:
        folder = siteExport.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Resultados traslados/{folderAño}/Semana{folderSemana}')
    else:
        folder = siteExport.Folder(f'Documentos%20compartidos/Compras/{folderAño}')
    with open(fileNamePath, mode='rb') as file:
        fileContent = file.read()
    folder.upload_file(fileContent, f'{fileName}.xlsx')

def eliminarEspacios (x):
    x = x.strip()
    return x

def eliminarPrimeraFila(df,type):
    df.columns = df.iloc[0]
    dfFinal = df.drop([df.index[0]])
    dfFinal = dfFinal.reset_index()
    #create_excel(dfFinal,"DfPrueba","Hoja1")
    if type==1:
        dfFinal.drop(['index','Disponible','Consumo a 2 Semanas','Consumo a 3 Semanas','Cierre a 2 Semanas'], inplace=True, axis=1)
    elif type==2:
        dfFinal.drop(['index'], inplace=True, axis=1)
    else:
        dfFinal = dfFinal[dfFinal["Estado"] == "ACTIVO"]
        dfFinal.drop(['index','Estado'], inplace=True, axis=1)
    dfFinal.rename(columns = {'Cierre Semana Siguiente':'Disponible'}, inplace = True)
    return dfFinal

def ajustarColumnas(df,codBodega1:str,codBodega2:str):
    if len(df.columns) == 30:
        listDfs = []
        df1 = df.iloc[:, 0:12]
        df.drop([codBodega1,'Unnamed: 4','Unnamed: 5','Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 10','Unnamed: 11','Grand Total','Unnamed: 22','Unnamed: 23','Unnamed: 24','Unnamed: 25','Unnamed: 26','Unnamed: 27','Unnamed: 28','Unnamed: 29'], inplace=True, axis=1)
        df1 = eliminarPrimeraFila(df1,1)
        df = eliminarPrimeraFila(df,1)
        df1['Archivo'] = codBodega1
        df['Archivo'] = codBodega2
        listDfs.append(df1)
        listDfs.append(df)
        return listDfs
    df = eliminarPrimeraFila(df,1)
    return df

def cerrarProgramaPorErrorEnSemana(df,añoArchivo,semanaArchivo,nombreArchivo):
    nombres_columnas = list(df.columns.values)
    if str(nombres_columnas[6]) != f'{añoArchivo}{semanaArchivo}': 
        input(f'Error en las semanas descargadas para el archivo de consumo "{nombreArchivo}" de la carpeta de año "{añoArchivo}". Verifique la base descargada.')
        exit()

# GUI of calendar
def createCalendarObject (title):
    ano = int(today.strftime("%Y"))
    mes = int(today.strftime("%m"))
    dia = int(today.strftime("%d"))
    tkobj = tk.Tk()
    tkobj.geometry("400x400")
    tkobj.title(title)
    tkc = Calendar(tkobj,selectmode = "day",year=ano,month=mes,date=dia)
    tkc.pack(pady=40)
    if title == "Fecha inicial":
        continueButton = tk.Button(tkobj, text="Seleccionar fecha", command=tkobj.destroy).pack()
    else: 
        continueButton = tk.Button(tkobj, text="Seleccionar", command=tkobj.destroy).pack()
    tkobj.mainloop()
    input = tkc.get_date()
    return input

meses = {1:'enero',
2:'febrero',
3:'marzo',
4:'abril',
5:'mayo',
6:'junio',
7: 'julio',
8: 'agosto',
9: 'septiembre',
10: 'octubre',
11: 'noviembre',
12: 'diciembre'
}
# -----------------------------
#-----------Se debe cambiar por la semana 49
año = input("Ingrese año de descarga de los archivos: ")
semanaInventario = int(input("Ingrese semana de descarga de los archivos: "))

#Verificar cómo es el otro a
semana = semanaInventario + 2
if semanaInventario == 51:
    semana = 1
if semanaInventario == 52:
    semana=2

diasSinMovimiento = int(input("¿Tomar inventario con cuántos días sin movimiento?: "))
fechaDescarga = createCalendarObject ("Fecha descargue de archivos")
fechaDescarga = datetime.strptime(fechaDescarga, "%m/%d/%y")
fechaDescarga = fechaDescarga.strftime("%d/%m/%Y")
fechaDescarga = datetime.strptime(fechaDescarga, "%d/%m/%Y")

aviso = print('Ingrese: ',end='\n'
              '1: Si se va a correr el programa normal')
print(end='\n')
aviso3 = print('2: Si se van a realizar adicionales')      
adicionales = int(input("Escriba alguna de las 2 opciones anteriores: "))

#-----------Inventario en almacenes e inventario Siesa-------------------------#
maestroProductos = get_excel_sh(site,'Agroquímicos','Diccionarios.xlsx','Maestro productos',1,'Parametro') #Diccionario de productos
maestroProductos = maestroProductos.drop_duplicates(subset=['Item'])

diccionarioFincas = get_excel_sh(site,'Agroquímicos','Diccionarios.xlsx','Fincas',1,'Parametro') #Diccionario de fincas
fincasCerradas = get_excel_sh(site,año,f'Check List - Fincas Cerradas.xlsx','Hoja1',11,f'Semana{semanaInventario}')
if semanaInventario % 2== 0:
    fincasCerradas = fincasCerradas[fincasCerradas["Semanas"] == 1]
fincasCerradas["Estado ('OK' o vacío)"] = fincasCerradas["Estado ('OK' o vacío)"].str.upper()
fincasCerradas = fincasCerradas[fincasCerradas["Estado ('OK' o vacío)"] == "OK"]
create_excel(fincasCerradas,"FincasCerradas","Hoja1")

if adicionales==1:
    inventarioSiesa = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet1',5,f'Semana{semanaInventario}')
    inventarioAlmacenes = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet',10,f'Semana{semanaInventario}')
else:
    fincasCerradas = fincasCerradas[fincasCerradas["Adicional"] == "Si"]
    inventarioSiesa = get_excel_sh(site,f'{año}',f'Semana{semanaInventario}.xlsx','Sheet1',5,f'Semana{semanaInventario}/Adicionales')
    inventarioAlmacenes = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet',10,f'Semana{semanaInventario}/Adicionales')

inventarioAlmacenes = inventarioAlmacenes.dropna(subset=['Uni'])
inventarioAlmacenes['Uni'] = inventarioAlmacenes['Uni'].apply(lambda x:eliminarEspacios(str(x)))
inventarioAlmacenes['Bodega'] = inventarioAlmacenes['Bodega'].str.split('-').str[0]
inventarioAlmacenes['Bodega'] = inventarioAlmacenes['Bodega'].apply(lambda x:eliminarEspacios(str(x)))
inventarioAlmacenes= pd.merge(inventarioAlmacenes, fincasCerradas, how='inner' ,left_on= ['Bodega'], right_on= ['Bodega'])

#Arreglar columnas de inventario en almacenes descargado de Colibrí
inventarioAlmacenes['Inventario'] = inventarioAlmacenes['Inventario'].fillna(0)
inventarioAlmacenes['Consumo Semana Actual'] = inventarioAlmacenes['Consumo Semana Actual'].fillna(0)
inventarioAlmacenes['Consumo Semana Siguiente'] = inventarioAlmacenes['Consumo Semana Siguiente'].fillna(0)
inventarioAlmacenes['Suma'] = inventarioAlmacenes['Inventario']+inventarioAlmacenes['Consumo Semana Actual'] + inventarioAlmacenes['Consumo Semana Siguiente']
inventarioAlmacenes = inventarioAlmacenes[inventarioAlmacenes["Suma"] != 0]
inventarioAlmacenes.drop(['Suma'], inplace=True, axis=1)

#Seleccionar las columnas que se van a dejar de este dataframe
diccionarioFincas['Nombre archivo']= diccionarioFincas['Nombre archivo'].str.upper()
inventarioAlmacenes= pd.merge(inventarioAlmacenes, diccionarioFincas, how='left' ,left_on= ['Bodega'], right_on= ['Bodega'])
if semanaInventario%2==0:
    inventarioAlmacenes = inventarioAlmacenes[inventarioAlmacenes["Semanas de abastecimiento"] == 1]
inventarioAlmacenes.rename(columns = {'Nombre archivo':'Finca'}, inplace = True)
inventarioAlmacenes= inventarioAlmacenes.groupby(['Bodega','Finca','Código','Agroquimico','Uni'],as_index=False)[['Inventario','Consumo Semana Actual','Cierre Semana Actual','Consumo Semana Siguiente','Disponible']].sum()
listaBodegasPrograma =  list(pd.unique(inventarioAlmacenes["Bodega"]))

#Combinar inventario en almacenes con inventario Siesa y tomar el disponible de este último
inventarioAlmacenes= pd.merge(inventarioAlmacenes, inventarioSiesa[['Item','Bodega','Cant. disponible','U.M.']], how='outer' ,left_on= ['Código','Bodega'], right_on= ['Item','Bodega'])
inventarioAlmacenes['Inventario'] = inventarioAlmacenes['Cant. disponible'].fillna(0)
inventarioAlmacenes['Código'] = inventarioAlmacenes['Código'].fillna(0)
inventarioAlmacenes['Item'] = inventarioAlmacenes['Item'].fillna(0)
inventarioAlmacenes['Código'] = np.where(inventarioAlmacenes['Código'] == 0,inventarioAlmacenes['Item'],inventarioAlmacenes['Código'])
inventarioAlmacenes.drop(['Cant. disponible'], inplace=True, axis=1)
inventarioAlmacenes['Uni'] = inventarioAlmacenes['Uni'].fillna(0)
inventarioAlmacenes['Uni'] = np.where(inventarioAlmacenes['Uni'] == 0,inventarioAlmacenes['U.M.'],inventarioAlmacenes['Uni'])
inventarioAlmacenes.drop(['Finca'], inplace=True, axis=1)
print("Procesados los archivos de inventarios en almacén")

inventarioSiesaUnicos = inventarioSiesa.drop_duplicates(subset=['Item'])
inventarioSiesaUnicos.rename(columns = {'Item':f'ItemSiesa'}, inplace = True)
inventarioAlmacenes= pd.merge(inventarioAlmacenes, inventarioSiesaUnicos[['ItemSiesa', 'Desc. item']], how='left' ,left_on= ['Código'], right_on= ['ItemSiesa'])
inventarioAlmacenes['Agroquimico'] = np.where(inventarioAlmacenes['Agroquimico'] == inventarioAlmacenes['Desc. item'],inventarioAlmacenes['Agroquimico'],inventarioAlmacenes['Desc. item'])

#-----------Consumos-------------------------#
diccionarioFincas = get_excel_sh(site,'Agroquímicos','Diccionarios.xlsx','Fincas',1,'Parametro')
consumos = obtenerConsumos(site,año,semana,semanaInventario,diccionarioFincas,"Bodega",adicionales)

if semanaInventario%2==1:
    consumos2 = obtenerConsumos(site,año,semana+1,semanaInventario,diccionarioFincas,"Bodega",adicionales)
    consumos2.rename(columns = {'Cantidad':f'Consumo Semana ({semana+1})'}, inplace = True)
    consumos2 = consumos2[consumos2["Semanas de abastecimiento"] == 2]
    consumos = pd.merge(consumos, consumos2, how='outer' ,left_on= ['Bodega','Finca','SisFinCode','Quimico','Uni','Dens'], right_on= ['Bodega','Finca','SisFinCode','Quimico','Uni','Dens'])
    consumos[f'Consumo Semana ({semana+1})'] = consumos[f'Consumo Semana ({semana+1})'].fillna(0)
    consumos['Semanas de abastecimiento'] = np.where(consumos['Semanas de abastecimiento_x'].fillna(0) == consumos['Semanas de abastecimiento_y'].fillna(0),consumos['Semanas de abastecimiento_x'],consumos['Semanas de abastecimiento_x'].fillna(0)+consumos['Semanas de abastecimiento_y'].fillna(0))
else:
    consumos = consumos[consumos["Semanas de abastecimiento"] == 1]

consumos['Uni']= consumos['Uni'].str.upper()
consumos['Cantidad'] = consumos['Cantidad'].fillna(0)
if semanaInventario%2==1: 
    consumos= consumos.groupby(['Bodega','Finca','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento'],as_index=False)[['Cantidad',f'Consumo Semana ({semana+1})']].sum()
else:
    consumos= consumos.groupby(['Bodega','Finca','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento'],as_index=False)[['Cantidad']].sum()

inventarioAlmacenes.rename(columns = {'Uni':'Uni2'}, inplace = True)
consumos = pd.merge(consumos, inventarioAlmacenes, how='outer' ,left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','Código'])
consumos = consumos[consumos['Bodega'].isin(listaBodegasPrograma)]
consumos.rename(columns = {'Cantidad':f'Consumo Semana ({semana})','Inventario':f'Inventario Semana Actual ({semanaInventario})','Consumo Semana Actual':f'Consumo Semana Actual ({semanaInventario})','Cierre Semana':f'Cierre Semana Actual ({semanaInventario})', 'Consumo Semana Siguiente':f'Consumo Semana ({semanaInventario+1})','Disponible':f'Cierre Semana ({semanaInventario+1})'}, inplace = True)
consumos['SisFinCode'] = consumos['SisFinCode'].fillna(0)
consumos['SisFinCode'] = np.where(consumos['SisFinCode'] == 0,consumos['Código'],consumos['SisFinCode'])


consumosCopiaUM = consumos.copy()
consumosCopiaUM = consumosCopiaUM.drop_duplicates(subset=['SisFinCode'])
consumosCopiaUM.rename(columns = {'Dens':'Dens2'}, inplace = True)
consumos = pd.merge(consumos, consumosCopiaUM[['SisFinCode',"Dens2"]], how='left' ,left_on= ['SisFinCode'], right_on= ['SisFinCode'])

consumos['Dens'] = np.where(consumos['Dens'] == consumos['Dens2'],consumos['Dens'],consumos['Dens2'])
consumos['Dens'] = consumos['Dens'].fillna(1)

consumosCopiaFinca = consumos.copy()
consumosCopiaFinca = consumosCopiaFinca.drop_duplicates(subset=['Bodega'])
consumosCopiaFinca.rename(columns = {'Finca':'Finca2'}, inplace = True)
consumos = pd.merge(consumos, consumosCopiaFinca[['Bodega',"Finca2"]], how='left' ,left_on= ['Bodega'], right_on= ['Bodega'])
consumos['Finca'] = np.where(consumos['Finca'] == consumos['Finca2'],consumos['Finca'],consumos['Finca2'])

consumosCopiaQuimico = consumos.copy()
consumosCopiaQuimico = consumosCopiaQuimico.drop_duplicates(subset=['SisFinCode'])
consumosCopiaQuimico.rename(columns = {'Quimico':'Quimico2'}, inplace = True)
consumos = pd.merge(consumos, consumosCopiaQuimico[['SisFinCode',"Quimico2"]], how='left' ,left_on= ['SisFinCode'], right_on= ['SisFinCode'])
consumos['Quimico'] = np.where(consumos['Quimico'] == consumos['Quimico2'],consumos['Quimico'],consumos['Quimico2'])

combinedDfCopia = inventarioAlmacenes.copy()
combinedDfCopia = combinedDfCopia.drop_duplicates(subset=['Código'])
#consumos = pd.merge(consumos, consumosCopiaQuimico[['Código',"Agroquimico"]], how='left' ,left_on= ['SisFinCode'], right_on= ['Código'])
consumos['Quimico'] = np.where(consumos['Quimico'] == consumos['Agroquimico'],consumos['Quimico'],consumos['Agroquimico'])
consumos.drop(['Dens2','Finca2','Quimico2','Agroquimico'], inplace=True, axis=1)

inventarioSiesaCopy = inventarioSiesa.copy()
inventarioSiesaCopy = inventarioSiesaCopy.drop_duplicates(subset=['Item'])
inventarioSiesaCopy.rename(columns = {'Desc. item':'Desc. item2','Item':'ItemCopia'}, inplace = True)
consumos= pd.merge(consumos, inventarioSiesaCopy[['ItemCopia', 'Desc. item2']], how='left' ,left_on= ['SisFinCode'], right_on= ['ItemCopia'])
consumos['Quimico'].fillna(consumos['Desc. item2'], inplace=True)
consumos['Uni'].fillna(consumos['Uni2'], inplace=True)
consumos.drop(['ItemCopia'], inplace=True, axis=1)

#---------------Ordenes de compras-----------------
if adicionales==1:
    ordenesDeCompras = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet1',2,f'Semana{semanaInventario}')
else:
    ordenesDeCompras = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet1',2,f'Semana{semanaInventario}/Adicionales')
valoresAFiltrar = ['Aprobado','Parcial']
ordenesDeCompras = ordenesDeCompras[ordenesDeCompras['Estado'].isin(valoresAFiltrar)]
listCropColumns = ['Desc. bodega','Desc. item','Detalle ext. 1','Detalle ext. 2','U.M. inv.']
for column in listCropColumns: ordenesDeCompras[column] = ordenesDeCompras[column].apply(lambda x:eliminarEspacios(str(x)))
ordenesDeCompras = pd.merge(ordenesDeCompras,diccionarioFincas, how='left' ,left_on= ['Bodega'], right_on= ['Bodega'])
ordenesDeCompras = ordenesDeCompras.dropna(subset=['Nombre archivo'])

calendario = get_excel_sh(siteDBLogistics,'Calendarios','calendarioSunshine.xlsx','Hoja1',3,'Parametro')#Calendario DBLogistics
ordenesDeCompras['Fecha entrega'] = pd.to_datetime(ordenesDeCompras['Fecha entrega'])
ordenesDeCompras = pd.merge(ordenesDeCompras,calendario, how='left' ,left_on= ['Fecha entrega'], right_on= ['Fecha'])
create_excel(ordenesDeCompras,"Ordenes","Hoja1")
ordenesDeCompras = ordenesDeCompras.groupby(['Nombre archivo','Bodega','Item','Desc. item','U.M. inv.','U.M.','semana'],as_index=False)[['Cant. pendiente']].sum()
ordenesDeCompras = pd.pivot_table(ordenesDeCompras,values='Cant. pendiente',columns=['semana'],index=['Nombre archivo','Item','Desc. item','U.M. inv.','U.M.','Bodega'], aggfunc=np.sum)

ordenesDeCompras['index_column'] = ordenesDeCompras.index
dfe = pd.DataFrame(ordenesDeCompras['index_column'].tolist()).fillna('').add_prefix('m')
ordenesDeCompras.drop(['index_column'], inplace=True, axis=1)
ordenesDeCompras['Key'] = range(1, len(ordenesDeCompras) + 1)

dfe['Key'] = range(1, len(dfe) + 1)
dfe = pd.merge(dfe,ordenesDeCompras,how='left' ,left_on= ['Key'], right_on= ['Key'])
dfe.drop(['Key'], inplace=True, axis=1)
dfe.rename(columns = {'m0':'Finca','m1':'Item','m2':'Desc. item','m3':'U.M. inv.','m4':'U.M.','m5':'Bodega',semanaInventario: f'Ingresos Semana ({semanaInventario})',semanaInventario+1: f'Ingresos Semana ({semanaInventario+1})',semana:f'Ingresos Semana ({semana})',semana+1:f'Ingresos Semana ({semana+1})'}, inplace = True)
#ordenesDeCompras = dfe.groupby(['Finca','Item','Desc. item','U.M. inv.'],as_index=False)[[f'Ingresos Semana ({semanaInventario})',f'Ingresos Semana ({semanaInventario+1})',f'Ingresos Semana ({semana})',f'Ingresos Semana ({semana+1})']].sum() #No se toma por el momento los ingresos de dos y tres semanas adelante porque estas pueden ocurrir al principio o final de la semana y no estar oportuno para la aplicación del agroquímico en el cultivo.
ordenesDeCompras = dfe.groupby(['Bodega','Finca','Item','Desc. item','U.M. inv.','U.M.'],as_index=False)[[f'Ingresos Semana ({semanaInventario})',f'Ingresos Semana ({semanaInventario+1})']].sum() #Evaluar con almacén si se elimina esta fila. Por el momento es el definitivo

ordenesDeCompras['U.M.'] = ordenesDeCompras['U.M.'].apply(lambda x:eliminarEspacios(str(x)))
ordenesDeCompras['Finca'] = ordenesDeCompras['Finca'].str.upper()

diccionarioUM = get_excel_sh(site,'Folder','Diccionarios.xlsx','Unidades compra',6,"Folder")
ordenesDeCompras = pd.merge(ordenesDeCompras,diccionarioUM, how='left' ,left_on= ['U.M.'], right_on= ['UM Compras'])

try:
    ordenesDeCompras[f'Ingresos Semana ({semanaInventario})'] = ordenesDeCompras[f'Ingresos Semana ({semanaInventario})'].fillna(0) * ordenesDeCompras['Factor conversión'].fillna(0)
    ordenesDeCompras[f'Ingresos Semana ({semanaInventario+1})'] = ordenesDeCompras[f'Ingresos Semana ({semanaInventario+1})'].fillna(0) * ordenesDeCompras['Factor conversión'].fillna(0)
except:
    input(f'Faltan unidades dentro del diccionario.Comuniquese con el área de investigación y desarrollo (proyectos).')

ordenesDeCompras = ordenesDeCompras.groupby(['Bodega','Finca','Item','Desc. item','U.M. inv.'],as_index=False)[[f'Ingresos Semana ({semanaInventario})',f'Ingresos Semana ({semanaInventario+1})']].sum()

#----------Traslados internos
if adicionales==1:
    trasladosInternos = get_excel_sh(site,f'{año}',f'Semana{semanaInventario}.xlsx','Sheet1',8,f'Semana{semanaInventario}')
else:
    trasladosInternos = get_excel_sh(site,f'{año}',f'Semana{semanaInventario}.xlsx','Sheet1',8,f'Semana{semanaInventario}/Adicionales')

listCropTran = ['U.M.']
for column in listCropTran: trasladosInternos[column] = trasladosInternos[column].apply(lambda x:eliminarEspacios(str(x)))

#fechaTraslados = fechaDescarga - dt.timedelta(days = 3)
#trasladosInternos = trasladosInternos.loc[(trasladosInternos['Fecha'] <= fechaTraslados)]
trasladosInternos = pd.merge(trasladosInternos,diccionarioUM, how='left' ,left_on= ['U.M.'], right_on= ['UM Compras'])

try:
    trasladosInternos[f'Ingresos Traslados Semana Actual ({semanaInventario})'] = trasladosInternos['Cant. Saldo'].fillna(0) * trasladosInternos['Factor conversión'].fillna(0)
except:
    input(f'Faltan unidades dentro del diccionario.Comuniquese con el área de investigación y desarrollo (proyectos).')
trasladosInternos= trasladosInternos.groupby(['Bod. entrada','Item'],as_index=False)[[f'Ingresos Traslados Semana Actual ({semanaInventario})']].sum()

consumos = pd.merge(consumos,trasladosInternos[['Bod. entrada','Item',f'Ingresos Traslados Semana Actual ({semanaInventario})']], how='left' ,left_on= ['Bodega','SisFinCode'], right_on= ['Bod. entrada','Item'])
consumos = pd.merge(consumos,ordenesDeCompras[['Bodega','Item','U.M. inv.',f'Ingresos Semana ({semanaInventario})',f'Ingresos Semana ({semanaInventario+1})']], how='left' ,left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','Item'])

consumos[[f"Inventario Semana Actual ({semanaInventario})"]] = consumos[[f"Inventario Semana Actual ({semanaInventario})"]].astype(float) 
consumos[[f"Consumo Semana Actual ({semanaInventario})"]] = consumos[[f"Consumo Semana Actual ({semanaInventario})"]].astype(float)
consumos[[f"Ingresos Semana ({semanaInventario})"]] = consumos[[f"Ingresos Semana ({semanaInventario})"]].astype(float)
consumos[f'Cierre Semana Actual ({semanaInventario})'] = consumos[f'Inventario Semana Actual ({semanaInventario})'].fillna(0) - consumos[f'Consumo Semana Actual ({semanaInventario})'].fillna(0) + consumos[f'Ingresos Semana ({semanaInventario})'].fillna(0) + consumos[f'Ingresos Traslados Semana Actual ({semanaInventario})'].fillna(0)
consumos[f'Cierre Semana ({semanaInventario+1})'] = consumos[f'Cierre Semana Actual ({semanaInventario})'].fillna(0) - consumos[f'Consumo Semana ({semanaInventario+1})'].fillna(0) + consumos[f'Ingresos Semana ({semanaInventario+1})'].fillna(0)
consumos[f'Cierre Semana ({semana})'] = consumos[f'Cierre Semana ({semanaInventario+1})'].fillna(0) - consumos[f'Consumo Semana ({semana})'].fillna(0)

#----Ajustar bichitos
listBichitos = [3868,4709,6602,7484,7485,7731,8056,8057,8653,10941]
consumos[f'Consumo Semana ({semana})'] = consumos[f'Consumo Semana ({semana})'].fillna(0)

if semanaInventario%2==1: 
    consumos[f'Consumo Semana ({semana+1})'] = consumos[f'Consumo Semana ({semana+1})'].fillna(0)
    consumos[f'Cierre Semana ({semana+1})'] = consumos[f'Cierre Semana ({semana})'].fillna(0) - consumos[f'Consumo Semana ({semana+1})'].fillna(0)
    consumos = consumos[['Bodega','Finca','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento',f'Inventario Semana Actual ({semanaInventario})',f'Ingresos Semana ({semanaInventario})',f'Ingresos Traslados Semana Actual ({semanaInventario})',f'Consumo Semana Actual ({semanaInventario})',f'Cierre Semana Actual ({semanaInventario})',f'Ingresos Semana ({semanaInventario+1})',f'Consumo Semana ({semanaInventario+1})',f'Cierre Semana ({semanaInventario+1})',f'Consumo Semana ({semana})',f'Cierre Semana ({semana})',f'Consumo Semana ({semana+1})',f'Cierre Semana ({semana+1})']]
    consumos['Inventario Faltante'] = np.where(consumos['Semanas de abastecimiento'] == 1,consumos[f'Cierre Semana ({semana})'],consumos[f'Cierre Semana ({semana+1})'])
    for i in listBichitos:
        consumos['Inventario Faltante'] = np.where(consumos['SisFinCode'] == i,(consumos[f'Consumo Semana ({semana})']+consumos[f'Consumo Semana ({semana+1})'])*(-1),consumos['Inventario Faltante'])
        consumos[f'Cierre Semana ({semana})'] = np.where(consumos['SisFinCode'] == i,(consumos[f'Consumo Semana ({semana})'])*(-1),consumos[f'Cierre Semana ({semana})'])
        consumos[f'Cierre Semana ({semana+1})'] = np.where(consumos['SisFinCode'] == i,(consumos[f'Consumo Semana ({semana})']+consumos[f'Consumo Semana ({semana+1})'])*(-1),consumos['Inventario Faltante'])
else:
    consumos = consumos[['Bodega','Finca','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento',f'Inventario Semana Actual ({semanaInventario})',f'Ingresos Semana ({semanaInventario})',f'Ingresos Traslados Semana Actual ({semanaInventario})',f'Consumo Semana Actual ({semanaInventario})',f'Cierre Semana Actual ({semanaInventario})',f'Ingresos Semana ({semanaInventario+1})',f'Consumo Semana ({semanaInventario+1})',f'Cierre Semana ({semanaInventario+1})',f'Consumo Semana ({semana})',f'Cierre Semana ({semana})']]
    consumos['Inventario Faltante'] = consumos[f'Cierre Semana ({semana})']
    for i in listBichitos:
        consumos['Inventario Faltante'] = np.where(consumos['SisFinCode'] == i,(consumos[f'Consumo Semana ({semana})'])*(-1),consumos['Inventario Faltante'])
        consumos[f'Cierre Semana ({semana})'] = np.where(consumos['SisFinCode'] == i,(consumos[f'Consumo Semana ({semana})'])*(-1),consumos['Inventario Faltante'])

consumos = consumos[consumos['Bodega'].isin(listaBodegasPrograma)]
diccionarioProductos = get_excel_sh(site,f'{año}','Base cotizaciones.xlsx','Base',7,f'Semana{semanaInventario}')
diccionarioProductos = diccionarioProductos.drop_duplicates(subset=['Item'])
diccionarioProductos.rename(columns = {'Desc. item':'Desc. item2'}, inplace = True)
consumos = pd.merge(consumos, diccionarioProductos[['Item',"Desc. item2"]], how='left' ,left_on= ['SisFinCode'], right_on= ['Item'])
consumos['Quimico'] = np.where(consumos['Quimico'] == consumos['Desc. item2'],consumos['Quimico'],consumos['Desc. item2'])
consumos.drop(['Item','Desc. item2'], inplace=True, axis=1)

inventarioSiesaCopia = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet1',5,f'Semana{semanaInventario}')
inventarioSiesaCopia = inventarioSiesaCopia.drop_duplicates(subset=['Item'])
inventarioSiesaCopia.rename(columns = {'Desc. item':'Desc. item3'}, inplace = True)
consumos = pd.merge(consumos, inventarioSiesaCopia[['Item',"Desc. item3"]], how='left' ,left_on= ['SisFinCode'], right_on= ['Item'])
consumos['Quimico'] = np.where(consumos['Quimico'] == consumos['Desc. item3'],consumos['Quimico'],consumos['Desc. item3'])

consumos.drop(['Item','Desc. item3'], inplace=True, axis=1)
diccionarioFincas.rename(columns = {'Semanas de abastecimiento':'Semanas de abastecimiento 2'}, inplace = True)
consumos= pd.merge(consumos, diccionarioFincas[['Bodega','Semanas de abastecimiento 2']], how='left' ,left_on= ['Bodega'], right_on= ['Bodega'])
consumos['Semanas de abastecimiento'].fillna(consumos['Semanas de abastecimiento 2'], inplace=True)
consumos[['SisFinCode']] = consumos[["SisFinCode"]].astype(float)
consumos.rename(columns = {f'Ingresos Traslados Semana Actual ({semanaInventario})':f'Ingresos Traslados Pendientes (Siesa)'}, inplace = True)
consumos= pd.merge(consumos, maestroProductos[['Item','Desc. item']], how='left' ,left_on= ['SisFinCode'], right_on= ['Item'])
consumos['Quimico'].fillna(consumos['Desc. item'], inplace=True)
consumos.drop(['Item','Desc. item','Semanas de abastecimiento 2'], inplace=True, axis=1)
create_excel(consumos,"Inventario disponible-faltante","Hoja1")

#---------------Inventario Siesa-----------------
#Filtrar los productos que ya tienen rotación mayor a X días y que no tengan consumo futuro
diaMaximoInventario = fechaDescarga - dt.timedelta(days = diasSinMovimiento)
inventarioSiesa = inventarioSiesa.groupby(['Bodega','Item','Desc. item','U.M.'],as_index=False).agg({'Fecha última salida': 'max','Fecha última entrada': 'max', 'Cant. disponible': 'sum','Costo prom. unit. (ins)':'mean'})

listCropColumnsSiesa = ['Bodega','Desc. item','U.M.']
for column in listCropColumnsSiesa: inventarioSiesa[column] = inventarioSiesa[column].apply(lambda x:eliminarEspacios(str(x)))

inventarioSiesa['Fecha último movimiento'] = np.where(inventarioSiesa['Fecha última salida']<inventarioSiesa['Fecha última entrada'],inventarioSiesa['Fecha última entrada'],inventarioSiesa['Fecha última salida'])
inventarioSiesa = inventarioSiesa.loc[(inventarioSiesa['Fecha último movimiento'] <= diaMaximoInventario)]
inventarioSiesa = pd.merge(inventarioSiesa, diccionarioFincas, how='left' ,left_on= ['Bodega'], right_on= ['Bodega'])
inventarioSiesa = inventarioSiesa[inventarioSiesa['Bodega'].isin(listaBodegasPrograma)]
consumosCopia = consumos.copy() #No tomar materiales que a pesar que no tengan movimiento mayor a 'X' días si tienen consumos futuros en esa finca

if semanaInventario%2==1: 
    consumosCopia['Consumos'] = consumosCopia[f"Consumo Semana Actual ({semanaInventario})"].fillna(0) + consumosCopia[f"Consumo Semana ({semanaInventario+1})"].fillna(0) + consumosCopia[f"Consumo Semana ({semana})"].fillna(0) + consumosCopia[f"Consumo Semana ({semana+1})"].fillna(0)
else:
    consumosCopia['Consumos'] = consumosCopia[f"Consumo Semana Actual ({semanaInventario})"].fillna(0) + consumosCopia[f"Consumo Semana ({semanaInventario+1})"].fillna(0) + consumosCopia[f"Consumo Semana ({semana})"].fillna(0)

consumosCopia = consumosCopia[consumosCopia["Consumos"] > 0]
inventarioSiesa = pd.merge(inventarioSiesa, consumosCopia[['Bodega','SisFinCode','Inventario Faltante']], how='left' ,left_on = ['Bodega','Item'], right_on = ['Bodega','SisFinCode'])
consumosCopia = consumosCopia.drop_duplicates(subset=['Bodega'])
inventarioSiesa = pd.merge(inventarioSiesa, consumosCopia[['Bodega','Finca']], how='left' ,left_on = ['Bodega'], right_on = ['Bodega'])
inventarioSiesa = inventarioSiesa[inventarioSiesa['Inventario Faltante'].isnull()]
inventarioSiesa['Nombre archivo'] = inventarioSiesa['Nombre archivo'].str.upper()
#inventarioSiesa.drop(['Código'], inplace=True, axis=1)

#---------------Generar archivo para fase 2--------
demanda = consumos.copy()
demanda = demanda[demanda['Inventario Faltante'] < 0 ]

#Falta quitar las fincas de abastecimiento a 2 cuando no le toca
productosFaltantes =  list(pd.unique(demanda["SisFinCode"]))
inventarioSiesa.rename(columns = {'Bodega':'Bodega Disponible','Finca':'Finca Disponible','Item':'SisFinCodeOferta','Costo prom. unit. (ins)':'Costo promedio unitario','Cant. disponible':'Inventario Disponible'}, inplace = True)
demandaOferta = pd.merge(demanda,inventarioSiesa[['Bodega Disponible','Finca Disponible','SisFinCodeOferta','Costo promedio unitario','Inventario Disponible','Fecha último movimiento']], how='inner' ,left_on= ['SisFinCode'], right_on= ['SisFinCodeOferta'])
demandaOferta = demandaOferta.rename(columns={'Finca':'Finca Necesidad','Bodega':'Bodega Necesidad'})
demandaOferta = demandaOferta[['Bodega Necesidad','Finca Necesidad','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Inventario Faltante','Bodega Disponible','Finca Disponible','Inventario Disponible','Fecha último movimiento','Costo promedio unitario']]
#////////////////////////////////////////////////compra de microogranismos/////////////////////////

#///////////////////////// Productos vencidos///////////////////////////////////
productosVencidos = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Base',9,f'Semana{semanaInventario}')
demandaOferta = pd.merge(demandaOferta,productosVencidos[['Bodega','Item','Cantidad vencida']], how='left' ,left_on= ['Bodega Disponible','SisFinCode'], right_on= ['Bodega','Item'])
demandaOferta['Inventario Disponible 2'] = demandaOferta['Inventario Disponible'].fillna(0)-demandaOferta['Cantidad vencida'].fillna(0)
demandaOferta = demandaOferta[demandaOferta['Inventario Disponible 2'] > 0 ]
demandaOferta = demandaOferta[['Bodega Necesidad','Finca Necesidad','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Inventario Faltante','Bodega Disponible','Finca Disponible','Inventario Disponible','Fecha último movimiento','Costo promedio unitario']]

create_excel(demandaOferta,f'OfertaDemandaSemana{semanaInventario}',"Hoja1")

#-----Producto traslado bodega IN021
if adicionales==1:
    inventarioIN021 = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet1',5,f'Semana{semanaInventario}')
else:
    inventarioIN021 = get_excel_sh(site,año,f'Semana{semanaInventario}.xlsx','Sheet1',5,f'Semana{semanaInventario}/Adicionales')
inventarioIN021 = inventarioIN021[inventarioIN021['Bodega'] == "IN021" ]
consumosNegativos = pd.merge(demanda,inventarioIN021[['Item','U.M.','Cant. disponible']],how='inner',left_on= ['SisFinCode'], right_on= ['Item'])

#-----Trayectos entre la bodega IN021 y demás fincas
fieldsTrayectos = ['ID','OrigenTray','DestinoTray','Distancia']
ajusteTrayectos = ['OrigenTray','DestinoTray']
trayectos = getList(siteMatEmpaque,'Trayecto',fieldsTrayectos,ajusteTrayectos)
trayectos = ajustarNombresFincas(trayectos,"OrigenTray")
trayectos = ajustarNombresFincas(trayectos,"DestinoTray")
trayectos = trayectos[trayectos['OrigenTray'] == "BODEGA LA PUNTA" ]
productosParaTraslado = pd.merge(consumosNegativos,trayectos[['DestinoTray','Distancia']], how='left' ,left_on= ['Finca'], right_on= ['DestinoTray'])

i=0
trasladosPunta = pd.DataFrame()
while i==0:
    productosParaTraslado = productosParaTraslado[productosParaTraslado['Cant. disponible'] > 0 ]
    minimaDistancia = productosParaTraslado.copy()
    minimaDistancia = minimaDistancia.groupby(['SisFinCode','Quimico','Uni','Dens'],as_index=False).agg({'Distancia':'min'})
    minimaDistancia['Concatenado'] = minimaDistancia["SisFinCode"].astype(str) + minimaDistancia['Quimico'] + minimaDistancia['Uni'] + minimaDistancia['Dens'].astype(str)+ minimaDistancia['Distancia'].astype(str)
    productosParaTraslado = pd.merge(productosParaTraslado,minimaDistancia, how='left' ,left_on= ['SisFinCode', 'Quimico','Uni','Dens','Distancia'], right_on= ['SisFinCode', 'Quimico','Uni','Dens','Distancia'])

    baseFinalTraslado = productosParaTraslado.dropna(subset=['Concatenado'])
    baseFinalTraslado = baseFinalTraslado.drop_duplicates(subset=['Concatenado'])
    baseFinalTraslado['Inventario Faltante Absoluto'] = np.where(baseFinalTraslado['Inventario Faltante'] <= 0, baseFinalTraslado[f'Inventario Faltante']*(-1) ,baseFinalTraslado['Inventario Faltante'])
    baseFinalTraslado['Inventario de Traslado'] = baseFinalTraslado[[f'Cant. disponible',f'Inventario Faltante Absoluto']].min(axis=1)
    baseFinalTraslado['Indice'] = 1

    #Extraer productos ya suplidos para filtrar posteriormente
    productosSuplidos = baseFinalTraslado.copy()
    productosSuplidos['Diferencia'] = productosSuplidos["Cant. disponible"] - productosSuplidos['Inventario de Traslado']
    productosSuplidos = productosSuplidos[productosSuplidos['Diferencia'] > 0.5 ]
    productosFaltantesPorSuplir =  list(pd.unique(productosSuplidos["SisFinCode"]))

    existencias =  list(pd.unique(baseFinalTraslado["Cant. disponible"]))
    if sum(existencias)==0: i=1
    baseFinalTrasladoCopia = baseFinalTraslado.copy()
    baseFinalTrasladoCopia = baseFinalTrasladoCopia[['Bodega','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Inventario Faltante','Inventario de Traslado']]

    trasladosPunta = pd.concat([trasladosPunta, baseFinalTrasladoCopia], ignore_index=True)
    productosParaTraslado = pd.merge(productosParaTraslado,baseFinalTraslado[['Bodega','SisFinCode','Indice','Inventario de Traslado']], how='left' ,left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','SisFinCode'])
    productosParaTraslado['Inventario Faltante'] = np.where(productosParaTraslado['Indice'] == 1,productosParaTraslado['Inventario Faltante'] + productosParaTraslado['Inventario de Traslado'],productosParaTraslado['Inventario Faltante'])
    productosParaTraslado['Cant. disponible'] = np.where(productosParaTraslado['Indice'] == 1,productosParaTraslado['Cant. disponible'] - productosParaTraslado['Inventario de Traslado'], productosParaTraslado['Cant. disponible'])
    inventarioActualizado = productosParaTraslado[productosParaTraslado['Indice'] == 1 ]
    inventarioActualizado.rename(columns = {'Cant. disponible':'Cant. disponible1'}, inplace = True)
    productosParaTraslado = pd.merge(productosParaTraslado,inventarioActualizado[['SisFinCode','Cant. disponible1']], how='left' ,left_on= ['SisFinCode'], right_on= ['SisFinCode'])
    
    productosParaTraslado = productosParaTraslado[productosParaTraslado['Indice'] != 1 ]
    productosParaTraslado = productosParaTraslado[productosParaTraslado['SisFinCode'].isin(productosFaltantesPorSuplir)]
    productosParaTraslado.drop(['Concatenado',"Indice",'Inventario de Traslado','Cant. disponible1'], inplace=True, axis=1)

create_excel(trasladosPunta,"Productos para traslado IN021","Hoja1")

#if len(excelesConErrores)>0:
#    print(f'Hay errores con los archivos de: {excelesConErrores}. Verifique que los códigos de bodega de la carpeta de "Inventario en almacenes" correspondan al nombre del archivo')

# Upload data to sharepoint
exit()
if adicionales==1:
    file_upload_to_sharepoint(siteDBLogistics,año,semanaInventario,f"OfertaDemandaSemana{semanaInventario}",2)
    file_upload_to_sharepoint(site,año,semanaInventario,"Inventario disponible-faltante",1)
    file_upload_to_sharepoint(site,año,semanaInventario,"Productos para traslado IN021",1)
else:
    file_upload_to_sharepoint(siteDBLogistics,f'{año}/Adicionales',semanaInventario,f"OfertaDemandaSemana{semanaInventario}",2)
    file_upload_to_sharepoint(site,año,f'{semanaInventario}/Adicionales',"Inventario disponible-faltante",1)
    file_upload_to_sharepoint(site,año,f'{semanaInventario}/Adicionales',"Productos para traslado IN021",1)

#Time
producto = f'Fase 1: Generación de archivo de sobrantes y faltantes'
elapsed_time = time() - start_time
print("Tiempo: %.10f segundos." % elapsed_time)
input(f'{producto} finalizado!... Presione enter para salir. ')