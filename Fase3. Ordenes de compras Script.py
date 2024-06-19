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
import math

today = date.today()
start_time = time()
row_limit = 5000

authcookie = Office365('https://sunshinebouquet1.sharepoint.com/', username='scastro@sunshinebouquet.com', password='CCyl3uwWUK6ZD6sf').GetCookies()
siteAprovisionamiento = Site('https://sunshinebouquet1.sharepoint.com/sites/aprovisionamiento',version=Version.v2019, authcookie=authcookie)

def get_excel_sh(site, folder1:str,folder2:str, namefile:str, sheetname:str,typeFolder:int):
#'Función para leer Excel Online Sharepoint'
    if typeFolder ==1:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Resultados traslados/{folder1}/{folder2}')
    elif typeFolder==2:
        folder = site.Folder(f'Shared%20Documents/Indicadores/{folder1}')
    else:
        folder = site.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Cotizaciones/{folder1}/{folder2}')
    df= pd.read_excel(folder.get_file(namefile), sheet_name=sheetname)
    return df

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

def create_sheet(df,namewb:str,namesh:str):
    filename = os.path.dirname(__file__)+f'\\{namewb}.xlsx'
    writer = pd.ExcelWriter(filename, engine = 'openpyxl', mode='a', if_sheet_exists ='replace')
    df.to_excel(writer, sheet_name =namesh, index=False)
    writer._save()
    writer.close

def file_upload_to_sharepoint(siteExport,folderAño:str,folderSemana:str,fileName:str):
    fileNamePath= os.path.dirname(__file__)+f'\\{fileName}.xlsx'
    folder = siteExport.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Resultados traslados/{folderAño}/{folderSemana}')
    with open(fileNamePath, mode='rb') as file:
        fileContent = file.read()
    folder.upload_file(fileContent, f'{fileName}.xlsx')

def separarProductosNoEncontradosAgotados(df,columnaFiltro,nombreExportableExcel,tipo,adicionales):
    if tipo == 1: dfFiltrado = df.dropna(subset=[columnaFiltro])
    else: dfFiltrado = df[df[columnaFiltro].isnull()]
    dfFiltrado = dfFiltrado[['Bodega','Finca','Producto','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Necesidad de compra (inv)']]
    listProductosNoEncontradosAgotados =  list(pd.unique(dfFiltrado["Producto"]))
    dfFiltrado.drop(['Producto'], inplace=True, axis=1)
    create_excel(dfFiltrado,nombreExportableExcel,"Hoja1")
    if adicionales==1:
        file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}',nombreExportableExcel)
    else:
        file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Adicionales',nombreExportableExcel)
    print(f'Productos no encontrados o agotados : {listProductosNoEncontradosAgotados}. Revisar la base de cotizaciones "Base Agroquímicos" o en la hoja "Unidades compra" del excel "Diccionario"')
    if tipo == 1: df = df[df[columnaFiltro].isnull()]
    else: df = df.dropna(subset=[columnaFiltro])
    df = df.reset_index()
    return df
    
def appendRowFromDfToAnother (df,dfExport,columnaAModificar,indiceDf,unidadesCompraXUm,varFactorConversion,varInventarioFaltantePorSuplir,varInventarioSuplido):
    varInventarioSuplido = varInventarioSuplido + (unidadesCompraXUm * varFactorConversion)
    varInventarioFaltantePorSuplir = varInventarioFaltantePorSuplir - (unidadesCompraXUm * varFactorConversion)
    df[columnaAModificar][indiceDf] = unidadesCompraXUm
    valorFiltro = df['Concatenado'][indiceDf]
    fila_a_copiar = df[df["Concatenado"] == valorFiltro]
    dfExport = pd.concat([dfExport, fila_a_copiar], ignore_index=True)
    return dfExport,varInventarioFaltantePorSuplir,varInventarioSuplido

def eliminarEspacios (x):
    x = x.strip()
    return x

def calculateColumns(df):
    df['Unidades compra decimal'] = df["Necesidad de compra (inv) UMCompras"]/ df['Factor conversión'].fillna(0)
    df['Unidades compra arriba'] = df['Unidades compra decimal'].apply(lambda x:math.ceil(x))
    df['Costo compra arriba'] = df["Unidades compra arriba"] * df['Precio Actual Compra'].fillna(0)
    df['Inventario comprado arriba'] = df["Unidades compra arriba"] * df['Factor conversión']
    df['% Inventario extra'] = (df["Inventario comprado arriba"] -  df['Necesidad de compra (inv) UMCompras']) / df['Necesidad de compra (inv) UMCompras']
    df['Unidades compra abajo'] = df['Unidades compra decimal'].astype(int)
    df['Costo compra abajo'] = df["Unidades compra abajo"] * df['Precio Actual Compra'].fillna(0)
    df['Inventario comprado abajo'] = df["Unidades compra abajo"] * df['Factor conversión']
    return df

def son_todos_iguales(lista):
    # Verificar si todos los elementos son iguales
    return all(elemento == lista[0] for elemento in lista)

def definirVariableTrue(list):
    if son_todos_iguales(list):
        variable = True
    else: 
        variable = False
    return variable

año = input("Ingrese año de extracción de archivos: ")
semanaExtraccionArchivos = int(input("Ingrese semana de extracción de archivos: "))
inventarioExtraMax = int(input("Ingrese % máximo de superación de inventario: "))/100 + 1
if semanaExtraccionArchivos==51: 
    semanaConsumo = 1
else: 
    semanaConsumo = semanaExtraccionArchivos + 2

aviso = print('Ingrese: ',end='\n'
              '1: Si se va a correr el programa normal')
print(end='\n')
aviso3 = print('2: Si se van a realizar adicionales')      
adicionales = int(input("Escriba alguna de las 2 opciones anteriores: "))
#-----------Dataframes
if adicionales==1:
    demanda = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}','Inventario disponible-faltante.xlsx','Hoja1',1)
    productosSuplidos = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}','Productos a trasladar.xlsx','Hoja1',1)
    productosSuplidosIN021 = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}','Productos para traslado IN021.xlsx','Hoja1',1)
else:
    demanda = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Adicionales','Inventario disponible-faltante.xlsx','Hoja1',1)
    productosSuplidos = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Adicionales','Productos a trasladar.xlsx','Hoja1',1)
    productosSuplidosIN021 = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Adicionales','Productos para traslado IN021.xlsx','Hoja1',1)    

productosSuplidosIN021.rename(columns = {'Inventario de Traslado':'Inventario de Traslado IN021'}, inplace = True)
baseCotizaciones = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}','Base cotizaciones.xlsx','Base',3)
diccionarioUM = get_excel_sh(siteAprovisionamiento,'Agroquímicos','','Diccionarios.xlsx','Unidades compra',2)
diccionarioFincas = get_excel_sh(siteAprovisionamiento,'Agroquímicos','','Diccionarios.xlsx','Fincas',2)
diccionarioFincas['Nombre archivo'] = diccionarioFincas['Nombre archivo'].str.upper()
spidex = get_excel_sh(siteAprovisionamiento,'Agroquímicos','','Diccionarios.xlsx','7485',2)
sueroDeLeche = get_excel_sh(siteAprovisionamiento,'Agroquímicos','','Diccionarios.xlsx','980',2)

#------------Obtener demanda final
productosSuplidos = productosSuplidos[productosSuplidos["Decisión"] == "Si"]
productosSuplidos = productosSuplidos.groupby(['Bodega Necesidad','Finca Necesidad','SisFinCode'],as_index=False).agg({'Inventario de Traslado':'sum','Inventario de Traslado (peso)':'sum'})
demanda = pd.merge(demanda,productosSuplidos[['Bodega Necesidad','Finca Necesidad','SisFinCode','Inventario de Traslado','Inventario de Traslado (peso)']],how='left' ,left_on= ['Bodega','SisFinCode'], right_on= ['Bodega Necesidad','SisFinCode'])
demanda = pd.merge(demanda,productosSuplidosIN021[['Bodega','SisFinCode','Inventario de Traslado IN021']],how='left' ,left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','SisFinCode'])

if semanaExtraccionArchivos%2==1: 
    demanda[f'Inventario Necesidad'] = np.where(demanda['Semanas de abastecimiento'] == 1,demanda[f'Cierre Semana ({semanaConsumo})'],demanda[f'Cierre Semana ({semanaConsumo+1})'])
else:
    demanda['Inventario Necesidad'] = demanda[f'Cierre Semana ({semanaConsumo})']

productosHomologos = get_excel_sh(siteAprovisionamiento,'Agroquímicos','','Diccionarios.xlsx','Productos homólogos',2)
demandaCopia = pd.merge(demanda,productosHomologos[['Item descontinuado','Item homólogo']],how='left',left_on= ['SisFinCode'], right_on= ['Item descontinuado'])
demandaCopia['SisFinCode'] = demandaCopia.apply(lambda row: row['Item homólogo'] if pd.notnull(row['Item homólogo']) else row['SisFinCode'], axis=1)
demandaCopia = demandaCopia.dropna(subset=['Item homólogo'])
demandaCopia =  demandaCopia.groupby(['Bodega','SisFinCode'],as_index=False).agg({'Inventario Necesidad':'sum'})
demandaCopia.rename(columns = {'Inventario Necesidad':'Inventario Necesidad Final'}, inplace = True)
demanda = pd.merge(demanda,demandaCopia,how='outer',left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','SisFinCode'])
demanda['Inventario Necesidad'] = demanda['Inventario Necesidad'].fillna(0)+demanda['Inventario Necesidad Final'].fillna(0)

demanda['Necesidad de compra (inv)'] = demanda[f"Inventario Necesidad"] + demanda['Inventario de Traslado'].fillna(0) + demanda['Inventario de Traslado IN021'].fillna(0)
demanda = demanda[demanda[f"Necesidad de compra (inv)"] < 0]
demanda['Quimico'] = demanda['Quimico'].str.upper()
listCropColumns = ['Quimico']
for column in listCropColumns: demanda[column] = demanda[column].apply(lambda x:eliminarEspacios(str(x)))

#-------------Base de proveedores cruzado con unidades de compra
baseCotizaciones['Observaciones'] = baseCotizaciones['Observaciones'].fillna(" ")
baseCotizaciones['Autorizado'] = baseCotizaciones['Autorizado'].fillna(" ")
baseCotizacionesFiltrada = baseCotizaciones[baseCotizaciones["Autorizado"] == "Si"]
baseCotizacionesFiltrada = baseCotizacionesFiltrada[baseCotizacionesFiltrada["Agotado"].isnull()]

baseCotizacionesFiltrada =  baseCotizacionesFiltrada.groupby(['Item','U.M.','Observaciones','Autorizado'],as_index=False).agg({'Precio Actual Compra':'min'})
baseCotizacionesFiltrada.rename(columns = {'Observaciones':'Observaciones borrar','Autorizado':'Autorizado borrar'}, inplace = True)
baseCotizacionesFiltrada = pd.merge(baseCotizacionesFiltrada,baseCotizaciones,how='left',left_on= ['Item','U.M.','Precio Actual Compra','Observaciones borrar','Autorizado borrar'], right_on= ['Item','U.M.','Precio Actual Compra','Observaciones','Autorizado'])
baseCotizacionesFiltrada['Observaciones'] = baseCotizacionesFiltrada['Observaciones'].replace(' ','0')
baseCotizacionesFiltrada['Autorizado'] = baseCotizacionesFiltrada['Autorizado'].replace(' ','0')
baseCotizacionesFiltrada.drop(['Observaciones borrar','Autorizado borrar'], inplace=True, axis=1)

baseCotizaciones = pd.merge(baseCotizacionesFiltrada,diccionarioUM,how='left',left_on= ['U.M.'], right_on= ['UM Compras'])
baseCotizaciones['Concatenado'] = baseCotizaciones["Item"].astype(str) + baseCotizaciones['U.M.'] + baseCotizaciones['Razón social proveedor'] + baseCotizaciones['Precio Actual Compra'].astype(str)
baseFiltradaEsmeralda = baseCotizaciones[baseCotizaciones["Observaciones"] == "IN080"]
listProductosExclusivosEsmeralda =  list(pd.unique(baseFiltradaEsmeralda["Item"]))

#Filtrar las observaciones que no son para la finca Esmeralda Med
baseFiltradaDemasFincas = baseCotizaciones[baseCotizaciones["Observaciones"] == "Demás fincas"]
listProductosExclusivosDemasFincas =  list(pd.unique(baseFiltradaDemasFincas["Item"]))

valoresAFiltrarNoEsm = ['De 1 a 10 bultos', 'Mayor a 10 bultos']
baseFiltradaNoEsmeralda = baseCotizaciones[baseCotizaciones['Observaciones'].isin(valoresAFiltrarNoEsm)]
listProductosExclusivosNoEsmeralda =  list(pd.unique(baseFiltradaNoEsmeralda["Item"]))

baseCotizacionesValoresUnicos = baseCotizaciones.drop_duplicates(subset=['Item'])
demanda = pd.merge(demanda,baseCotizacionesValoresUnicos[['Item', 'UM Inv']],how='left' ,left_on= ['SisFinCode'], right_on= ['Item'])
demanda['Necesidad de compra (inv)'] = np.where(demanda['Necesidad de compra (inv)'] <= 0, demanda['Necesidad de compra (inv)']*(-1) ,demanda['Necesidad de compra (inv)']) #No sirvió calculando con -1
demanda['Producto'] = demanda["SisFinCode"].astype(str) + demanda['Quimico']

#------------Productos agotados y no encontrados
demanda = separarProductosNoEncontradosAgotados(demanda,"UM Inv","Productos no encontrados",2,adicionales)
demanda['Necesidad de compra (inv) UMCompras'] = np.where(demanda['Uni'] != demanda['UM Inv'], demanda['Necesidad de compra (inv)']*demanda['Dens'] ,demanda['Necesidad de compra (inv)'])
demanda = demanda.sort_values(by = ['Bodega','SisFinCode'], ascending = [True,True],ignore_index=True )

#Separar productos de compra urgente
listBichitos = [3868,4709,6602,7484,7485,7731,8056,8057,8653,10941]
baseAdicionales = demanda[~demanda['SisFinCode'].isin(listBichitos)]
baseAdicionales['Necesidad de compra (inv)'] = demanda[f'Cierre Semana ({semanaExtraccionArchivos+1})'] + demanda['Inventario de Traslado'].fillna(0) + demanda['Inventario de Traslado IN021'].fillna(0)
baseAdicionales = baseAdicionales[baseAdicionales["Necesidad de compra (inv)"] < 0]
baseAdicionales['Necesidad de compra (inv)'] = np.where(baseAdicionales['Necesidad de compra (inv)'] <= 0, baseAdicionales['Necesidad de compra (inv)']*(-1),'Necesidad de compra (inv)') #No sirvió calculando con -1
baseAdicionales['Necesidad de compra (inv)'] = baseAdicionales['Necesidad de compra (inv)'].astype(float)
baseAdicionales['Necesidad de compra (inv) UMCompras'] = np.where(baseAdicionales['Uni'] != baseAdicionales['UM Inv'], baseAdicionales['Necesidad de compra (inv)']*baseAdicionales['Dens'],baseAdicionales['Necesidad de compra (inv)'])
baseAdicionales = baseAdicionales.sort_values(by = ['Bodega','SisFinCode'], ascending = [True,True],ignore_index=True )

def logicaCompra(dataframe):
    dataframe = dataframe.reset_index()
    necesidadCompraFinal = pd.DataFrame()
    indice = 0
    while indice<len(dataframe):
        inventarioTotal = dataframe['Necesidad de compra (inv) UMCompras'][indice]
        inventarioFaltantePorSuplir = inventarioTotal.copy()
        bodega = dataframe['Bodega'][indice]
        item = dataframe['SisFinCode'][indice]
        baseCotizacionesFiltrado = baseCotizaciones[baseCotizaciones["Item"] == item]
        baseCotizacionesFiltrado['Bodega'] = bodega
        baseCotizacionesFiltrado['Unidades de compra'] = 0
        baseCotizacionesFiltrado['Necesidad inventario'] = inventarioTotal.copy()
        baseCotizacionesFiltrado['ColumnaOrden'] = baseCotizacionesFiltrado['Precio Actual Compra'] / baseCotizacionesFiltrado['Factor conversión']    
        #Casos especiales (observaciones)
        if bodega == "IN080":
            if item in listProductosExclusivosEsmeralda: baseCotizacionesFiltrado = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Observaciones"] == bodega]
        else:
            if item in listProductosExclusivosDemasFincas: 
                baseCotizacionesFiltrado = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Observaciones"] == 'Demás fincas']
            elif item in listProductosExclusivosNoEsmeralda:
                baseCotizacionesFiltrado = baseCotizacionesFiltrado[baseCotizacionesFiltrado['Observaciones'].isin(valoresAFiltrarNoEsm)]
                #Agregar casos en el caso que se presenten mayores o menores bultos
                if item == 995:
                    if inventarioTotal>400000:
                        baseCotizacionesFiltrado = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Observaciones"] == 'Mayor a 10 bultos']
                    else:
                        baseCotizacionesFiltrado = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Observaciones"] == 'De 1 a 10 bultos']
                else:
                    if inventarioTotal>500000:
                        baseCotizacionesFiltrado = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Observaciones"] == 'Mayor a 10 bultos']
                    else:
                        baseCotizacionesFiltrado = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Observaciones"] == 'De 1 a 10 bultos']
    
        if item == 7485:
            baseCotizacionesFiltrado = pd.merge(baseCotizacionesFiltrado,spidex[['Bodega', 'Precio']],how='left',left_on= ['Bodega'], right_on= ['Bodega'])
            baseCotizacionesFiltrado['Precio Actual Compra'] = baseCotizacionesFiltrado['Precio']
            baseCotizacionesFiltrado.drop(['Precio'], inplace=True, axis=1)

        if item == 980:
            baseCotizacionesFiltrado = pd.merge(baseCotizacionesFiltrado,sueroDeLeche,how='left',left_on= ['Bodega'], right_on= ['Bodega'])
            baseCotizacionesFiltrado['Observaciones'] = baseCotizacionesFiltrado['Tipo de compra']
            baseCotizacionesFiltrado.drop(['Tipo de compra'], inplace=True, axis=1)        

        baseCotizacionesFiltrado['Inventario abastecido'] =  baseCotizacionesFiltrado.apply(lambda row: math.ceil(row['Necesidad inventario'] / row['Factor conversión']), axis=1) * baseCotizacionesFiltrado['Factor conversión']
        baseCotizacionesFiltradoCopia = baseCotizacionesFiltrado.copy()
        baseCotizacionesFiltradoCopia = baseCotizacionesFiltradoCopia.groupby(['Inventario abastecido'],as_index=False).agg({'ColumnaOrden':'min'})#Si dos unidades suplen lo mismo dejar la de costo más económico
        
        if len(baseCotizacionesFiltrado)<=2:
            baseCotizacionesFiltrado = pd.merge(baseCotizacionesFiltrado,baseCotizacionesFiltradoCopia,how='inner',left_on= ['Inventario abastecido','ColumnaOrden'], right_on= ['Inventario abastecido','ColumnaOrden'])
        baseCotizacionesFiltrado = baseCotizacionesFiltrado.reset_index()
        observacion = baseCotizacionesFiltrado['Observaciones'][0]
        indice2 = 0
        inventarioSuplido = 0
        baseCotizacionesFiltrado = baseCotizacionesFiltrado.sort_values(by = ['ColumnaOrden'], ascending = [True],ignore_index=True )

        while indice2<len(baseCotizacionesFiltrado): 
            if inventarioSuplido>=inventarioTotal or inventarioFaltantePorSuplir == 0:
                break
            factorConversion = baseCotizacionesFiltrado['Factor conversión'][indice2]
            unidadesCompra = inventarioFaltantePorSuplir/factorConversion
            unidadesCompraMin = int(unidadesCompra)
            unidadesCompraMax = math.ceil(unidadesCompra)
            if item==980:
                if observacion == "Mínimo 100 y múltiplos de 20":
                    if inventarioTotal<100000: 
                        unidadesCompraMin = 100
                        unidadesCompraMax = 100
                    else:
                        unidadesCompraMin = math.ceil(inventarioTotal/20000)*20
                        unidadesCompraMax = math.ceil(inventarioTotal/20000)*20
                else:
                    if inventarioTotal<200000: 
                        unidadesCompraMin = 200
                        unidadesCompraMax = 200
                    else:
                        unidadesCompraMin = math.ceil(inventarioTotal/20000)*20
                        unidadesCompraMax = math.ceil(inventarioTotal/20000)*20
            if observacion == "Múltiplos de 500":
                if inventarioTotal < 1000000: 
                    unidadesCompraMin = 1000
                    unidadesCompraMax = 1000
                else:
                    unidadesCompraMin = math.ceil(inventarioTotal/500000)*500
                    unidadesCompraMax = math.ceil(inventarioTotal/500000)*500
            if observacion == "Múltiplos de 25":
                unidadesCompraMin = math.ceil(inventarioTotal/25000)*25
                unidadesCompraMax = math.ceil(inventarioTotal/25000)*25
            if observacion == "Múltiplos de 10":
                unidadesCompraMin = math.ceil(inventarioTotal/10000)*10
                unidadesCompraMax = math.ceil(inventarioTotal/10000)*10

            if observacion == "Múltiplos de 20":
                unidadesCompraMin = math.ceil(inventarioTotal/20000)*20
                unidadesCompraMax = math.ceil(inventarioTotal/20000)*20      

            inventarioCompraMax = unidadesCompraMax * factorConversion
            inventarioCompraMin = unidadesCompraMin * factorConversion
            inventarioASuplir = inventarioSuplido + inventarioCompraMax
            sobreAbastecimiento = inventarioASuplir/inventarioTotal

            if len(baseCotizacionesFiltrado)==1:
                necesidadCompraFinal,inventarioFaltantePorSuplir,inventarioSuplido = appendRowFromDfToAnother(baseCotizacionesFiltrado,necesidadCompraFinal,'Unidades de compra',
                0,unidadesCompraMax,factorConversion,inventarioFaltantePorSuplir,inventarioSuplido)
                break
            if sobreAbastecimiento <= inventarioExtraMax:
                necesidadCompraFinal,inventarioFaltantePorSuplir,inventarioSuplido = appendRowFromDfToAnother(baseCotizacionesFiltrado,necesidadCompraFinal,'Unidades de compra',
                indice2,unidadesCompraMax,factorConversion,inventarioFaltantePorSuplir,inventarioSuplido)
                break
            if indice2 == (len(baseCotizacionesFiltrado)-1):
                baseCotizacionesFiltradoCopia = baseCotizacionesFiltrado.copy()
                #/////////////////// Agregar agrupacion por minimo unidades que suplen lo mismo /////////////////////////
                agrupoporminimos = baseCotizacionesFiltradoCopia.groupby(['Item','Bodega', 'Inventario abastecido'])['ColumnaOrden'].min().reset_index()
                baseCotizacionesFiltradoCopia2 = pd.merge(baseCotizacionesFiltradoCopia,agrupoporminimos[['Item', 'Bodega', 'ColumnaOrden','Inventario abastecido']], how='inner', left_on=['Item', 'Bodega', 'ColumnaOrden'], right_on=['Item', 'Bodega', 'ColumnaOrden'])
                factorConversion = baseCotizacionesFiltradoCopia2['Factor conversión'].min()
                baseCotizacionesFiltradoMin = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Factor conversión"] == factorConversion]
                factorConversion = baseCotizacionesFiltradoCopia['Factor conversión'].min()
                baseCotizacionesFiltradoMin = baseCotizacionesFiltrado[baseCotizacionesFiltrado["Factor conversión"] == factorConversion]
                baseCotizacionesFiltradoMin = baseCotizacionesFiltradoMin.reset_index()
                unidadesCompraMax = math.ceil(inventarioFaltantePorSuplir/factorConversion)
                necesidadCompraFinal,inventarioFaltantePorSuplir,inventarioSuplido = appendRowFromDfToAnother(baseCotizacionesFiltradoMin,necesidadCompraFinal,'Unidades de compra',
                0,unidadesCompraMax,factorConversion,inventarioFaltantePorSuplir,inventarioSuplido)
                break
            if unidadesCompraMin>0:
                necesidadCompraFinal,inventarioFaltantePorSuplir,inventarioSuplido = appendRowFromDfToAnother(baseCotizacionesFiltrado,necesidadCompraFinal,'Unidades de compra',
                indice2,unidadesCompraMin,factorConversion,inventarioFaltantePorSuplir,inventarioSuplido)
            indice2+=1
        indice+=1
        
    dataframe = pd.merge(dataframe,necesidadCompraFinal,how='left',left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','Item'])
    dataframe['Costo compra total'] = dataframe['Precio Actual Compra'] * dataframe['Unidades de compra']
    dataframe['Inventario suplido'] = dataframe['Factor conversión'] * dataframe['Unidades de compra']
    dataframe['SisFinCodeText'] = dataframe['SisFinCode'].astype(str)
    dataframe['ConcatenadoBodegaProducto'] = dataframe['Bodega'] + dataframe['SisFinCodeText']

    #Dejar solo una UM cuando dos unidades suplen el mismo inventario
    indice = 1
    concatenadoBodegaProductoAnterior = dataframe['ConcatenadoBodegaProducto'][0]
    factorConversionAnterior = dataframe['Factor conversión'][0]
    unidadesCompraAnterior = dataframe['Unidades de compra'][0]
    while indice<len(dataframe):
        concatenadoBodegaProducto = dataframe['ConcatenadoBodegaProducto'][indice]
        inventarioSuplido = dataframe['Inventario suplido'][indice]
        if concatenadoBodegaProducto == concatenadoBodegaProductoAnterior:
            if factorConversionAnterior == inventarioSuplido:
                dataframe['Unidades de compra'][indice-1] = unidadesCompraAnterior + 1
                dataframe['Unidades de compra'][indice] = 0
        factorConversionAnterior = dataframe['Factor conversión'][indice]
        unidadesCompraAnterior = dataframe['Unidades de compra'][indice]
        concatenadoBodegaProductoAnterior = concatenadoBodegaProducto
        indice+=1

    dataframe = dataframe[dataframe["Unidades de compra"] != 0]
    dataframe['Inventario suplido'] = dataframe['Factor conversión'] * dataframe['Unidades de compra']
    dataframe['Inventario sobrante'] = dataframe['Unidades de compra'] * dataframe['Factor conversión'] - dataframe['Necesidad de compra (inv)'] #Calcularlas mejor
    dataframe['% Inventario sobrante'] = dataframe['Inventario sobrante'] / dataframe['Necesidad de compra (inv)'] #Calcularlas mejor
    return dataframe

baseAdicionales = logicaCompra(baseAdicionales)
baseAdicionales['Inventario suplido real adicionales'] = np.where(baseAdicionales['Uni'] != baseAdicionales['UM Inv_x'], baseAdicionales['Inventario suplido']/baseAdicionales['Dens'],baseAdicionales['Inventario suplido'])
baseAdicionalesAgrupada =  baseAdicionales.groupby(['Bodega','SisFinCode'],as_index=False).agg({'Inventario suplido real adicionales':'sum'})
demanda = pd.merge(demanda,baseAdicionalesAgrupada[['Bodega','SisFinCode','Inventario suplido real adicionales']],how='left',left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','SisFinCode'])
demanda['Necesidad de compra (inv)'] = demanda['Necesidad de compra (inv)'].fillna(0) - demanda['Inventario suplido real adicionales'].fillna(0)
demanda = demanda[demanda["Necesidad de compra (inv)"] > 0]
demanda['Uni'] = demanda['Uni'].fillna(demanda['UM Inv'])
demanda['Necesidad de compra (inv) UMCompras'] = np.where(demanda['Uni'] != demanda['UM Inv'], demanda['Necesidad de compra (inv)']*demanda['Dens'],demanda['Necesidad de compra (inv)'])
demanda = logicaCompra(demanda)

#------------Productos con observaciones
demandaProductosConObservaciones = demanda.copy()
demandaProductosConObservaciones['Observaciones'] = demandaProductosConObservaciones['Observaciones'].astype(str)
demandaProductosConObservaciones = demandaProductosConObservaciones[demandaProductosConObservaciones["Observaciones"] != "0"]
demandaProductosConObservaciones = demandaProductosConObservaciones.groupby(['Bodega','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Necesidad de compra (inv)','Razón social proveedor','UM Compras','Descripción UMCompras','Precio Actual Compra','Factor conversión','Observaciones'],as_index=False).agg({'Unidades de compra':'sum','Costo compra total':'sum','Inventario suplido':'sum'})
demandaProductosConObservaciones = demandaProductosConObservaciones[['Bodega','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Necesidad de compra (inv)','Razón social proveedor','UM Compras','Descripción UMCompras','Precio Actual Compra','Factor conversión','Unidades de compra','Costo compra total','Inventario suplido','Observaciones']]
create_excel(demandaProductosConObservaciones,"Productos con observaciones","Hoja1")
if adicionales==1:
    file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}',"Productos con observaciones")
else:
    file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Adicionales',"Productos con observaciones")

#------------Definir el 70-30% para la compra de melaza (verificar plan comercial en el futuro)
melaza = demanda.copy()
melaza = melaza[melaza["SisFinCode"] == 1027]
melaza = melaza.sort_values(by = ['Unidades de compra'], ascending = [False],ignore_index=True )
baseCotizacionesMelaza = get_excel_sh(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}','Base cotizaciones.xlsx','Base',3)
baseCotizacionesMelaza['Observaciones'] = baseCotizacionesMelaza['Observaciones'].fillna(" ")
baseCotizacionesMelaza['Autorizado'] = baseCotizacionesMelaza['Autorizado'].fillna(" ")
baseCotizacionesMelaza = baseCotizacionesMelaza[baseCotizacionesMelaza["Autorizado"] == "Si"]
baseCotizacionesMelaza = baseCotizacionesMelaza[baseCotizacionesMelaza["Agotado"].isnull()]
baseCotizacionesMelaza = baseCotizacionesMelaza[baseCotizacionesMelaza["Item"] == 1027]
print(baseCotizacionesMelaza)
 
indice = 0
inventarioSuplido = 0
inventarioTotalMelaza =  melaza['Unidades de compra'].sum()
treintaInventarioTotal = inventarioTotalMelaza * 0.3
baseCotizacionesMelaza = baseCotizacionesMelaza.reset_index()
melaza = melaza.reset_index()
while indice<len(melaza):
    if inventarioSuplido>treintaInventarioTotal:
        melaza['Razón social proveedor'][indice] = baseCotizacionesMelaza['Razón social proveedor'][0]  
    inventarioSuplido = inventarioSuplido + melaza['Unidades de compra'][indice]
    indice+=1
 
baseCotizacionesMelaza.rename(columns = {'Precio Actual Compra':'Precio Actual Compra 2'}, inplace = True)
melaza = pd.merge(melaza,baseCotizacionesMelaza[['Razón social proveedor','Precio Actual Compra 2']],how='left',left_on= ['Razón social proveedor'], right_on= ['Razón social proveedor'])
melaza['Precio Actual Compra'] = np.where(melaza['Precio Actual Compra'] == melaza['Precio Actual Compra 2'],melaza['Precio Actual Compra'],melaza['Precio Actual Compra 2'])
 
melaza.rename(columns = {'Razón social proveedor':'Razón social proveedor 2','Precio Actual Compra':'Precio Actual Compra 3'}, inplace = True)
demanda = pd.merge(demanda,melaza[['Bodega','SisFinCode','Razón social proveedor 2','Precio Actual Compra 3']],how='left',left_on= ['Bodega','SisFinCode'], right_on= ['Bodega','SisFinCode'])

demanda['Razón social proveedor 2'] = demanda['Razón social proveedor 2'].fillna(0)
demanda['Precio Actual Compra 3'] = demanda['Precio Actual Compra 3'].fillna(0)
demanda['Razón social proveedor'] = np.where(demanda['Razón social proveedor 2'] != 0,demanda['Razón social proveedor 2'],demanda['Razón social proveedor'])
demanda['Precio Actual Compra'] = np.where(demanda['Precio Actual Compra 3'] != 0,demanda['Precio Actual Compra 3'],demanda['Precio Actual Compra'])

create_excel(demanda,"QuePasoMelazita","Hoja1")
exit()

#Dejar columnas de dataframes finales a exportar
def definirColumnasDataframeOrdenes(dataframe):
    columnasAgrupar = ['Bodega','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Necesidad de compra (inv)','Necesidad inventario','Razón social proveedor','UM Compras','Descripción UMCompras','Precio Actual Compra','Factor conversión','Concatenado']
    for i in columnasAgrupar:
        dataframe[i] = dataframe[i].fillna(0)
    dataframe = dataframe.groupby(['Bodega','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Necesidad de compra (inv)','Necesidad inventario','Razón social proveedor','UM Compras','Descripción UMCompras','Precio Actual Compra','Factor conversión','Concatenado'],as_index=False).agg({'Unidades de compra':'sum','Costo compra total':'sum','Inventario suplido':'sum'})
    dataframe = dataframe[['Bodega','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Necesidad de compra (inv)','Necesidad inventario','Razón social proveedor','UM Compras','Descripción UMCompras','Precio Actual Compra','Factor conversión','Unidades de compra','Costo compra total','Inventario suplido']]
    dataframe = dataframe.sort_values(by = ['Bodega','SisFinCode','Factor conversión'], ascending = [True,True,False],ignore_index=True )
    dataframe = pd.merge(dataframe,diccionarioFincas[['Bodega','Asignación','Bodega - Descripción']],how='left',left_on= ['Bodega'], right_on= ['Bodega'])
    return dataframe

demanda = definirColumnasDataframeOrdenes(demanda)
baseAdicionales = definirColumnasDataframeOrdenes(baseAdicionales)
create_excel(baseAdicionales,"Adicionales","Hoja1")
if adicionales==1:
    file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Productos adicionales semana {semanaExtraccionArchivos+1}','Adicionales')
else:
    file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Adicionales/Productos adicionales semana {semanaExtraccionArchivos+1}','Adicionales')

#Separar dataframe según comprador
demandaAndrea = demanda[demanda["Asignación"] == "Andrea Navarrete"]
demandaClaudia = demanda[demanda["Asignación"] == "Claudia Quiroga"]
demandaSandro = demanda[demanda["Asignación"] == "Sandro Murillo"]

listBodegasAndrea =  list(pd.unique(demandaAndrea["Bodega - Descripción"]))
listBodegasClaudia =  list(pd.unique(demandaClaudia["Bodega - Descripción"]))
listBodegasSandro =  list(pd.unique(demandaSandro["Bodega - Descripción"]))
listBodegas = []
listBodegas.append(listBodegasAndrea)
listBodegas.append(listBodegasClaudia)
listBodegas.append(listBodegasSandro)

for i in listBodegas:
    for j in i:
        demandaFiltradaPorBodega = demanda[demanda["Bodega - Descripción"] == j]
        demandaFiltradaPorBodega = demandaFiltradaPorBodega.reset_index()
        asignacion = demandaFiltradaPorBodega["Asignación"][0]
        demandaFiltradaPorBodega = demandaFiltradaPorBodega[['Bodega','SisFinCode','Quimico','Uni','Dens','Semanas de abastecimiento','Necesidad de compra (inv)','Necesidad inventario','Razón social proveedor','UM Compras','Descripción UMCompras','Precio Actual Compra','Factor conversión','Unidades de compra','Costo compra total','Inventario suplido']]
        create_excel(demandaFiltradaPorBodega,j,"Hoja1")
        if adicionales==1:
            file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/{asignacion}',j)
        else:
            file_upload_to_sharepoint(siteAprovisionamiento,año,f'Semana{semanaExtraccionArchivos}/Adicionales/{asignacion}',j)
#Time
producto = f'Fase 3: Generación de ordenes de compras'
elapsed_time = time() - start_time
print("Tiempo: %.10f segundos." % elapsed_time)
input(f'{producto} finalizado!... Presione enter para salir. ')