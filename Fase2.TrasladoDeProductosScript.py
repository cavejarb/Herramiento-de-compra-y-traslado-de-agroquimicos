import math
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
from math import radians, sin, cos, sqrt, atan2
import itertools

today = date.today()
start_time = time()
row_limit = 5000

authcookie = Office365('https://sunshinebouquet1.sharepoint.com/', username='scastro@sunshinebouquet.com', password='CCyl3uwWUK6ZD6sf').GetCookies()
siteAprovisionamiento = Site('https://sunshinebouquet1.sharepoint.com/sites/aprovisionamiento',version=Version.v2019, authcookie=authcookie)
siteDBLogistics = Site('https://sunshinebouquet1.sharepoint.com/sites/CosteodeTransporte',version=Version.v2019, authcookie=authcookie)
siteMatEmpaque = Site('https://sunshinebouquet1.sharepoint.com/sites/MatEmpaque',version=Version.v2019, authcookie=authcookie)

print("Funciona perra")


def get_excel_sh(site, folder1:str,folder2:str, namefile:str, sheetname:str,typeFolder:int):
#'Función para leer Excel Online Sharepoint'
    if typeFolder==1:
        folder = site.Folder(f'Documentos%20compartidos/{folder1}/{folder2}')
    else:
        folder = site.Folder(f'Shared%20Documents/{folder1}/{folder2}')
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

def file_upload_to_sharepoint(siteExport,folderAño:str,folderSemana:str,fileName:str):
    fileNamePath= os.path.dirname(__file__)+f'\\{fileName}.xlsx'
    folder = siteExport.Folder(f'Shared%20Documents/Indicadores/Agroquímicos/Resultados traslados/{folderAño}/{folderSemana}')
    with open(fileNamePath, mode='rb') as file:
        fileContent = file.read()
    folder.upload_file(fileContent, f'{fileName}.xlsx')

def create_sheet(df,namewb:str,namesh:str):
    filename = os.path.dirname(__file__)+f'\\{namewb}.xlsx'
    writer = pd.ExcelWriter(filename, engine = 'openpyxl', mode='a', if_sheet_exists ='replace')
    df.to_excel(writer, sheet_name =namesh, index=False)
    writer._save()
    writer.close

#Getting lists of sharepoint
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

def getListTarifa(site):
    listAux = site.List('Tarifa')
    data = listAux.GetListItems()
    tarifa = pd.DataFrame(data[0:])
    return tarifa

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

def tsp_branch_and_bound_no_return(graph, start_node):
    # Número de nodos en el grafo
    n = len(graph)
    # Función auxiliar para calcular el costo de una ruta parcial
    def partial_cost(path):
        cost = 0
        for i in range(len(path) - 1):
            cost += graph[path[i]][path[i+1]]
        return cost
    # Inicialización de variables
    best_path = None
    best_cost = float('inf')
    queue = [(0, [start_node])]  # Cola de nodos por explorar: (costo acumulado, ruta parcial)
    while queue:
        # Obtener el nodo de la cola con el menor costo acumulado
        cost, path = queue.pop(0)
        # Si la ruta parcial ya visita todos los nodos, actualizar la mejor solución
        if len(path) == n:
            if cost < best_cost:
                best_cost = cost
                best_path = path
            continue
        
        # Ramificar y podar: agregar a la cola rutas parciales prometedoras
        for next_node in range(n):
            if next_node not in path:
                new_cost = cost + graph[path[-1]][next_node]
                if new_cost < best_cost:
                    new_path = path + [next_node]
                    queue.append((new_cost, new_path))
    
    return best_path, best_cost

def filtrarDataframe(df,list,contador,columnaFincaDestino):
    des = list[contador+1]
    destino = df[df[columnaFincaDestino] == des]
    destino = destino.reset_index(drop=True)
    return destino

def recorrerListaYSumarKmCc(df,recorrido,columnaOrigen,columnaDestino,distanciaInventario,contador,ccKm,kmCcTotales):
    origen = df[df[columnaOrigen] == recorrido[0]]
    while contador<len(recorrido)-1:
        if ccKm==2:
            origen = df[df[columnaOrigen] == recorrido[contador]]
        destino = filtrarDataframe(origen,recorrido,contador,columnaDestino)
        kmCcTotales = kmCcTotales + destino[distanciaInventario][0]
        contador+=1
    return kmCcTotales

def getKmCc(recorrido,df,columnaOrigen:str,columnaDestino:str,distanciaInventario:str,ccKm,recorridoAnterior = []):
    contador = 0
    kmCcTotales = 0
    if ccKm == 1 or ccKm ==2:
        kmCcTotales = recorrerListaYSumarKmCc(df,recorrido,columnaOrigen,columnaDestino,distanciaInventario,contador,ccKm,kmCcTotales)
    else:
        try:
            kmCcTotales = recorrerListaYSumarKmCc(df,recorrido,columnaOrigen,columnaDestino,distanciaInventario,contador,ccKm,kmCcTotales)
        except:
            recorridoCopia = recorrido.copy()
            recorridoAnteriorCopia = recorridoAnterior.copy()
            for i in range(1, len(recorridoCopia)):
                recorridoAnteriorCopia.append(recorridoCopia[i])
            origen = df[df[columnaOrigen] == recorridoAnteriorCopia[0]]
            while contador<len(recorridoAnteriorCopia)-1:
                destino = filtrarDataframe(origen,recorridoAnteriorCopia,contador,columnaDestino)
                kmCcTotales = kmCcTotales + destino[distanciaInventario][0]
                contador+=1
    return kmCcTotales

def restarKmCC(recorrido,df,columnaOrigen:str,columnaDestino:str,distanciaInventario:str,tipo,ccKmTotales,capacidad):
    listaARetornar = []
    #Dividir trayectos según la cantidad de cc o km
    ccKm = 0
    while ccKmTotales>1:
        trayectosDivididos = []   
        trayectosDivididos.append(recorrido[0])
        ccKm = 0
        contador = 0
        origen = df[df[columnaOrigen] == recorrido[0]]
        while (contador<len(recorrido)-1):
            if tipo ==2:
                origen = df[df[columnaOrigen] == recorrido[contador]]
            destino = filtrarDataframe(origen,recorrido,contador,columnaDestino)
            if ccKm + destino[distanciaInventario][0]>=capacidad:
                break
            ccKm = ccKm + destino[distanciaInventario][0]
            trayectosDivididos.append(recorrido[contador+1])
            contador+=1
        listaARetornar.append(trayectosDivididos)
        recorridoCopia = []
        if tipo==1: recorridoCopia.append(recorrido[0])
        else: recorridoCopia.append(recorrido[contador])
        contador+=1
        for i in range(contador,len(recorrido)):
            recorridoCopia.append(recorrido[i])
        recorrido = recorridoCopia
        ccKmTotales = ccKmTotales-ccKm
    return listaARetornar,ccKmTotales
    
def haversine(lat1, lon1, lat2, lon2):
    # Radio de la Tierra en kilómetros
    R = 6371.0
    # Convierte las coordenadas de grados a radianes
    lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
    # Diferencias de latitud y longitud
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    # Fórmula de Haversine para calcular la distancia
    a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    # Distancia en kilómetros
    distance = R * c
    return distance

def esta_en_trayectoria_coordenadas(A, B, C):
    lat_A, lon_A = A
    lat_B, lon_B = B
    lat_C, lon_C = C
    # Distancias entre los puntos
    distancia_AC = haversine(lat_A, lon_A, lat_C, lon_C)
    distancia_BC = haversine(lat_B, lon_B, lat_C, lon_C)
    distancia_AB = haversine(lat_A, lon_A, lat_B, lon_B)
    # Verificar si el punto C está en la trayectoria entre A y B
    if distancia_AC + distancia_BC - distancia_AB < 10.05: #Modificar si se nota incorrecta la ruta
        return True
    return False

#---------Identificar si un punto C se encuentra en la ruta AB-------------
def definir_si_esta_en_trayectoria (fincaA,fincaB,fincaC,df):
    locA = locaciones[locaciones["NombreLocacion"] == fincaA]
    locA = locA.reset_index(drop=True)
    A = (locA['Latitud'][0],locA['Longitud'][0] )
    locB = locaciones[locaciones["NombreLocacion"] == fincaB]
    locB = locB.reset_index(drop=True)
    B = (locB['Latitud'][0],locB['Longitud'][0] )
    locC = locaciones[locaciones["NombreLocacion"] == fincaC]
    locC = locC.reset_index(drop=True)
    C = (locC['Latitud'][0],locC['Longitud'][0])

    if esta_en_trayectoria_coordenadas(A, B, C):
        distanciaAB = getKmCc([locA['NombreLocacion'][0],locB['NombreLocacion'][0]],df,"OrigenTray","DestinoTray","Distancia",2)
        distanciaAC = getKmCc([locA['NombreLocacion'][0],locC['NombreLocacion'][0]],df,"OrigenTray","DestinoTray","Distancia",2)
        if distanciaAC<=distanciaAB:
            return True
        else:
            return False
    else:
        return False

aviso = print('Para generar el archivo de manera correcta se debe llenar la hoja de "Parámetros" del Excel en línea llamado "Diccionarios".',end='\n'
              'Presione Enter si está seguro que estos datos ya están correctos en el sitio de Sharepoint.')
aviso2 = input("")

aviso = print('Ingrese: ',end='\n'
              '1: Si se va a correr el programa normal')
print(end='\n')
aviso3 = print('2: Si se van a realizar adicionales')      
adicionales = int(input("Escriba alguna de las 2 opciones anteriores: "))

parametros = get_excel_sh(siteAprovisionamiento,'Indicadores','Agroquímicos','Diccionarios.xlsx','Parámetros',2)
añoConsumo = parametros['Valor'][0]
semanaDescarga = parametros['Valor'][1]
#if semanaDescarga == 51: semanaConsumo =1
#else: semanaConsumo = semanaDescarga+2
tipologia = parametros['Valor'][2]
porcentajeUsoCamion = parametros['Valor'][3]
parametro = parametros['Valor'][4]
kmDiarios = parametros['Valor'][5]
print(semanaDescarga)

#----------------------------------------------------Arreglo de data--------------------------------------------#
#---------------Trayectos
fieldsTrayectos = ['ID','OrigenTray','DestinoTray','Distancia']
ajusteTrayectos = ['OrigenTray','DestinoTray']
trayectos = getList(siteMatEmpaque,'Trayecto',fieldsTrayectos,ajusteTrayectos)
trayectos = ajustarNombresFincas(trayectos,"OrigenTray")
trayectos = ajustarNombresFincas(trayectos,"DestinoTray")

#---------------Tarifa
fieldsTipologias = ['ID','TarifaKilometro','TarifaNodo','Capacidad']
tarifa = getListTarifa(siteMatEmpaque)
tarifa = tarifa[['ID','Tipologia','TarifaKilometro','TarifaNodo','Capacidad']]
tarifa = tarifa[tarifa["Tipologia"] == str(tipologia)]
tarifa = tarifa.reset_index(drop=True)
trayectos['Costo transporte'] = trayectos['Distancia'].fillna(0) * tarifa['TarifaKilometro'][0] + tarifa['TarifaNodo'][0]

#--------------Locaciones
fieldsLocaciones = ['ID','NombreLocacion','Latitud','Longitud']
locaciones = getList(siteMatEmpaque,'LocacionV3',fieldsLocaciones)
locaciones = ajustarNombresFincas(locaciones,"NombreLocacion")
#----------------Demanda
if adicionales==1:
    demanda = get_excel_sh(siteDBLogistics,'Compras',f'{añoConsumo}',f'OfertaDemandaSemana{semanaDescarga}.xlsx','Hoja1',1)
else:
    demanda = get_excel_sh(siteDBLogistics,'Compras',f'{añoConsumo}/Adicionales',f'OfertaDemandaSemana{semanaDescarga}.xlsx','Hoja1',1)
demanda = pd.merge(demanda, trayectos, how='left' ,left_on= ['Finca Disponible','Finca Necesidad'], right_on= ['OrigenTray','DestinoTray'])
demanda['Origen-Destino'] = demanda['Finca Disponible']+demanda['Finca Necesidad']
demanda['Inventario Disponible (peso)'] = np.where(demanda['Uni'] == "GR",demanda[f'Inventario Disponible'],demanda[f'Inventario Disponible']*demanda['Dens'])
demanda['Inventario Faltante (peso)'] = np.where(demanda['Uni'] == "GR",demanda[f'Inventario Faltante'],demanda[f'Inventario Faltante']*demanda['Dens'])

#-----------------------------------------Fase 2.1 (Generar los trayectos individuales de mayor ganancia)-------------------#
origenes =  list(pd.unique(demanda["Finca Disponible"]))
trayectosProductos = pd.DataFrame()
gananciaMax = 0
ponderadoGanancia = 0.7
ponderadoAbastecimiento = 0.3

#try:
for i in range(len(demanda)):
    demanda[f'Inventario de Traslado (peso)'] = np.where(demanda[f'Inventario Disponible (peso)'] >= abs(demanda[f'Inventario Faltante (peso)']),abs(demanda[f'Inventario Faltante (peso)']),demanda[f'Inventario Disponible (peso)'])
    demanda[f'Inventario de Traslado'] = np.where(demanda[f'Inventario Disponible'] >= abs(demanda[f'Inventario Faltante']),abs(demanda[f'Inventario Faltante']),demanda[f'Inventario Disponible'])
    demanda['Ahorro Traslado'] = (demanda['Costo promedio unitario'] * demanda[f'Inventario de Traslado'])

    necesidad = demanda.groupby(['Bodega Necesidad','SisFinCode'],as_index=False).agg({f'Inventario Disponible':'sum',f'Inventario Faltante':'mean', 'Costo promedio unitario': 'mean'})
    necesidad[f'Inventario Faltante Absoluto'] = np.where(necesidad[f'Inventario Faltante'] <= 0, necesidad[f'Inventario Faltante']*(-1) ,necesidad[f'Inventario Faltante'])
    necesidad['Inventario de Traslado'] = necesidad[[f'Inventario Disponible',f'Inventario Faltante Absoluto']].min(axis=1)
    necesidad['Necesidad Total Finca'] = (necesidad['Costo promedio unitario'] * abs(necesidad['Inventario de Traslado']))
    necesidad = necesidad.groupby(['Bodega Necesidad'],as_index=False)['Necesidad Total Finca'].sum()

    demandaGanancia = demanda.groupby(['Bodega Disponible','Bodega Necesidad'],as_index=False).agg({f'Inventario de Traslado (peso)': 'sum', 'Ahorro Traslado': 'sum','Costo transporte':'mean'})
    demandaGanancia['Ganancia'] = demandaGanancia['Ahorro Traslado'] - demandaGanancia[f'Costo transporte']
    demandaGanancia = pd.merge(demandaGanancia, necesidad, how='left' ,left_on= ['Bodega Necesidad'], right_on= ['Bodega Necesidad'])
    demandaGanancia['Porcentaje suplido'] = demandaGanancia['Ahorro Traslado'] / demandaGanancia['Necesidad Total Finca']
    demandaGanancia['Ganancia ponderada'] = (demandaGanancia['Ganancia'].fillna(0)*ponderadoGanancia) + (demandaGanancia['Ganancia'].fillna(0)*abs(demandaGanancia['Porcentaje suplido'].fillna(0)))*ponderadoAbastecimiento
    demandaGanancia = demandaGanancia.sort_values(by = ['Ganancia ponderada','Ganancia','Costo transporte'], ascending = [False,False,True],ignore_index=True )

    if demandaGanancia['Ganancia ponderada'][0]<= parametro*(demandaGanancia['Costo transporte'][0]):
        break

    bodegaOrigen = demandaGanancia['Bodega Disponible'][0]
    bodegaDestino = demandaGanancia['Bodega Necesidad'][0]

    demandaFiltrado = demanda[demanda["Bodega Disponible"] == bodegaOrigen]
    demandaFiltrado = demandaFiltrado[demandaFiltrado["Bodega Necesidad"] == bodegaDestino]
    trayectosProductos = pd.concat([trayectosProductos, demandaFiltrado], ignore_index=True)

    demanda = demanda.loc[(demanda['Origen-Destino'] != f'{bodegaOrigen}{bodegaDestino}')]
    demanda = pd.merge(demanda, demandaFiltrado[['Bodega Disponible','Finca Disponible','SisFinCode','Finca Necesidad','Inventario de Traslado (peso)','Inventario de Traslado']], how='left' ,left_on= ['Bodega Disponible','Finca Disponible','SisFinCode'], right_on= ['Bodega Disponible','Finca Disponible','SisFinCode'])
    demanda.rename(columns = {'Finca Necesidad_x':'Finca Necesidad','Inventario de Traslado (peso)_x':'Inventario de Traslado (peso)','Inventario de Traslado_x':'Inventario de Traslado'}, inplace = True)
    demanda['Inventario Disponible (peso)'] = np.where(demanda['Bodega Disponible'] == bodegaOrigen,demanda['Inventario Disponible (peso)']-demanda['Inventario de Traslado (peso)_y'].fillna(0), demanda['Inventario Disponible (peso)']) 
    demanda[f'Inventario Disponible'] = np.where(demanda['Bodega Disponible'] == bodegaOrigen,demanda[f'Inventario Disponible']-demanda['Inventario de Traslado_y'].fillna(0), demanda[f'Inventario Disponible'])
    demanda.drop(['Finca Necesidad_y','Inventario de Traslado (peso)_y','Inventario de Traslado_y'], inplace=True, axis=1)

    demanda = pd.merge(demanda, demandaFiltrado[['Finca Disponible','SisFinCode','Bodega Necesidad','Finca Necesidad','Inventario de Traslado (peso)','Inventario de Traslado']], how='left' ,left_on= ['Bodega Necesidad','Finca Necesidad','SisFinCode'], right_on= ['Bodega Necesidad','Finca Necesidad','SisFinCode'])
    demanda.rename(columns = {'Finca Disponible_x':'Finca Disponible','Inventario de Traslado (peso)_x':'Inventario de Traslado (peso)','Inventario de Traslado_x':'Inventario de Traslado'}, inplace = True)
    demanda['Inventario Faltante (peso)'] = np.where(demanda['Bodega Necesidad'] == bodegaDestino,demanda['Inventario Faltante (peso)']+demanda['Inventario de Traslado (peso)_y'].fillna(0), demanda['Inventario Faltante (peso)'])
    demanda[f'Inventario Faltante'] = np.where(demanda['Bodega Necesidad'] == bodegaDestino,demanda[f'Inventario Faltante']+demanda['Inventario de Traslado_y'].fillna(0), demanda[f'Inventario Faltante'])
    demanda.drop(['Finca Disponible_y','Inventario de Traslado (peso)_y','Inventario de Traslado_y'], inplace=True, axis=1)

demanda = demanda[['Bodega Disponible','Finca Disponible','SisFinCode','Quimico','Uni','Dens',f'Inventario Disponible (peso)','Bodega Necesidad','Finca Necesidad',f'Inventario Faltante (peso)','Fecha último movimiento',f'Inventario de Traslado (peso)','Costo promedio unitario','Ahorro Traslado','Distancia','Costo transporte']]
trayectosProductos.rename(columns = {'Inventario Disponible (peso)':f'Inventario Disponible (peso)','Inventario Faltante (peso)':f'Inventario Faltante (peso)'}, inplace = True)
print(trayectosProductos)
create_excel(trayectosProductos,"Trayectos","Hoja1")
trayectosProductos = trayectosProductos[['Bodega Disponible','Finca Disponible','SisFinCode','Quimico','Uni','Dens',f'Inventario Disponible (peso)','Bodega Necesidad','Finca Necesidad',f'Inventario Faltante (peso)','Fecha último movimiento',f'Inventario de Traslado (peso)','Inventario de Traslado','Costo promedio unitario']]
trayectosProductos['Costo carga'] = trayectosProductos['Inventario de Traslado'] * trayectosProductos['Costo promedio unitario']
trayectosProductos = trayectosProductos[trayectosProductos["Costo carga"] >= 30000]
trayectosProductos['Decisión'] = "No"
fincas = get_excel_sh(siteAprovisionamiento,'Indicadores','Agroquímicos','Diccionarios.xlsx','Fincas',2)
trayectosProductos = pd.merge(trayectosProductos,fincas[['Bodega','Descripción Bodega']],how='left' ,left_on= ['Bodega Disponible'], right_on= ['Bodega'])
fincas.rename(columns = {'Bodega':'Bodega 2','Descripción Bodega':'Descripción Bodega 2'}, inplace = True)
trayectosProductos = pd.merge(trayectosProductos,fincas[['Bodega 2','Descripción Bodega 2']],how='left' ,left_on= ['Bodega Necesidad'], right_on= ['Bodega 2'])
trayectosProductos.rename(columns = {'Descripción Bodega':'Descripción Bodega Disponible','Descripción Bodega 2':'Descripción Bodega Necesidad '}, inplace = True)
trayectosProductos.drop(['Bodega',"Bodega 2"], inplace=True, axis=1)
create_excel(trayectosProductos,"Productos a trasladar","Hoja1")

#--------------------------------Fase 2.2 Código del TSP (Salesman vendor model)--------------------------------# 
trayectosNodosTotales = trayectosProductos.groupby(['Finca Disponible','Finca Necesidad'],as_index=False)[f'Inventario Disponible (peso)'].sum()
fincasOrigen =  list(pd.unique(trayectosNodosTotales["Finca Disponible"]))
trayectosParciales = []
inventarioTraslado = trayectosProductos.groupby(['Finca Disponible','Finca Necesidad'],as_index=False)[f'Inventario de Traslado'].sum()

for i in fincasOrigen:
    print(f'Origen: {i}')
    trayectosNodos = trayectosNodosTotales[trayectosNodosTotales["Finca Disponible"] == i]
    trayectosNodos = trayectosNodos.reset_index(drop=True)
    nodos =  list(pd.unique(trayectosNodos["Finca Necesidad"]))
    nodos.insert(0,i) #Generar lista de nodos a recorrer
    nodos.sort()

    #Generar la matriz de distancias con los nodos a recorrer
    fincasTrayectos = trayectos[trayectos.OrigenTray.isin(nodos)]
    fincasTrayectos = fincasTrayectos[fincasTrayectos.DestinoTray.isin(nodos)]
    fincasTrayectos = pd.pivot_table(fincasTrayectos,values='Distancia',columns=['DestinoTray'],index=['OrigenTray'], aggfunc="mean")
    fincasTrayectos['index_column'] = fincasTrayectos.index
    matrizDistancias = pd.DataFrame(fincasTrayectos['index_column'].tolist()).fillna('').add_prefix('m')
    fincasTrayectos.drop(['index_column'], inplace=True, axis=1)
    fincasTrayectos['Key'] = range(1, len(fincasTrayectos) + 1)
    matrizDistancias['Key'] = range(1, len(matrizDistancias) + 1)
    matrizDistancias = pd.merge(matrizDistancias,fincasTrayectos,how='left' ,left_on= ['Key'], right_on= ['Key'])
    matrizDistancias.drop(['Key'], inplace=True, axis=1)
    matrizDistancias.rename(columns = {'m0':'Origen'}, inplace = True)
    matrizDistancias = matrizDistancias.reset_index(drop=True)

    #Parámetros para realizar el modelo de Branch and Bound
    fincaOrigen = trayectosNodos['Finca Disponible'][0]
    origen = matrizDistancias.index[matrizDistancias["Origen"] == i].tolist()
    diccionarioFincas = matrizDistancias.loc[:, 'Origen']
    diccionarioFincas = diccionarioFincas.to_dict()
    
    # Convertir dataframe a matriz
    matrizDistancias.to_numpy()
    matrix = matrizDistancias[nodos].to_numpy()
    #if capacidadCarro>=totalPesoTrasladoFincaOrigen:
    start_node = origen[0]  # Origen definido
    best_path, best_cost = tsp_branch_and_bound_no_return(matrix, start_node)
    best_path = [diccionarioFincas[indice] for indice in best_path]
    trayectosParciales.append(best_path)
    print("Mejor ruta: ", best_path)
    print("Km recorridos: ", best_cost)
    print("Costo total recorrido: ", best_cost* tarifa['TarifaKilometro'][0])

if adicionales==1:
    file_upload_to_sharepoint(siteAprovisionamiento,añoConsumo,f'Semana{int(semanaDescarga)}',"Productos a trasladar")
else:
    file_upload_to_sharepoint(siteAprovisionamiento,añoConsumo,f'Semana{int(semanaDescarga)}/Adicionales',"Productos a trasladar")
print("Cargado el archivo de productos a trasladar en el Sharepoint")


#--------------------Fase 2.3 (Ajustar trayectos según capacidad de camión y kms máximos diarios)
#-------------------- 2.3.1. Ajustar trayectos según capacidad del camión
trayectosFinales = []
capacidadCarro = int(tarifa['Capacidad'][0])*(porcentajeUsoCamion/100)*(1e6)

for recorridoCapacidad in trayectosParciales:
    ccTotales = getKmCc(recorridoCapacidad,inventarioTraslado,'Finca Disponible','Finca Necesidad','Inventario de Traslado',1)
    #Dividir trayectos según la cantidad de cc
    trayectosDivididos,ccKmRecorridos = restarKmCC(recorridoCapacidad,inventarioTraslado,"Finca Disponible","Finca Necesidad","Inventario de Traslado",1,ccTotales,capacidadCarro)                           
    for i in trayectosDivididos: trayectosFinales.append(i)

#--------------------2.3.2 Ajustar trayectos según cantidad de km recorridos
listaDeListas = []
for recorrido in trayectosFinales:
    kmTotales = getKmCc(recorrido,trayectos,'OrigenTray','DestinoTray','Distancia',2)
    trayectosDivididosKm,ccKmRecorridos = restarKmCC(recorrido,trayectos,"OrigenTray","DestinoTray","Distancia",2,kmTotales,kmDiarios)                        
    #Dividir trayectos según la cantidad de km
    for i in trayectosDivididosKm: listaDeListas.append(i)

#-------------------Fase 2.4 (Mochilero style)---------------------
#----Obtener para cada trayecto su km y asignar una posición en un diccionario
dictTrayectorias = {}
listaDeListas = sorted(listaDeListas, key=len, reverse=True)

for i in listaDeListas:
    indiceLista =  listaDeListas.index(i)
    km = getKmCc(i,trayectos,'OrigenTray','DestinoTray','Distancia',2)
    dictTrayectorias[indiceLista] = [i,km]

#-----Ejercicio del mochilero
dictAsAList = list(dictTrayectorias)
listTrayectosFinales = [] #Final, final, no va más
saltarUnElemento = 1000    #Porque ya se introdujo en otra lista, se coloca cualquier valor para inicializar la variable

for key in dictTrayectorias:
    if key == saltarUnElemento:
        continue
    kmKey = dictTrayectorias[key][1]
    listTrayectosKey = dictTrayectorias[key][0]
    if key>0: 
        listTrayectosKeyAnterior = dictTrayectorias[key-1][0]
    else:
        listTrayectosKeyAnterior = dictTrayectorias[key][0] 
    listaIntroducirElementos = listTrayectosKey.copy()
    cargarKey = True
    for i in range((key+1),len(dictAsAList)):             
        listTrayectosKeyI = dictTrayectorias[i][0]
        kmKeyI = dictTrayectorias[i][1]
        listaBooleanos = []
        if key == 0:
            #ccCapacidadInstanteKey = getKmCc(listTrayectosKey,inventarioTrasladoBorrar,"Finca Disponible","Finca Necesidad","Inventario de Traslado",1)#Borrar después          
            ccCapacidadInstanteKey = getKmCc(listTrayectosKey,inventarioTraslado,"Finca Disponible","Finca Necesidad","Inventario de Traslado",1)             
        else:
            #ccCapacidadInstanteKey = getKmCc(listTrayectosKey,inventarioTrasladoBorrar,"Finca Disponible","Finca Necesidad","Inventario de Traslado",3,listTrayectosKeyAnterior)
            ccCapacidadInstanteKey = getKmCc(listTrayectosKey,inventarioTraslado,"Finca Disponible","Finca Necesidad","Inventario de Traslado",3,listTrayectosKeyAnterior)

        #ccCapacidadInstanteKeyI = getKmCc(listTrayectosKeyI,inventarioTrasladoBorrar,"Finca Disponible","Finca Necesidad","Inventario de Traslado",3,listTrayectosKeyAnterior)
        ccCapacidadInstanteKeyI = getKmCc(listTrayectosKeyI,inventarioTraslado,"Finca Disponible","Finca Necesidad","Inventario de Traslado",3,listTrayectosKeyAnterior)

        if kmKey + kmKeyI<200:
            posicionIngreso = 0
            for k in range (0,len(listTrayectosKeyI)):
                cargarElKActual = False
                soloCargarUnJenUnaK = False
                for j in range (posicionIngreso,len(listTrayectosKey)-1):  
                    fincaA,fincaB,fincaC = listTrayectosKey[j],listTrayectosKey[j+1],listTrayectosKeyI[k]
                    #Tienen que estar todos los de listTrayectosKeyI en alguna trayectoria entre dos fincas de listTrayectosKey
                    if definir_si_esta_en_trayectoria(fincaA,fincaB,fincaC,trayectos):
                        contador = 1
                        # Las demás fincas deben estar después, no se puede recorrer toda la lista para meter los demás elementos
                        for finca in listTrayectosKey:
                            if finca == fincaA:
                                break
                            try:
                                #origen = inventarioTrasladoBorrar[inventarioTrasladoBorrar["Finca Disponible"] == listTrayectosKey[0]] #Cambiar después el nombre del dataframe
                                origen = inventarioTraslado[inventarioTraslado["Finca Disponible"] == listTrayectosKey[0]] #Cambiar después el nombre del dataframe
                                destino = filtrarDataframe(origen,listTrayectosKey,contador,"Finca Necesidad")
                                ccCapacidadInstanteKey = ccCapacidadInstanteKey - destino["Inventario de Traslado"][0]
                                contador+=1
                            except:
                                listTrayectosKeyCopia = listTrayectosKey.copy()
                                listTrayectosKeyAnteriorCopia = listTrayectosKeyAnterior.copy()
                                for i in range(1,len(listTrayectosKeyCopia)): listTrayectosKeyAnteriorCopia.append(listTrayectosKeyCopia[i])
                                #origen = inventarioTrasladoBorrar[inventarioTrasladoBorrar["Finca Disponible"] == listTrayectosKeyAnteriorCopia[0]] #Cambiar después el nombre del dataframe
                                origen = inventarioTraslado[inventarioTraslado["Finca Disponible"] == listTrayectosKeyAnteriorCopia[0]] #Cambiar después el nombre del dataframe
                                destino = filtrarDataframe(origen,listTrayectosKeyAnteriorCopia,contador,"Finca Necesidad")
                                ccCapacidadInstanteKey = ccCapacidadInstanteKey - destino["Inventario de Traslado"][0]
                                contador+=1
                        if ccCapacidadInstanteKey+ccCapacidadInstanteKeyI<=capacidadCarro:
                            indiceFincaA =  listaIntroducirElementos.index(fincaA)
                            listaIntroducirElementos.insert(indiceFincaA + 1, fincaC)
                            listaBooleanos.append(True)
                            soloCargarUnJenUnaK = True
                            cargarElKActual = True
                            posicionIngreso = indiceFincaA+1
                            break
                    if soloCargarUnJenUnaK == True: break
                if cargarElKActual == False: break
        if len(listaBooleanos) == len(listTrayectosKeyI): 
            listTrayectosFinales.append(listaIntroducirElementos)
            cargarKey = False
            saltarUnElemento = i
            break
    if cargarKey == True: listTrayectosFinales.append(listTrayectosKey)

print(f'Lista de trayectos finales: {listTrayectosFinales}')

'''except:
    column_names = ['Bodega Disponible', 'Finca Disponible', 'SisFinCode', 'Quimico', 'Uni','Dens', 'Inventario Disponible (peso)', 'Bodega Necesidad', 'Finca Necesidad', 'Inventario Faltante (peso)', 'Fecha último movimiento', 'Inventario de Traslado (peso)', 'Inventario de Traslado', 'Costo promedio unitario', 'Decisión', 'Descripción Bodega Disponible', 'Descripción Bodega Necesidad']
    # Crear un DataFrame vacío con los nombres de las columnas
    df = pd.DataFrame(columns=column_names)
    create_excel(df,"Productos a trasladar","Hoja1") 
    print("No se presentan traslados sugeridos de productos")'''

#Time
producto = f'Fase 2: Generación de productos a trasladar entre bodegas'
elapsed_time = time() - start_time
print("Tiempo: %.10f segundos." % elapsed_time)
input(f'{producto}!... Presione enter para salir. ')