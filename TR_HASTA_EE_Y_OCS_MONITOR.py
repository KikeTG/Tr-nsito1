# %%
#En esta sección del codigo se importan todas las librerias que 

import pandas as pd
import numpy as np
import os
import datetime
import win32com.client
import time
import getpass

usuario = getpass.getuser()

inicio = time.time()

# %%
#Lectura de los archivos de la ME2L en el repositorio de Sharepoint
me2l1_df = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/AFM/ME2L1.XLSX", dtype={'Documento compras':'str','Posición':'str', 'Material':'str'})
me2l2_df = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/AFM/ME2L2.XLSX", dtype={'Documento compras':'str','Posición':'str', 'Material':'str'})
me2l3_df = pd.read_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/AFM/ME2L3.XLSX", dtype={'Documento compras':'str','Posición':'str', 'Material':'str'})

#eliminacion de filas en blanco que vienen por defecto en SAP
me2l1_df = me2l1_df.drop(0)
me2l2_df = me2l2_df.drop(0)
me2l3_df = me2l3_df.drop(0)


#consolidacion de todos los dataframes de me2l en uno solo
me2l_consolidado = pd.concat([me2l1_df,me2l2_df,me2l3_df])
me2l_consolidado.reset_index(drop=True, inplace=True)

# %%
#localizacion y lectura de base planificable 

base = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros/Base Planificable"
base_carpeta = os.listdir(base)
for i in base_carpeta:
    if str(datetime.datetime.today().year) + '-' + str(datetime.datetime.today().month).zfill(2) in i:
        ruta = os.path.join(base, i)
        print(ruta)

        ruta_base = os.listdir(ruta)
        for i in ruta_base:
            if 'Base' in i and not 'R3' in i:
                ruta_final = os.path.join(ruta, i)
                print(ruta_final)
                base = pd.read_excel(ruta_final,header=1)
                base_df = base[['Material','Texto breve de material','NomSector_actual','TIPO','Corresponde']]
                base_df_prov = base[['Proveedor', 'Pais (Proveedor)']]
        
base_df_prov = base_df_prov.drop_duplicates(subset=['Proveedor'])

# %%
#localizacion y lectura de archivo de cadena de reemplazo

ruta = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras Maestros"
lista_ruta = os.listdir(ruta)
for i in lista_ruta:
    if str(datetime.datetime.today().year) in i:
        ruta = os.path.join(ruta , i)
        lista_ruta = os.listdir(ruta)
        for i in lista_ruta:
            
            if str(datetime.datetime.today().month).zfill(2) in i:
                ruta = os.path.join(ruta, i)
                lista_ruta = os.listdir(ruta)
                for i in lista_ruta:
                    if 'COD_ACTUAL_S4' in i and not 'R3' in i:
                        ruta = os.path.join(ruta, i)
                        print(ruta)
                        cod_actual_df = pd.read_excel(ruta, usecols = ['Nro_pieza_fabricante_1', 'Cod_Actual_1'])

# %% [markdown]
# TRATAMIENTO DE DFS (OC, POSiCION)

# %%
#medida de control para cantidades de me2l

print(f'Dimensiones del df1: {me2l1_df.shape}')
print('-' * 40)
print(f'Dimensiones del df2: {me2l2_df.shape}')
print('-' * 40)
print(f'Dimensiones del df3: {me2l3_df.shape}')
print(f'Dimensiones de consolidado: {me2l_consolidado.shape}')

# %%
print("Cantidades de lineas por clase de documento:" + "\t" + str(me2l_consolidado['Cl.documento compras'].value_counts()))

# %% [markdown]
# UNIR DFS

# %%
#Cruce de base planificable con codigo actual cadena de reemplazo
base_df_ue = pd.merge(base_df, cod_actual_df, left_on="Material", right_on="Nro_pieza_fabricante_1", how="left")
base_df_ue['Cod_Actual_1'] = base_df_ue['Cod_Actual_1'].fillna(base_df_ue['Material'])

# %% [markdown]
# ULTIMO ESLABON A ME2L

# %%
#cruce de me2l con codigo actual cadena de reemplazo
me2l_ue = pd.merge(me2l_consolidado, cod_actual_df, left_on="Material", right_on="Nro_pieza_fabricante_1", how="left")
me2l_ue['Cod_Actual_1'].fillna(me2l_ue['Material'], inplace=True)

# %%
#reduccion y reordenamiento de columnas de me2l
me2l_ue = me2l_ue[['Documento compras', 'Posición', 'Reparto','Cl.documento compras','Tipo doc.compras',
       'Grupo de compras','Historial pedido/Docu.orden entrega', 'Fecha documento', 'Material','Cod_Actual_1', 'Texto breve','Grupo de artículos',
       'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
       'Unidad medida pedido', 'Precio neto', 'Moneda',
       'Fecha de entrega','Hora', 'Fecha entrega estad.', 'Cantidad anterior',
       'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
       'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación',
       'Cantidad de posiciones', 'Nombre de proveedor',
       'Por entregar (cantidad)', 'Por entregar (valor)',
       'Por calcular (cantidad)', 'Por calcular (valor)'
       ]]

me2l_ue = me2l_ue.rename(columns={'Material':'Material Antiguo'})
me2l_ue = me2l_ue.rename(columns={'Cod_Actual_1':'Material'})


# %%
#Cruce de me2l con base planificable y posterior reduccion de columnas
me2l_cruce_sector = pd.merge(me2l_ue, base_df_ue[['Material','Texto breve de material', 'NomSector_actual', 'TIPO','Corresponde']], left_on="Material", right_on="Material", how="left")
#me2l_ue.to_excel("C:/Users/lravlic/PROYECTOS DATA/PRUEBAS TRANSITO/me2l_sector.xlsx")
me2l_cruce_sector.shape
me2l_cruce_sector = me2l_cruce_sector[['Documento compras', 'Posición', 'Reparto','Cl.documento compras','Tipo doc.compras',
       'Grupo de compras','Historial pedido/Docu.orden entrega', 'Fecha documento', 'Material Antiguo','Material', 'Texto breve','Grupo de artículos',
       'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
       'Unidad medida pedido', 'Precio neto', 'Moneda', 
       'Fecha de entrega','Hora', 'Fecha entrega estad.', 'Cantidad anterior',
       'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
       'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación',
       'Cantidad de posiciones', 'Nombre de proveedor',
       'Por entregar (cantidad)', 'Por entregar (valor)',
       'Por calcular (cantidad)', 'Por calcular (valor)', 'NomSector_actual',
       'Texto breve de material', 'TIPO', 'Corresponde']]
me2l_cruce_sector['Cod_Prov'] = me2l_cruce_sector['Nombre de proveedor'].str.split(' ', expand=True)[0]
me2l_cruce_sector = me2l_cruce_sector.merge(base_df_prov,left_on='Cod_Prov', right_on='Proveedor', how='left')
me2l_cruce_sector['Posición'] = me2l_cruce_sector['Posición'].astype('str')
me2l_cruce_sector['AUX'] = me2l_cruce_sector['Documento compras'] + me2l_cruce_sector['Posición']
me2l_cruce_sector['Origen'] = me2l_cruce_sector['Pais (Proveedor)']
me2l_cruce_sector = me2l_cruce_sector[['AUX','Documento compras', 'Posición', 'Reparto','Cl.documento compras','Tipo doc.compras',
       'Grupo de compras','Historial pedido/Docu.orden entrega', 'Fecha documento', 'Material Antiguo','Material', 'Texto breve','Grupo de artículos',
       'Indicador de borrado','Tipo de posición', 'Tipo de imputación', 'Centro', 'Almacén', 'Cantidad de reparto', 'Cantidad de pedido',
       'Unidad medida pedido', 'Precio neto', 'Moneda',
       'Fecha de entrega','Hora', 'Fecha entrega estad.', 'Cantidad anterior',
       'Cantidad entrada', 'Cantidad de salida', 'Cantidad entregada',
       'Solicitud de pedido', 'Pos.solicitud pedido', 'Indicador creación',
       'Cantidad de posiciones', 'Nombre de proveedor',
       'Por entregar (cantidad)', 'Por entregar (valor)',
       'Por calcular (cantidad)', 'Por calcular (valor)', 'NomSector_actual','Origen', 'TIPO',  'Corresponde']]
me2l_base_tr = me2l_cruce_sector[me2l_cruce_sector['NomSector_actual'].notna()]
me2l_base_tr['Corresponde'].fillna(0, inplace=True)
me2l_base_tr['Corresponde'].replace({0:1}, inplace=True)


# %%
#agrupacion de ordenes de compra por documento y posicion para importacion en sap 
bases_oc = me2l_base_tr.groupby(['Documento compras'])['Posición'].count().sort_values(ascending=False).reset_index()


# %%
# Calculate the length of the string
str_len = len(bases_oc['Documento compras'])

# Calculate the start and end indices for each part
part_len = str_len // 5
part_indices = [(i * part_len, (i + 1) * part_len) for i in range(5)]
part_indices[-1] = (part_indices[-1][0], str_len)  # Adjust the end index of the last part

# Divide the string into 5 parts using list slicing
parts = [bases_oc['Documento compras'][start:end] for start, end in part_indices]


# %%
# Define the number of parts you want to divide the string into
num_parts = 5

# Calculate the length of the string
str_len = len(bases_oc['Documento compras'])

# Calculate the start and end indices for each part
part_len = str_len // num_parts
part_indices = [(i * part_len, (i + 1) * part_len) for i in range(num_parts)]
part_indices[-1] = (part_indices[-1][0], str_len)  # Adjust the end index of the last part

# Divide the string into 'num_parts' parts using list slicing and maintain index starting at 0
parts = [bases_oc['Documento compras'][start:end].reset_index(drop=True) for start, end in part_indices]

# Convert the divided parts into a DataFrame with each part as a column
df = pd.concat(parts, axis=1)
df.columns = [f'Part {i+1}' for i in range(num_parts)]
df.to_excel(f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/AFM/OCS_A_DESCARGAR_MONITOR.xlsx", index=False)

# %%
bases_oc['Documento compras'].to_clipboard(header=False, index=False)

# %%
#exportacion de me2l cruzada con base planificable a sharepoint
carpeta_tubo = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Tubo Semanal/"
#{str((datetime.datetime.today()).strftime('%Y-%m-%d'))}/{str((datetime.datetime.today()).strftime('%Y-%m-%d'))} ME2L S4.xlsx"
lista_tubo = os.listdir(carpeta_tubo)
sem_actual = lista_tubo[-2]
destino_archivo = (os.path.join(carpeta_tubo,sem_actual,sem_actual) + ' ME2L S4.xlsx')
print(destino_archivo)




me2l_base_tr.to_excel(destino_archivo)

# %%
#advertencia para setear sap gui para ingresar a sap
import easygui as eg
eg.msgbox('Se iniciara descarga de VL06IF desde SAP, por favor asegura tener iniciada tu sesión de SAP y no tener mas sesiones abiertas', 'Transito S4', 'OK')

# %%
#descarga automatica de vl06if a traves de ordenes de compras copiadas en codigo desde me2l filtrada

SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "VL06IF"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").text = ""
session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").setFocus()
session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").setFocus()
session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").caretPosition = 0
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/btn%_IT_EBELN_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/tbar[0]/btn[24]").press()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/tbar[0]/btn[8]").press()
session.findById("wnd[0]").sendVKey(8)
session.findById("wnd[0]/tbar[1]/btn[18]").press()
session.findById("wnd[0]/tbar[1]/btn[33]").press()
session.findById("wnd[1]/usr/lbl[1,6]").setFocus()
session.findById("wnd[1]/usr/lbl[1,6]").caretPosition = 7

session.findById("wnd[1]").sendVKey(2)
session.findById("wnd[0]").sendVKey(43)
session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = f"C:/Users/{usuario}/Inchcape/Planificación y Compras Chile - Documentos/Planificación y Compras KPI-Reportes/Bases Transito (python)/AFM"
session.findById("wnd[1]/tbar[0]/btn[11]").press()
session.findById("wnd[0]").sendVKey(3)
session.findById("wnd[0]").sendVKey(3)




# %%



