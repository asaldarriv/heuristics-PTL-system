import pandas as pd

instance_name = 'Data_40_Salidas_composición_zonas_homogéneas.xlsx'

excel_modelo = pd.ExcelFile(f'instances-ptl/{instance_name}')

Conjunto_pedidos = pd.read_excel(excel_modelo, 'Pedidos', index_col=0) #Pedidos
Conjunto_zonas = pd.read_excel(excel_modelo, 'Zonas', index_col=0) #Zonas
Conjunto_salidas = pd.read_excel(excel_modelo, 'Salidas', index_col=0) #Salidas
Conjunto_skus = pd.read_excel(excel_modelo, 'SKU', index_col=0) #SKUs
#Conjunto_trabajadores = pd.read_excel(excel_modelo, 'Trabajadores', index_col=0) #Trabajadores

#Lectura de parametros de salidas
N_Salidas = pd.read_excel(excel_modelo, 'Salidas_en_cada_zona', index_col=0) #Cantidad de salidas por cada zona
Salidas_por_zona = pd.read_excel(excel_modelo, 'Salidas_pertenece_zona', index_col=0) #Parametro binario, salidas que están incluidas en cada zona
Tiempo_salidas = pd.read_excel(excel_modelo, 'Tiempo_salida', index_col=0) #Tiempo para desplazarse desde el lector al punto medio de cada salida

#Lectura de parametros de SKUs
SKUS_por_pedido = pd.read_excel(excel_modelo, 'SKU_pertenece_pedido', index_col=0) #Parámetro binario, SKUS que están incluidas en un pedido
Tiempo_SKU = pd.read_excel(excel_modelo, 'Tiempo_SKU', index_col=0) #Tiempo total de lectura, conteo, separación, depósito de cada ref por pedidotr

#Lectura de parametros adicionales
Parametros = pd.read_excel(excel_modelo, 'Parametros', index_col=0) #Parametros

# Conjuntos
pedidos = list(Conjunto_pedidos.index)
zonas = list(Conjunto_zonas.index)
salidas = list(Conjunto_salidas.index)
skus = list(Conjunto_skus.index)

# Parametros
V = Parametros['v']
ZN = Parametros ['zn']

v=float(V.iloc[0])
zn=float(ZN.iloc[0])

#Pertenencia de las salidas a las zonas
#Salidas que están incluidas en cada zona
s={(j,k):Salidas_por_zona.at[j,k] for k in salidas for j in zonas}

#Pertenencia de una SKU a un pedido
#SKUS que están incluidas en un pedido
rp={(i,m):SKUS_por_pedido.at[i,m] for m in skus for i in pedidos}

#Distancia a cada una de las salidas
#Tiempo de surtir un SKU en una salida
d={(j,k):Tiempo_salidas.at[j,k] for k in salidas for j in zonas}

#trim : Tiempo de surtir todas las unidades de una SKU dada ( m∈R ) en un pedido particular ( i∈P ).
tra={(i,m):Tiempo_SKU.at[i,m] for m in skus for i in pedidos}

#Numero de salidas en cada zona
n_sal = N_Salidas.Num_Salidas