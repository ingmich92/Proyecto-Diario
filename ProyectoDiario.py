from decimal import ROUND_DOWN
from msilib import Table
import sqlite3
from turtle import color
import psycopg2
import pandas as pd
import smtplib
from datetime import date
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import matplotlib.pyplot as plt
import numpy as np
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from tabulate import tabulate
import dataframe_image as dfi
import numpy
import geopandas as gpd

##Conexion a BDD
PSQL_HOST = "creze-aws-aurora-cluster.cluster-cmiufpwtxjkk.us-east-1.rds.amazonaws.com"
PSQL_PORT = "5432"
PSQL_USER = "laguilar"
PSQL_PASS = "0FsAC42&.+j.7zjM"
PSQL_DB   = "creze"

connstr = "host=%s port=%s user=%s password=%s dbname=%s" % (PSQL_HOST, PSQL_PORT, PSQL_USER, PSQL_PASS, PSQL_DB)
conn = psycopg2.connect(connstr)

today = date.today()
dia = (today.day)
meses = ["Enero", "Febrero", "Marzo", "Abri", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes = meses[today.month - 1]

###Tabla 1 Nueva   balance y bucket nuevo ######   
df = pd.read_sql("select distinct piv.Buckets_nif, sum(piv.balance) as balance,  "
                +"sum(piv.reserva_nif) as reserva_nif "
                +"from(select distinct *, case when dpd_final=0 then 'a.0 DPD' "
                +"when dpd_final<=14 then 'b.1-14 DPD' "
                +"when dpd_final<=30 then 'c.15-30 DPD' "
                +"when dpd_final<=60 then 'd.31-60 DPD' "
                +"when dpd_final<=90 then 'e.61-90 DPD' "
                +"when dpd_final<=120 then 'f.91-120 DPD' "
                +"when dpd_final<=150 then 'g.121-150 DPD' "
                +"when dpd_final<=179 then 'h.151-179 DPD' "
                +"else 'i.Castigos' "
                +"end as Buckets_nif "                
                +"from riesgo.seguimiento_diario_mes_reserva_nif  "      
                +f"where dia={dia} and dpd_final<=179) piv  "                                                                                                                           
                +"group by 1",conn)

##Tabla1 balance y bucket nuevo    ANTERIOR     
# df = pd.read_sql("select buckets_cub,sum(balance) "
#                 +"from riesgo.seguimiento_diario_mes "
#                 +f"where dia= {dia} and buckets<>'i.Castigos' "
#                 +"group by 1 "
#                 +"order by 1",conn)

###Tabla 2 MONTO OPERADO######   
df2 = pd.read_sql("select piv2.status2 as status, sum(piv2.monto_neto_total) as monto_operado_total from(select fn.*,case when fn.garantia is not null then  'Secured' else fn.status end as status2, "
                +"case when fn.status not IN ('Anticipo a Capital ','Nuevo ','Refinanciamiento ','Refinanciamiento Plus ','Renov Bullet ','Subsecuente ') then 0 else fn.monto_neto end as monto_neto_total "
                +"from(select distinct piv.*,b.monto_neto,b.status,cs.garantia "
                +"from (select fecha_reporte,fecha_contrato,cast(to_char(fecha_contrato, 'DD')as numeric) as dia_firma,folio,vintage,monto_operado "
                +"from cartera_corte_credito() ccc "
                +"where ccc.folio <> 'null' and ccc.id_status<>10 and /*ccc.fecha_contrato>='2022-10-01' and ccc.fecha_contrato<'2022-11-01'*/ ccc.fecha_contrato>=date_trunc('month', current_date - interval '1' month) and ccc.fecha_contrato<cast(date_trunc('month', current_date) as date )) piv "
                +"left join cobranza.base_originacion b on piv.folio=b.folio left join riesgo.creditos_secured cs on piv.folio=cs.folio) fn)piv2 where piv2.monto_neto_total>0 group by 1",conn)  

##Tabla3 MONTO NETO######  
df3 = pd.read_sql("select piv2.status2 as status, sum(piv2.monto_neto_total) as monto_neto_total from(select fn.*,case when fn.garantia is not null then  'Secured' else fn.status end as status2, "
                +"case when fn.status not IN ('Anticipo a Capital ','Nuevo ','Refinanciamiento ','Refinanciamiento Plus ','Renov Bullet ','Subsecuente ') then 0 else fn.monto_neto end as monto_neto_total "
                +"from(select distinct piv.*,b.monto_neto,b.status,cs.garantia "
                +"from (select fecha_reporte,fecha_contrato,cast(to_char(fecha_contrato, 'DD')as numeric) as dia_firma,folio,vintage,monto_operado "
                +"from cartera_corte_credito() ccc "
                +"where ccc.folio <> 'null' and ccc.id_status<>10 and /*ccc.fecha_contrato>='2022-11-01'*/ ccc.fecha_contrato>=cast(date_trunc('month', current_date) as date )) piv "
                +"left join cobranza.base_originacion b on piv.folio=b.folio left join riesgo.creditos_secured cs on piv.folio=cs.folio) fn)piv2 where piv2.monto_neto_total>0 group by 1",conn)         

##Tabla4 Pronostico anterior ICV
# df4 = pd.read_sql("select distinct fn.dia, (cast(to_char(fn_getlastdayofmonth(date_trunc('month', current_date)::date),'DD') as numeric)) as dia_ultimo, "
#                 +"round(sum(fn.balance)/1000000,2) as balance_hoy, "
#                 +"round(sum(fn.balance_90_mas)/1000000,2) as vencido_hoy, "
#                 +"round((sum (fn.balance_90_mas)/ sum(fn.balance))*100,2) as ICV, "
#                 +"round(sum(fn.pron_balance_90_mas)/1000000,2) as vencido_pronostico, "
#                 +"round((sum (fn.pron_balance_90_mas)/ sum(fn.balance))*100,2) as ICV_pronostico "
#                 +"from (select distinct *, "
#                 +"case when a.buckets in ('f.90-119 DPD','g.120-149 DPD','h.150-179 DPD') then a.balance else 0 end as balance_90_mas "
#                 +",(dias_vencido+(fn_getlastdayofmonth(date_trunc('month', current_date)::date)- current_date)) as pron_dias_vencido, "
#                 +"case when (dias_vencido+(fn_getlastdayofmonth(date_trunc('month', current_date)::date)- current_date))>=90 then a.balance else 0 end as pron_balance_90_mas "
#                 +f"from riesgo.seguimiento_diario_mes a where a.dia= {dia} and a.buckets<>'i.Castigos') fn group by 1,2",conn)     

##Tabla5 se anexa bucket_cub
# df5 = pd.read_sql("select a.buckets_cub,sum(reserva),sum(reserva_cub) as s_reserva_cub "
#                 +"from riesgo.seguimiento_diario_mes_reservas b "
#                 +"left join riesgo.seguimiento_diario_mes a on a.folio=b.folio "
#                 +f"where a.dia= {dia} and b.dia= {dia} and a.buckets<>'i.Castigos' "
#                 +"group by 1 "
#                 +"order by 1",conn)

##Tabla6 MOVER 2 POR 1 : CIERRE BALANCE Y RESERVA###### 
df6 = pd.read_sql("select distinct bucket_nif,sum(balance) as balance,sum(reserva_nif) as reserva_nif "
                +"from riesgo.tmensual  "
                +f"where mes=/*202210*/to_char(date_trunc('month', current_date - interval '1' month)::date,'YYYYMM')::int and bucket_nif<>'i.180+' "
                +"group by 1  "                      
                +"order by 1",conn)

##Tabla7 MOVER 2 POR 1: CASTIGOS DEL MES###### 
df7 = pd.read_sql("select sum(piv.balance_ccc)/1000000 as capital_castigado_mes "
                +"from(select distinct a.*,b.dias_vencido as dias_vencido_ccc,(b.monto_operado-b.capital_pagado) as balance_ccc  "
                +"from(select distinct * from riesgo.tmensual "
                +f"where mes=/*202210*/to_char(date_trunc('month', current_date - interval '1' month)::date,'YYYYMM')::int and mora_cub<7  )a "
                +"left join cartera_corte_credito() b on a.folio=b.folio) piv "
                +"where piv.dias_vencido_ccc>=180",conn) 

# ##Tabla7 MOVER 2 POR 1: CASTIGOS DEL MES###### 
# df7 = pd.read_sql("select sum(piv.balance_ccc)/1000000 as capital_castigado_mes "
#                 +"from(select distinct a.*,b.dias_vencido as dias_vencido_ccc,(b.monto_operado-b.capital_pagado) as balance_ccc  "
#                 +"from(select distinct * from riesgo.tmensual "
#                 +f"where mes=to_char(date_trunc('month', current_date - interval '1' month)::date,'YYYYMM')::int and bucket not in ('k.180+'))a "
#                 +"left join cartera_corte_credito() b on a.folio=b.folio) piv "
#                 +"where piv.dias_vencido_ccc>=180",conn)                 

##Tabla8: TASA###### 
df8 = pd.read_sql("select distinct sum(fn.monto_tasa_ins)/sum(fn.monto_operado) as prom_tasa_insoluta "
                +"from(select distinct piv.*,c.tasa_insoluta_anual,(c.tasa_insoluta_anual*piv.monto_operado) as monto_tasa_ins "
                +"from (select fecha_reporte,fecha_contrato,cast(to_char(fecha_contrato, 'DD')as numeric) as dia_firma,folio,vintage,monto_operado "
                +"from cartera_corte_credito() ccc "
                +"where ccc.folio <> 'null' and ccc.id_status<>10 and ccc.fecha_contrato>=/*'2022-11-01'*/cast(date_trunc('month', current_date) as date )) piv "
                +"left join (select distinct a.folio, avg(a.tasa_diaria_interes)*365*100 tasa_insoluta_anual "
                +"from (select distinct folio, no_pago,tasa_diaria_interes from cobranza.pago_amortizacion "
                +"where tasa_diaria_interes>0 ) a "
               +"group by 1) c on piv.folio=c.folio) fn where fn.tasa_insoluta_anual is not null ",conn)     

###Tabla9: COMISION###### 
df9 = pd.read_sql("select (sum(comision_montoneto)/sum(monto_neto))*100 as comision_ponderada "
                +"from cobranza.base_originacion "  
                +"where fecha_contrato >=/*'2022-11-01'*/cast(date_trunc('month', current_date) as date )  and status<>'Reestructura  ' ",conn)      

###Tabla10:Vencido al cierre###
df10 = pd.read_sql("select distinct piv.cartera_vencida, sum(piv.balance) as balance  "
                +"from(select *,case when bucket_nif IN ('f.[91-120]','g.[121-150]','h.[151-179]','i.180+') then 'Vencido' "  
                +"else 'Vigente' "  
                +"end as cartera_vencida "  
                +"from riesgo.tmensual a "
                +f"where a.mes=/*202210*/to_char(date_trunc('month', current_date - interval '1' month)::date,'YYYYMM')::int and bucket_nif<>'i.180+') piv "
                +"group by 1 "
                +"order by 1 desc ",conn)    

###Tabla11:Vencido HOY###
df11 = pd.read_sql("select b.vencimiento_cub,sum(a.balance) as balance "
                +"from riesgo.seguimiento_diario_mes a "  
                +"left join (select distinct folio,vencimiento_cub from cartera_corte_credito_vencimiento() )b on a.folio=b.folio "  
                +f"where a.dia= {dia} and a.buckets<>'i.Castigos'  "  
                +"group by 1 "  
                +"order by 1 ",conn)   

#Tabla de concentracion para el mapa 
df12 = pd.read_sql("select fn.entidad_federativa_empresa,sum(fn.balance) as balance "
                +"from (select distinct a.fecha_reporte, a.folio, a.clasificacion,a.mora_cub,a.bucket_folio_ant ,a.balance,a.reserva_cub,cs.garantia, "
                +"case when cs.garantia is not null then 'Secured' "
                +"when a.clasificacion='Reestructurado' then a.clasificacion "
                +"else 'No reestructurado' end as producto, "
                +"case when bucket='k.180+' then 1 else 0 end as castigo, "
                +"case when a.clasificacion<>'Reestructurado' then a.mora_cub "
                +"when a.clasificacion='Reestructurado' and a.pago_sostenido='No' then greatest(a.mora_cub_f_ant,a.bucket_folio_ant) "
                +"when a.clasificacion='Reestructurado' and a.pago_sostenido='Si' then a.mora_cub end as mora_cub_fn, "
                +"vcp.entidad_federativa_empresa "
                +"from riesgo.seguimiento_diario_mes_reservas a "
                +"left join riesgo.creditos_secured cs on a.folio=cs.folio "
                +"left join v_colocacion_pld vcp on a.folio=vcp.contrato "
                +"where a.castigos is null and a.dia=6 and bucket<>'k.180+') fn "         
                +"group by 1 "                                                                                                                              
                +"order by 1 ",conn)                                                                  

# # Meter valores al plot1(Buckets)
# def addlabels(x,y):
#     for i in range(len(x)):
#         plt.text(i, y[i]//2, y[i], ha = 'center')
x=df["buckets_nif"]                                     ##---------Buckets o Moras
y=df["balance"]                                             ##---------Balance hoy

#Grafica1 (Barplot- Buckets)
# fig=plt.figure()
# plt.bar(x,y/1000000)
# addlabels(x,round(y/1000000))
# plt.xticks(x[::2])
# plt.yticks(fontsize=10)
# plt.xlabel("Buckets", fontsize=15)
# plt.ylabel("$ mdp", fontsize=15)
# plt.savefig("grafica_diaria.png")

#Grafica2 (Pie-Colocación)
##Hacer los colores iguales por Producto
my_data=df2["monto_operado_total"]
print(df2)
sum_monto_operado=sum (my_data)
my_labels=df2["status"]

def absolute_value(val):
    a  = round((numpy.round(val/100.*my_data.sum(), 0))/1000000,2)
    return a

#Asigna los colores por producto para la grafica 1
colores1 = [] #Arreglo donde se guardan los colores 

for p in my_labels:
    if p == 'Nuevo ':
        colores1.append('Green')
    if p == 'Refinanciamiento ':
        colores1.append('DeepSkyBlue')
    if p == 'Refinanciamiento Plus ':
        colores1.append('Blue')
    if p == 'Secured':
        colores1.append('Red')
    if p == 'Subsecuente ':
        colores1.append('DarkViolet')

fig2=plt.figure()
plt.pie(my_data,labels=my_labels,autopct=absolute_value, colors=colores1)
plt.title('Colocación Neta - mes anterior (mdp)')
plt.savefig("C:\Proyecto Diario\grafica_diaria3.png")##Ruta


#Grafica3 (Pie-Colocación)
##Hacer los colores iguales por Producto
my_data2=df3["monto_neto_total"]
sum_monto_neto=sum (my_data2)
my_labels2=df3["status"]
def absolute_value2(val):
    a  = round((numpy.round(val/100.*my_data2.sum(), 0))/1000000,2)
    return a

#Asigna los colores por porducto para la segunda grafica
colores2 = [] #arreglo 2 donde se guardan los colores

for p in my_labels2:
    if p == 'Nuevo ':
        colores2.append('Green')
    if p == 'Refinanciamiento ':
        colores2.append('DeepSkyBlue')
    if p == 'Refinanciamiento Plus ':
        colores2.append('Blue')
    if p == 'Secured':
        colores2.append('Red')
    if p == 'Subsecuente ':
        colores2.append('DarkViolet')    

fig3=plt.figure()   
plt.pie(my_data2,labels=my_labels2,autopct=absolute_value2, colors=colores2)
plt.title('Colocación Neta - mes actual (mdp)')
plt.savefig("C:\Proyecto Diario\grafica_diaria4.png") ##Ruta

#Datos Pronosticos
# vencido_hoy=df4["vencido_hoy"]
# icv_hoy=df4["icv"]
# vencido_pron=df4["vencido_pronostico"] #ahora cub
# icv_pron=df4["icv_pronostico"] #ahora cub

#Datos Reservas
# x2=df5["buckets_cub"]
# y2=df5["sum"]
z2=df["reserva_nif"] #Reserva CUB

#Datos cierre anterior
x3=df6["balance"]
# y3=df6["reserva"]
z3=df6["reserva_nif"]

##Datos Entrada a castigs
x4=df7["capital_castigado_mes"]
# print (x4[0])

##Datos Tasa Insoluta
x5=df8["prom_tasa_insoluta"]
# print (x5[0])

##Datos Comision
x6=df9["comision_ponderada"]
# print (x6[0])

#Nuevas variables cub
nc22=df10["balance"] #sum suM(vencido) CIERRE
n22=df11["balance"] #sum suM(vencido) hoy

#INICIA CODIGO DE MAPA
#Directorio donde está alojado el archivo shape
shape_file = 'C:\Proyecto Diario\gdb_ref_esta_ine_2009.shp' #Modificar la ruta
mexico_gdf = gpd.read_file(shape_file)

#Operaciones para obtener % de concentracion
df12 ['concentracion'] =(df12 ['balance'] / (df12 ['balance'].sum()) * 100)
df12 ['porcentaje'] = df12['concentracion'].round()

#Creacion tabla manual para extraer colores de df12 y agregarlos
df_state = pd.DataFrame()
df_state['entidad_federativa_empresa'] = ['CAMPECHE', 'BAJA CALIFORNIA SUR', 'BAJA CALIFORNIA', 'AGUASCALIENTES', 'ZACATECAS', 'YUCATAN', 'VERACRUZ DE IGNACIO DE LA LLAVE', 'TLAXCALA', 'TAMAULIPAS', 'TABASCO', 'SONORA', 'SINALOA', 'SAN LUIS POTOSI', 'QUINTANA ROO', 'QUERETARO', 'PUEBLA', 'OAXACA', 'NUEVO LEON', 'NAYARIT', 'MORELOS', 'MICHOACAN DE OCAMPO', 'MEXICO', 'JALISCO', 'HIDALGO', 'GUERRERO', 'GUANAJUATO', 'DURANGO', 'CIUDAD DE MEXICO', 'CHIHUAHUA', 'CHIAPAS', 'COLIMA', 'COAHUILA DE ZARAGOZA']
#seleccionar index para posterior comparacion
df12 = df12.set_index('entidad_federativa_empresa')
dfe = df_state.set_index('entidad_federativa_empresa')
#Metodo marge para unir tablas
comp = pd.merge(df12['porcentaje'], dfe , how = "right", on = 'entidad_federativa_empresa')
#Arreglo para leer etiquetas de tabla y recorrer mediante arreglo y asignar colores 
label_map=comp["porcentaje"]
colors_map = []

for n in label_map:
    if n < 5 :
        colors_map.append('dodgerblue')
    if n >=5 and n < 10 :
        colors_map.append('yellow')
    if n >=10 :
        colors_map.append('lime')
#Creacion de mapa con los colores de arreglo colors_map
mexico_gdf.plot(color=colors_map)
plt.title('Distribución de saldo activo')
plt.axis('off')
plt.savefig("C:\Proyecto Diario\mapa.png") #Modificar ruta

# print(my_data2)
# print(sum_monto_neto)

##Valores para definir tamano de graficas
# DPI = 60
# CARDSIZE = (int(2.49 * DPI), int(3.48 * DPI)) #Para definir tamaño de imagenes
##CANVAS
canvas = canvas.Canvas("reporte_semanal.pdf",pagesize=letter)
canvas.setLineWidth(.3) 

canvas.drawImage("C:\Proyecto Diario\creze.jpg", 500, 700)## Logo Creze

canvas.setFillColorRGB(0.4,0.6,0.4) ## COlor Creze
canvas.setFont('Times-Roman', 18)## Tamaño Creze
canvas.drawString(515,690,'creze')

canvas.setFont('Times-Roman', 15)
canvas.setFillColorRGB(0.1, 0.1, 0.1) #Negro
canvas.drawString(250,670,'Reporte Diario de Cartera')
canvas.setFont('Times-Roman', 13)
canvas.setFillColorRGB(0.1, 0.1, 0.1) #Negro
canvas.drawString(287,650,f"{today}")

canvas.drawString(33,615,f'Cifras mdp')
canvas.drawString(151,615,f'Cierre')
canvas.drawString(233,615,f'Hoy')
# canvas.drawString(283,615,f'Cierre')
# canvas.drawString(350,615,f'Hoy')
canvas.drawString(350,615,f'Cierre')
canvas.drawString(430,615,f'Hoy')

canvas.drawString(33,600,f'Buckets')
canvas.drawString(150,600,f'Balance')
canvas.drawString(233,600,f'Balance')
# canvas.drawString(260,600,f'Reserva')
# canvas.drawString(340,600,f'Reserva')
canvas.drawString(350,600,f'Res NIF')
canvas.drawString(430,600,f'Res NIF')
canvas.line(20,595,580,595)#Línea tabla

canvas.drawString(35,577,f'{x[0]}')                          #------MORA
canvas.drawString(150,577,"$"f'{round(x3[0]/1000000)}')             #------balance CIERRE
canvas.drawString(233,577,"$"f'{round(y[0]/1000000)}')              #------balance HOY
# canvas.drawString(270,577,"$"f'{round(y3[0]/1000000,1)}')
# canvas.drawString(345,577,"$"f'{round(y2[0]/1000000,1)}') #reserva bucket 1
canvas.drawString(350,577,"$"f'{round(z3[0]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,577,"$"f'{round(z2[0]/1000000,1)}')           #------reserva cub HOY
canvas.drawString(35,550,f'{x[1]}')                                 #------MORA
canvas.drawString(150,550,"$"f'{round(x3[1]/1000000,1)}')             #------balance CIERRE
canvas.drawString(233,550,"$"f'{round(y[1]/1000000,1)}')               #------balance HOY
# canvas.drawString(270,550,"$"f'{round(y3[1]/1000000,1)}')
# canvas.drawString(345,550,"$"f'{round(y2[1]/1000000,1)}')
canvas.drawString(350,550,"$"f'{round(z3[1]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,550,"$"f'{round(z2[1]/1000000,1)}')           #------reserva cub HOY
canvas.drawString(35,525,f'{x[2]}')                                 #------MORA
canvas.drawString(150,525,"$"f'{round(x3[2]/1000000,1)}')             #------balance CIERRE
canvas.drawString(233,525,"$"f'{round(y[2]/1000000,1)}')               #------balance HOY
# canvas.drawString(270,525,"$"f'{round(y3[2]/1000000,1)}')
# canvas.drawString(345,525,"$"f'{round(y2[2]/1000000,1)}')
canvas.drawString(350,525,"$"f'{round(z3[2]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,525,"$"f'{round(z2[2]/1000000,1)}')           #------reserva cub HOY
canvas.drawString(35,500,f'{x[3]}')                                 #------MORA
canvas.drawString(150,500,"$"f'{round(x3[3]/1000000,1)}')             #------balance CIERRE
canvas.drawString(233,500,"$"f'{round(y[3]/1000000,1)}')               #------balance HOY
# canvas.drawString(270,500,"$"f'{round(y3[3]/1000000,1)}')
# canvas.drawString(345,500,"$"f'{round(y2[3]/1000000,1)}')
canvas.drawString(350,500,"$"f'{round(z3[3]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,500,"$"f'{round(z2[3]/1000000,1)}')           #------reserva cub HOY
canvas.drawString(35,475,f'{x[4]}')                                 #------MORA
canvas.drawString(150,475,"$"f'{round(x3[4]/1000000,1)}')             #------balance CIERRE
canvas.drawString(233,475,"$"f'{round(y[4]/1000000,1)}')               #------balance HOY
# canvas.drawString(270,475,"$"f'{round(y3[4]/1000000,1)}')
# canvas.drawString(345,475,"$"f'{round(y2[4]/1000000,1)}')
canvas.drawString(350,475,"$"f'{round(z3[4]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,475,"$"f'{round(z2[4]/1000000,1)}')           #------reserva cub HOY
canvas.drawString(35,450,f'{x[5]}')                                 #------MORA
canvas.drawString(150,450,"$"f'{round(x3[5]/1000000,1)}')             #------balance CIERRE
canvas.drawString(233,450,"$"f'{round(y[5]/1000000,1)}')               #------balance HOY
# canvas.drawString(270,450,"$"f'{round(y3[5]/1000000,1)}')
# canvas.drawString(345,450,"$"f'{round(y2[5]/1000000,1)}')
canvas.drawString(350,450,"$"f'{round(z3[5]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,450,"$"f'{round(z2[5]/1000000,1)}')           #------reserva cub HOY
canvas.drawString(35,425,f'{x[6]}')                                 #------MORA
canvas.drawString(150,425,"$"f'{round(x3[6]/1000000,1)}')             #------balance CIERRE
canvas.drawString(233,425,"$"f'{round(y[6]/1000000,1)}')               #------balance HOY
# canvas.drawString(270,425,"$"f'{round(y3[6]/1000000,1)}')
# canvas.drawString(345,425,"$"f'{round(y2[6]/1000000,1)}')
canvas.drawString(350,425,"$"f'{round(z3[6]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,425,"$"f'{round(z2[6]/1000000,1)}')           #------reserva cub HOY
canvas.drawString(35,400,f'{x[7]}')                                 #------MORA
canvas.drawString(150,400,"$"f'{round(x3[7]/1000000,1)}')             #------balance CIERRE
canvas.drawString(233,400,"$"f'{round(y[7]/1000000,1)}')               #------balance HOY
# canvas.drawString(270,400,"$"f'{round(y3[7]/1000000,1)}')#270
# canvas.drawString(345,400,"$"f'{round(y2[7]/1000000,1)}')#270
canvas.drawString(350,400,"$"f'{round(z3[7]/1000000,1)}')           #------reserva CIERRE
canvas.drawString(430,400,"$"f'{round(z2[7]/1000000,1)}')#270         #------reserva cub HOY
canvas.line(20,390,580,390)#Línea abajo tabla
# canvas.line(20,420,580,420)#Línea abajo tabla
canvas.setFont('Times-Roman', 15)
# canvas.drawString(45,375,"Total")
canvas.drawString(45,370,"Total")
total_m1=x3[0]+x3[1]+x3[2]+x3[3]+x3[4]+x3[5]+x3[6]+x3[7]
# total_m2=y3[0]+y3[1]+y3[2]+y3[3]+y3[4]+y3[5]+y3[6]
total1=y[0]+y[1]+y[2]+y[3]+y[4]+y[5]+y[6]+y[7]
# total2=y2[0]+y2[1]+y2[2]+y2[3]+y2[4]+y2[5]+y2[6]
total3=z3[0]+z3[1]+z3[2]+z3[3]+z3[4]+z3[5]+z3[6]+z3[7]
total4=z2[0]+z2[1]+z2[2]+z2[3]+z2[4]+z2[5]+z2[6]+z2[7]
print(total1)
ic22=(n22[1]/total1)*100                                                            #ICV CUB hoy
ic22_ci=(nc22[1]/total_m1)*100                                                      #ICV CUB cierre
# cv_cierre=round((x3[4]+x3[5]+x3[6])/1000000,2) #CV normal cierre
# icv_cierre=round(((x3[4]+x3[5]+x3[6])/total_m1)*100,2) #CV normal cierre
# canvas.drawString(120,375,"$"f'{round(total_m1/1000000)}')#153
# canvas.drawString(265,375,"$"f'{round(total_m2/1000000,1)}')#153
# canvas.drawString(185,375,"$"f'{round(total1/1000000)}')#153
# canvas.drawString(342,375,"$"f'{round(total2/1000000,1)}')#267
# canvas.drawString(419,375,"$"f'{round(total3/1000000,1)}')#267
# canvas.drawString(493,375,"$"f'{round(total4/1000000,1)}')#267
canvas.drawString(150,370,"$"f'{round(total_m1/1000000)}')#153
# canvas.drawString(265,400,"$"f'{round(total_m2/1000000,1)}')#153
canvas.drawString(233,370,"$"f'{round(total1/1000000)}')#153
# canvas.drawString(342,400,"$"f'{round(total2/1000000,1)}')#267
canvas.drawString(350,370,"$"f'{round(total3/1000000,1)}')#267
canvas.drawString(430,370,"$"f'{round(total4/1000000,1)}')#267

##Cartera Vencida e ICV CIERRE
canvas.setFont('Times-Roman', 12)
canvas.setFillColorRGB(0.5, 0.1, 0.1) #Negro
# canvas.drawString(135,365,"Cart. Ven. normal") #
# canvas.drawString(145,353,"    (cierre)    ")  #
# # canvas.drawString(153,338,"$"f'{cv_cierre}')#Vencido cierre #
# canvas.drawString(240,365,"ICV ")  #
# canvas.drawString(240,353,"(cierre)")  #
# canvas.drawString(240,338,f'{icv_cierre}'"%")#ICV cierre  #

# canvas.drawString(300,365,"Cart. Ven. CUB")  #-15
# canvas.drawString(310,353,"     (cierre)   ")  #-15
# # canvas.drawString(316,338,"$"f'{round(nc22[1]/1000000,2)}')                             #VENCIDO CIERRE
# canvas.drawString(415,365,"ICV CUB") #-15
# canvas.drawString(410,353,"(cierre)")  #-15
# canvas.drawString(410,338,f'{round(ic22_ci,2)}'"%")                                     #ICV CIERRE

##Cartera Vencida e ICV HOOYY
canvas.setFont('Times-Roman', 12)
canvas.setFillColorRGB(0.5, 0.1, 0.1) #Negro
canvas.drawString(135,320,"Cart. Vencida") #-15
canvas.drawString(145,308,"    (cierre)    ")  #-15
canvas.drawString(153,293,"$"f'{round(nc22[1]/1000000,2)}')   
# canvas.drawString(153,293,"$"f'{vencido_hoy[0]}')#Vencido hoy #-15
canvas.drawString(240,320,"ICV ")  #-15
canvas.drawString(240,308,"(cierre)")  #-15
canvas.drawString(240,293,f'{round(ic22_ci,1)}'"%")       
# canvas.drawString(240,293,f'{icv_hoy[0]}'"%")#ICV hoy  #-15

canvas.drawString(300,320,"Cart. Vencida")  #-15
canvas.drawString(310,308,"     (hoy)   ")  #-15
canvas.drawString(316,293,"$"f'{round(n22[1]/1000000,2)}')                                  #Vencido  HOY
canvas.drawString(415,320,"ICV") #-15
canvas.drawString(410,308,"(hoy)")  #-15
canvas.drawString(410,293,f'{round(ic22,1)}'"%")                                            #ICV  HOY         

canvas.drawString(505,320,"Entrada a")
canvas.drawString(505,308,"Castigos")
canvas.drawString(505,293,"$"f'{round(x4[0],2)}')                                               #---Entrada a Castigos

#Graficas (Colocacion)
canvas.drawImage("C:\Proyecto Diario\grafica_diaria3.png", 30, 30, 300, 250,) ##Modificar la ruta
canvas.drawImage("C:\Proyecto Diario\grafica_diaria4.png", 300, 30, 300, 250,) ##Modificar la ruta

##Montos totales
canvas.setFillColorRGB(0.4,0.6,0.4) ## COlor Creze
canvas.setFont('Times-Roman', 18)## Tamaño Creze
canvas.drawString(165,27,"$"f'{round(sum_monto_operado/1000000,2)}')                                 #------MONTO OPERADO
canvas.drawString(435,27,"$"f'{round(sum_monto_neto/1000000,2)}')                                     # ------MONTO NETO

##Promedio Tasa Insoluta
canvas.setFont('Times-Roman', 12)## Tamaño Creze
canvas.drawString(270,60,"Prom Pond Tasa Insoluta")
canvas.setFont('Times-Roman', 15)## Tamaño Creze
canvas.drawString(300,45,f'{round(x5[0],2)}'"%")                                               #---TASA DEL MES

##Promedio Tasa Insoluta
canvas.setFont('Times-Roman', 12)## Tamaño Creze
canvas.drawString(270,25,"Prom Pond Comision")
canvas.setFont('Times-Roman', 15)## Tamaño Creze
canvas.drawString(300,10,f'{round(x6[0],2)}'"%")                                               #---Comision DEL MES
canvas.showPage()

#Comienza la segunda pagina
#Agregamos mapa & leyenda
canvas.drawImage("C:\Proyecto Diario\mapa.png", 50, 350, width=500, height=400) #Modificar ruta
canvas.drawImage("C:\Proyecto Diario\legends.png", 400, 600, width=140, height=50) #Modificar ruta
canvas.drawImage("C:\Proyecto Diario\creze.jpg", 500, 700)## Logo Creze  #Modificar ruta
canvas.setFillColorRGB(0.4,0.6,0.4) ## COlor Creze
canvas.setFont('Times-Roman', 18)## Tamaño Creze
canvas.drawString(503,690,'CREZE')
canvas.showPage()
canvas.save()

###AQUI
# Iniciamos los parámetros del script
remitente = 'cjuarez@creze.com'
# destinatarios = ['cjuarez@creze.com']
destinatarios = ['bprum@creze.com','stabares@creze.com','jahedo@creze.com','sgazca@creze.com','ecervantes@creze.com','caguilar@creze.com','apedroza@creze.com','celine@creze.com','jmartinez@creze.com','cjuarez@creze.com','asantiago@creze.com','pislas@creze.com']
asunto = f'Resumen Cartera {today}'
cuerpo = f'Hola a todos. \n\nPara el día {today} \n\nSe anexa seguimiento diario de la Cartera con los nuevos indicadores de la CUB \n\nAnexo liga de QS con mayor detalle: https://us-east-1.quicksight.aws.amazon.com/sn/dashboards/21aa4c17-05ab-4120-aa4d-7409230cc76c/views/a748f86e-b04b-4b0d-b624-65c7f27d178d \n\nSaludos \n\nCristian Juarez'
ruta_adjunto = 'reporte_semanal.pdf'
nombre_adjunto = 'reporte_semanal.pdf'

# # Creamos el objeto mensaje
mensaje = MIMEMultipart()

# # Establecemos los atributos del mensaje
mensaje['From'] = remitente
mensaje['To'] = ", ".join(destinatarios)
mensaje['Subject'] = asunto

# # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
mensaje.attach(MIMEText(cuerpo, 'plain'))

# # Abrimos el archivo que vamos a adjuntar
archivo_adjunto = open(ruta_adjunto, 'rb')

# # Creamos un objeto MIME base
adjunto_MIME = MIMEBase('application', 'octet-stream')
# Y le cargamos el archivo adjunto
adjunto_MIME.set_payload((archivo_adjunto).read())
# Codificamos el objeto en BASE64
encoders.encode_base64(adjunto_MIME)
# Agregamos una cabecera al objeto
adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
# Y finalmente lo agregamos al mensaje
mensaje.attach(adjunto_MIME)

# Creamos la conexión con el servidor
sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)

# Ciframos la conexión
sesion_smtp.starttls()

# Iniciamos sesión en el servidor
sesion_smtp.login('','') 

# Convertimos el objeto mensaje a texto
texto = mensaje.as_string()

# Enviamos el mensaje
sesion_smtp.sendmail(remitente, destinatarios, texto)

# Cerramos la conexión
sesion_smtp.quit()

print("Correo enviado con exito")
