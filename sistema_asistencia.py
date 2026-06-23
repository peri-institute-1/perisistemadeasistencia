#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv
import json
import smtplib
from email.message import EmailMessage

def configurar_google_sheets():
    load_dotenv()
    SHEET_ID = os.getenv('GOOGLE_SHEET_ID')
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    SERVICE_ACCOUNT_FILE = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE')
    if SERVICE_ACCOUNT_FILE and os.path.exists(SERVICE_ACCOUNT_FILE):
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    else:
        SERVICE_ACCOUNT_JSON = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
        if SERVICE_ACCOUNT_JSON:
            service_account_info = json.loads(SERVICE_ACCOUNT_JSON)
            creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
        else:
            raise ValueError("❌ No se encontraron credenciales de Google.")
    
    service = build('sheets', 'v4', credentials=creds)
    return service, SHEET_ID

def leer_hoja(service, sheet_id, nombre_hoja):
    try:
        result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=nombre_hoja).execute()
        values = result.get('values', [])
        if values:
            if len(values) > 1:
                headers = values[0]
                num_columns = len(headers)
                data_rows = [row + [''] * (num_columns - len(row)) if len(row) < num_columns else row for row in values[1:]]
                df = pd.DataFrame(data_rows, columns=headers)
            else:
                df = pd.DataFrame(values)
            return df
        return None
    except Exception as e:
        print(f"❌ Error al leer la hoja {nombre_hoja}: {e}")
        return None

def procesar_reportes_hoy():
    print("🚀 === PROCESANDO REPORTES PARA HOY ===")
    try:
        service, sheet_id = configurar_google_sheets()
        
        hoy = datetime.today()
        fecha_hoy_str = hoy.strftime("%d/%m/%Y")
        
        df_pagos_check = leer_hoja(service, sheet_id, 'PAGOSCHECK')
        df_pagos_data = leer_hoja(service, sheet_id, 'PAGOS')
        df_calendario = leer_hoja(service, sheet_id, 'REGISTRO_CALENDARIO')
        df_vendedoras = leer_hoja(service, sheet_id, 'VENDEDORAS')
        
        if any(df is None for df in [df_pagos_check, df_pagos_data, df_calendario, df_vendedoras]):
            return False
        
        df_pagos_check['fecha_pago_dt'] = pd.to_datetime(df_pagos_check['fecha_pago'], format="%d/%m/%Y", errors='coerce')
        hoy_dt = pd.to_datetime(hoy.strftime("%Y-%m-%d"))
        
        colaboradores_hoy = df_pagos_check[df_pagos_check['fecha_pago_dt'] == hoy_dt]
        
        if colaboradores_hoy.empty:
            print("ℹ️ No hay colaboradores con fecha de pago para hoy.")
            return True
            
        directorio_reportes = "Reportes_Asistencia"
        if not os.path.exists(directorio_reportes):
            os.makedirs(directorio_reportes)
            
        load_dotenv()
        gmail_user = os.getenv('GMAIL_USER')
        gmail_password = os.getenv('GMAIL_APP_PASSWORD')
        enviar_correos = bool(gmail_user and gmail_password)
        
        resultados = []
        
        for idx, row in colaboradores_hoy.iterrows():
            colaborador = str(row.get('Colaborador', '')).strip()
            fecha_pago = str(row.get('fecha_pago', '')).strip()
            periodo_inicio = str(row.get('periodo_inicio', '')).strip()
            periodo_fin = str(row.get('periodo_fin', '')).strip()
            
            try:
                fecha_inicio_dt = datetime.strptime(periodo_inicio, "%d/%m/%Y")
                fecha_fin_dt = datetime.strptime(periodo_fin, "%d/%m/%Y")
                
                df_colaborador = df_calendario[df_calendario['Colaborador'].astype(str).str.strip() == colaborador].copy()
                if df_colaborador.empty:
                    continue
                
                df_colaborador['FechaEntrada_dt'] = pd.to_datetime(df_colaborador['FechaEntrada'], format="%d/%m/%Y", errors='coerce')
                df_asistencia = df_colaborador[
                    (df_colaborador['FechaEntrada_dt'] >= fecha_inicio_dt) & 
                    (df_colaborador['FechaEntrada_dt'] <= fecha_fin_dt)
                ].drop(columns=['FechaEntrada_dt'])
                
                if df_asistencia.empty:
                    continue
                
                # --- MAGIA DEL RESUMEN EN EL EXCEL ---
                pago_match = df_pagos_data[
                    (df_pagos_data['Colaborador'].astype(str).str.strip() == colaborador) &
                    (df_pagos_data['periodo_inicio'].astype(str).str.strip() == periodo_inicio)
                ]
                
                horas_totales = 0
                monto_total = 0
                moneda = "Soles"
                
                if not pago_match.empty:
                    reg = pago_match.iloc[0]
                    h_norm = float(str(reg.get('horas_normales', '0')).replace(',', '.'))
                    h_ext = float(str(reg.get('horas_extra', '0')).replace(',', '.'))
                    horas_totales = h_norm + h_ext
                    monto_total = float(str(reg.get('monto_total', '0')).replace(',', '.'))
                    moneda = str(reg.get('moneda', 'Soles')).strip()
                
                columnas = list(df_asistencia.columns)
                fila_vacia = {col: '' for col in columnas}
                fila_titulo = {col: '' for col in columnas}
                # Corrección del ERROR en Excel usando guiones
                fila_titulo[columnas[0]] = '--- RESUMEN DEL PERIODO ---'
                
                fila_horas = {col: '' for col in columnas}
                fila_horas[columnas[0]] = 'TOTAL HORAS:'
                fila_horas[columnas[1]] = f"{round(horas_totales, 2)} hrs"
                
                fila_monto = {col: '' for col in columnas}
                fila_monto[columnas[0]] = 'MONTO TOTAL:'
                fila_monto[columnas[1]] = f"{round(monto_total, 2)} {moneda}"
                
                df_resumen = pd.DataFrame([fila_vacia, fila_titulo, fila_horas, fila_monto])
                df_final = pd.concat([df_asistencia, df_resumen], ignore_index=True)
                # ----------------------------------------

                nombre_archivo = f"Reporte_{colaborador.replace(' ', '_')}_{fecha_pago.replace('/', '-')}.xlsx"
                ruta_archivo = os.path.join(directorio_reportes, nombre_archivo)
                df_final.to_excel(ruta_archivo, index=False, engine='openpyxl')
                
                estado_correo = "No configurado"
                if enviar_correos:
                    v_match = df_vendedoras[df_vendedoras['Colaborador'].astype(str).str.strip() == colaborador]
                    if not v_match.empty and str(v_match.iloc[0].get('Correo', '')).strip() != 'nan':
                        email_colab = str(v_match.iloc[0]['Correo']).strip()
                        if enviar_correo_con_excel(email_colab, colaborador, ruta_archivo, fecha_pago, gmail_user, gmail_password):
                            estado_correo = "Enviado"
                        else:
                            estado_correo = "Error"
                
                resultados.append({
                    'Colaborador': colaborador,
                    'Fecha_Pago': fecha_pago,
                    'Archivo_Excel': nombre_archivo,
                    'Estado_Correo': estado_correo,
                    'Periodo': f"{periodo_inicio} - {periodo_fin}"
                })
                
            except Exception as e:
                print(f"❌ Error con {colaborador}: {e}")
                
        if resultados:
            resumen_horas = calcular_resumen_horas(resultados, df_pagos_data)
            enviar_resumen_administrativo(resumen_horas, fecha_hoy_str, resultados, gmail_user, gmail_password)
            
        return True
    except Exception as e:
        print(f"❌ Error general: {e}")
        return False

def enviar_correo_con_excel(destinatario, colaborador, archivo_excel, fecha_pago, user, password):
    try:
        msg = EmailMessage()
        msg['Subject'] = f'Reporte de Asistencia - {colaborador} ({fecha_pago})'
        msg['From'] = user
        msg['To'] = destinatario
        
        cuerpo = f"Estimado/a {colaborador},\n\nTe enviamos tu reporte de asistencia correspondiente al periodo de pago: {fecha_pago}.\nAl final del archivo adjunto encontrarás el resumen de tus horas totales y el monto a pagar.\n\nSaludos,\nEquipo Peri Company"
        msg.set_content(cuerpo)
        
        with open(archivo_excel, 'rb') as f:
            msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=os.path.basename(archivo_excel))
            
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(user, password)
            server.send_message(msg)
        return True
    except:
        return False

def calcular_resumen_horas(resultados, df_pagos_data):
    resumen = {}
    for res in resultados:
        colaborador = res['Colaborador']
        per_str = res['Periodo'].split(" - ")
        pago_match = df_pagos_data[
            (df_pagos_data['Colaborador'].astype(str).str.strip() == colaborador) &
            (df_pagos_data['periodo_inicio'].astype(str).str.strip() == per_str[0])
        ]
        if pago_match.empty: continue
        reg = pago_match.iloc[0]
        h_norm = float(str(reg.get('horas_normales', '0')).replace(',', '.'))
        h_ext = float(str(reg.get('horas_extra', '0')).replace(',', '.'))
        monto = float(str(reg.get('monto_total', '0')).replace(',', '.'))
        moneda = str(reg.get('moneda', 'Soles')).strip()
        resumen[colaborador] = {'horas': h_norm + h_ext, 'monto': monto, 'moneda': moneda}
    return resumen

def enviar_resumen_administrativo(resumen, fecha_hoy_str, resultados, user, password):
    if not user or not password: return False
    try:
        msg = EmailMessage()
        msg['Subject'] = f'📊 Resumen de Reportes de Asistencia - {fecha_hoy_str}'
        msg['From'] = user
        msg['To'] = 'nitza.peri.d@gmail.com'
        cuerpo = f"Se han procesado los reportes de la fecha: {fecha_hoy_str}\n\nResumen de Pagos:\n"
        for col, dat in resumen.items():
            cuerpo += f"• {col}: {round(dat['horas'], 2)} horas -> {round(dat['monto'], 2)} {dat['moneda']}\n"
        cuerpo += "\nSistema Automatizado de Asistencia."
        msg.set_content(cuerpo)
        for res in resultados:
            ruta = os.path.join("Reportes_Asistencia", res['Archivo_Excel'])
            if os.path.exists(ruta):
                with open(ruta, 'rb') as f:
                    msg.add_attachment(f.read(), maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=res['Archivo_Excel'])
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(user, password)
            server.send_message(msg)
        return True
    except:
        return False

if __name__ == "__main__":
    if procesar_reportes_hoy():
        exit(0)
    else:
        exit(1)
