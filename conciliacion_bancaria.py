import pandas as pd
import numpy as np
import re
import xlsxwriter

# --- Constantes de negocio ---
CHEQUE_DEV_OPERATIVO = CHEQUE_DEV_OPERATIVO
CHEQUE_RECHAZADO_CLEARING = CHEQUE_RECHAZADO_CLEARING
DATE_FMT = DATE_FMT
LABEL_NRO_CHEQUE = LABEL_NRO_CHEQUE


def leer_extracto(path):
    df = pd.read_excel(path, dtype=str)
    df.columns = df.columns.str.upper().str.strip()
    columnas = ["DIACONT", "MOVIMIENTO", "DESCRIP", "DEBE", "HABER", "SALDO"]  # <-- Agregar MOVIMIENTO
    df = df[columnas]
    for col in ["DEBE", "HABER", "SALDO"]:
        df[col] = df[col].fillna("0")
        df[col] = df[col].replace('[^0-9,]', '', regex=True)
        df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

def extraer_cheques_detallado(extracto_df):
    patrones = r"PAGO CHEQUE|CHEQUE DEP|CLEARING|CHEQUE RECHAZADO X CLEARING|CHEQUE DEV.OPERATIVO"
    cheques_df = extracto_df[extracto_df['DESCRIP'].str.contains(patrones, case=False, na=False)].copy()

    def buscar_nro_cheque(row):
        desc = str(row['DESCRIP']).upper()
        mov = str(row['MOVIMIENTO']) if 'MOVIMIENTO' in row else ""
        if desc.startswith('CLEARING REC') or CHEQUE_RECHAZADO_CLEARING in desc or CHEQUE_DEV_OPERATIVO in desc:
            partes = mov.split()
            if len(partes) >= 2 and partes[1].isdigit():
                return partes[1]
        match = re.search(r'(\d{5,})', desc)
        if match:
            return match.group(1)
        return ''

    def buscar_monto(row):
        desc = str(row['DESCRIP']).upper()
        if CHEQUE_RECHAZADO_CLEARING in desc or CHEQUE_DEV_OPERATIVO in desc:
            return row['HABER']
        else:
            return row['DEBE']

    cheques_df['NRO_CHEQUE'] = cheques_df.apply(buscar_nro_cheque, axis=1)
    cheques_df['MONTO'] = cheques_df.apply(buscar_monto, axis=1)
    resultado = cheques_df[['DIACONT', 'NRO_CHEQUE', 'MONTO', 'DESCRIP']]
    resultado = resultado.rename(columns={'DIACONT': 'FECHA'})
    return resultado


def leer_cheques_vista(path):
    df = pd.read_csv(path, encoding='latin1', sep=';')
    df.columns = df.columns.str.upper().str.strip()
    df['TOTAL'] = df['TOTAL'].astype(str).replace('[^0-9,]', '', regex=True).str.replace('.', '').str.replace(',', '.').astype(float)
    if 'FECHA MOVIMIENTO' in df.columns:
        df['FECHA'] = pd.to_datetime(df['FECHA MOVIMIENTO'], dayfirst=True, errors='coerce')
    return df

def estado_final_cheque(nro_cheque, historial_df):
    """
    Devuelve el estado final del cheque según la secuencia de eventos en el extracto.
    - Si el último evento es rechazo (o devolución operativa), devuelve 'pendiente'
    - Si el último evento es cobro, devuelve 'cobrado'
    """
    if historial_df.empty:
        return 'pendiente'
    eventos = []
    for _, row in historial_df.iterrows():
        desc = str(row['DESCRIP']).upper()
        if CHEQUE_RECHAZADO_CLEARING in desc or CHEQUE_DEV_OPERATIVO in desc:
            eventos.append('rechazado')
        elif ("PAGO CHEQUE" in desc) or ("CHEQUE DEP" in desc) or ("CLEARING" in desc):
            eventos.append('cobrado')
    if eventos:
        return eventos[-1]  # El último evento manda
    return 'pendiente'


def leer_cheques_diferidos(path):
    import pandas as pd
    fechas_cobro = []
    montos = []
    numeros = []
    ordenes = []

    with open(path, 'r', encoding='latin1', errors='ignore') as file:
        for line in file:
            if "che  dif" in line.lower() and ' e ' in line.lower():
                partes = line.strip().split()
                if len(partes) >= 10:
                    try:
                        fecha_cobro = partes[4]
                        nro_cheque = partes[5]
                        monto_str = partes[-1].replace('.', '').replace(',', '.')
                        monto = float(monto_str)
                        if monto > 0:
                            fechas_cobro.append(pd.to_datetime(fecha_cobro, format='%m/%d/%Y', errors='coerce'))
                            numeros.append(nro_cheque)
                            montos.append(monto)
                            orden = ' '.join(partes[7:-1])
                            ordenes.append(orden)
                    except:
                        continue

    return pd.DataFrame({
        "FECHA_COBRO": fechas_cobro,
        "IMPORTE": montos,
        "NRO": numeros,
        "ORDEN": ordenes
    })

def clasificar_ingreso(desc):
    prefijos = {
        "MOV.POS": "Infonet",
        "CR.COM.BEPSA": "Bepsa",
        "CRED. CABAL": "Cabal",
        "CRED. COMERCIO PANAL": "Panal",
        "DEPOSITO": "Depositos"
    }
    for prefijo, nombre in prefijos.items():
        if desc.upper().startswith(prefijo):
            return nombre
    return desc.strip()

def clasificar_egreso(desc):
    desc = desc.upper().strip()
    if desc.startswith("ATESORAMIENTO Y TRASLADO"):
        return "Prosegur"
    elif desc.startswith("DB X CUOTA"):
        return "Prestamo"
    elif desc.startswith("DEB.X TARJ"):
        return "Tarjeta de Credito"
    elif desc.startswith("DEV.INTRBN"):
        return "Devolucion Sipap"
    elif desc.startswith("MOV.POS.:BANCARD"):
        return "Alquiler POS Bancard"
    elif desc.startswith("DB.COM.BEPSA"):
        return "Alquiler POS Bepsa"
    elif desc.startswith("SET"):
        return "SET"
    elif desc.startswith("SEGUROS"):
        return "Seguros Pagados"
    elif desc.startswith("IPS"):
        return "IPS"
    return desc

def generar_excel_conciliacion(saldo_inicial, extracto_path, vista_path, diferidos_path, salida_excel):
    extracto_df = leer_extracto(extracto_path)
    cheques_en_extracto = extraer_cheques_detallado(extracto_df)
    cheques_vista_df = leer_cheques_vista(vista_path)
    cheques_diferidos_df = leer_cheques_diferidos(diferidos_path)

    extracto_df['CLASIFICACION'] = extracto_df['DESCRIP'].apply(clasificar_ingreso)
    ingresos_prioritarios = ["Depositos", "Infonet", "Bepsa", "Cabal", "Panal", "Bancard"]
    ingresos_sumados_df = extracto_df[extracto_df['CLASIFICACION'].isin(ingresos_prioritarios)]
    ingresos_sumados = ingresos_sumados_df.groupby('CLASIFICACION')['HABER'].sum().reset_index()
    ingresos_sumados['FECHA'] = ""

    ingresos_individuales_df = extracto_df[(extracto_df['HABER'] > 0) & (~extracto_df['CLASIFICACION'].isin(ingresos_prioritarios))]
    ingresos_individuales_df = ingresos_individuales_df[~ingresos_individuales_df['DESCRIP'].str.contains("CHEQUE DEVUELTO|RECHAZADO", case=False, na=False)]
    ingresos_individuales = ingresos_individuales_df[['CLASIFICACION', 'DIACONT', 'HABER']].rename(columns={
        'CLASIFICACION': 'NOMBRE', 'DIACONT': 'FECHA', 'HABER': 'MONTO'
    })

    ingresos_final_df = pd.concat([
        ingresos_sumados.rename(columns={'CLASIFICACION': 'NOMBRE', 'HABER': 'MONTO'})[['NOMBRE', 'FECHA', 'MONTO']],
        ingresos_individuales
    ], ignore_index=True)

    cheques_cobrados_df = extracto_df[
        extracto_df['DESCRIP'].str.contains(r"PAGO CHEQUE|CHEQUE DEP|CLEARING", case=False, na=False)
    ].copy()
    cheques_cobrados_df['NRO_CHEQUE'] = cheques_cobrados_df['DESCRIP'].str.extract(r'(\d{5,})', expand=False).astype(str).str.strip()

    cheques_vista_df['NRO'] = cheques_vista_df['NRO'].astype(str).str.strip()
    pendientes_vista = []
    for _, cheque in cheques_vista_df.iterrows():
        nro = str(cheque['NRO']).strip()
        historial = cheques_en_extracto[cheques_en_extracto['NRO_CHEQUE'] == nro]
        estado = estado_final_cheque(nro, historial)
        if estado in ['pendiente', 'rechazado']:
            pendientes_vista.append(cheque)
    cheques_pendientes_vista_df = pd.DataFrame(pendientes_vista)
    suma_pendientes_vista = cheques_pendientes_vista_df['TOTAL'].sum()

    cheques_diferidos_df['NRO'] = cheques_diferidos_df['NRO'].astype(str).str.strip()
    cheques_diferidos_df['IMPORTE_REDONDEADO'] = cheques_diferidos_df['IMPORTE'].round(-3)

    cheques_cobrados_df['NRO_CHEQUE'] = cheques_cobrados_df['NRO_CHEQUE'].astype(str).str.strip()
    cheques_cobrados_df['MONTO_REDONDEADO'] = cheques_cobrados_df['DEBE'].round(-3)


    pendientes_diferidos = []
    for _, cheque in cheques_diferidos_df.iterrows():
        nro = str(cheque['NRO']).strip()
        historial = cheques_en_extracto[cheques_en_extracto['NRO_CHEQUE'] == nro]
        estado = estado_final_cheque(nro, historial)
        if estado in ['pendiente', 'rechazado']:
            pendientes_diferidos.append(cheque)
    diferidos_no_cobrados = pd.DataFrame(pendientes_diferidos)
    cheques_pendientes_diferido = diferidos_no_cobrados['IMPORTE'].sum()

    egresos_df = extracto_df[(~extracto_df['DESCRIP'].str.contains("PAGO CHEQUE|CHEQUE DEP|CLEARING", case=False, na=False)) & (extracto_df['DEBE'] > 0)].copy()
    egresos_df['CLASIFICACION'] = egresos_df['DESCRIP'].apply(clasificar_egreso)

    # Orden personalizado para egresos
    orden_egresos = [
        "Cheques Emitidos", 
        "Cheque Adelantado (Diferidos)", 
        "Prestamo", 
        "Tarjeta de Credito", 
        "Alquiler POS Bepsa", 
        "Alquiler POS Bancard",
        "SET",
        "Seguros Pagados",
        "IPS",
        "Prosegur"
    ]
    egresos_df['ORDEN'] = egresos_df['CLASIFICACION'].apply(
        lambda x: orden_egresos.index(x) if x in orden_egresos else 999
    )
    egresos_otros = egresos_df[egresos_df['ORDEN'] == 999].sort_values('CLASIFICACION')
    egresos_ordenados = pd.concat([
        *[egresos_df[egresos_df['CLASIFICACION'] == o] for o in orden_egresos],
        egresos_otros
    ], ignore_index=True)

    total_cheques_vista = cheques_vista_df['TOTAL'].sum()
    total_cheques_diferidos = cheques_diferidos_df['IMPORTE'].sum()
    total_egresos = egresos_df['DEBE'].sum() + total_cheques_vista + total_cheques_diferidos
    total_ingresos = ingresos_final_df['MONTO'].sum()

    saldo_contable_final = saldo_inicial + total_ingresos - total_egresos
    saldo_bancario = saldo_contable_final + suma_pendientes_vista + cheques_pendientes_diferido
    saldo_final_extracto = extracto_df['SALDO'].iloc[-1]
    diferencia = saldo_bancario - saldo_final_extracto

    wb = xlsxwriter.Workbook(salida_excel)
    ws = wb.add_worksheet("Conciliacion")
    bold = wb.add_format({'bold': True})
    money = wb.add_format({'num_format': '#,##0'})

    row = 0
    ws.write(row, 0, "Saldo Contable Inicial", bold)
    ws.write(row, 3, saldo_inicial, money)
    row += 2

    ws.write(row, 0, "INGRESOS", bold)
    row += 1
    for _, r in ingresos_final_df.iterrows():
        ws.write(row, 0, r['NOMBRE'])
        ws.write(row, 1, str(r['FECHA']))
        ws.write(row, 2, r['MONTO'], money)
        row += 1
    ws.write(row, 0, "TOTAL INGRESOS", bold)
    ws.write(row, 2, total_ingresos, money)
    row += 2

    ws.write(row, 0, "EGRESOS", bold)
    row += 1
    ws.write(row, 0, "Cheques Emitidos", bold)
    ws.write(row, 3, total_cheques_vista, money)
    row += 1
    ws.write(row, 0, "Cheque Adelantado (Diferidos)", bold)
    ws.write(row, 3, total_cheques_diferidos, money)
    row += 1

    # Escribir los egresos en el orden especificado
    for _, r in egresos_ordenados.iterrows():
        ws.write(row, 0, r['CLASIFICACION'])
        ws.write(row, 1, r['DIACONT'])
        ws.write(row, 3, r['DEBE'], money)
        row += 1

    ws.write(row, 0, "TOTAL EGRESOS", bold)
    ws.write(row, 3, total_egresos, money)
    row += 2

    ws.write(row, 0, "Saldo Contable Final", bold)
    ws.write(row, 3, saldo_contable_final, money)
    row += 1
    extracto_df['FECHA_DT'] = pd.to_datetime(extracto_df['DIACONT'], format='%d/%m/%Y',dayfirst=True,errors='coerce')

    max_fecha = extracto_df['FECHA_DT'].max()
    if isinstance(max_fecha, pd.Timestamp) and not pd.isnull(max_fecha):
        mes_anio = max_fecha.strftime("%m/%Y")
    else:
        mes_anio = ''

    ws.write(row, 0, f"Cheques Pendientes Vista {mes_anio}", bold)
    ws.write(row, 3, suma_pendientes_vista, money)
    row += 1
    ws.write(row, 0, f"Cheques Pendientes Diferido {mes_anio}", bold)
    ws.write(row, 3, cheques_pendientes_diferido, money)
    row += 1
    ws.write(row, 0, "Saldo Bancario", bold)
    ws.write(row, 3, saldo_bancario, money)
    row += 1
    ws.write(row, 0, "Saldo Final del Extracto", bold)
    ws.write(row, 3, saldo_final_extracto, money)
    row += 1
    ws.write(row, 0, "Diferencia", bold)
    ws.write(row, 3, diferencia, money)

    # Hoja Cheques Pendientes Vista ORDENADA por número de cheque
    ws2 = wb.add_worksheet("Cheques Pendientes Vista")
    ws2.write(0, 0, "Fecha de Emision", bold)
    ws2.write(0, 1, LABEL_NRO_CHEQUE, bold)
    ws2.write(0, 2, "A la Orden De", bold)
    ws2.write(0, 3, "Monto", bold)
    cheques_pendientes_vista_df_ordenado = cheques_pendientes_vista_df.sort_values(by='NRO', key=lambda x: x.astype(str))
    for i, r in cheques_pendientes_vista_df_ordenado.reset_index(drop=True).iterrows():
        fecha = r.get('FECHA', '') or r.get('FECHA EMISION', '')
    # Si es Timestamp
        if isinstance(fecha, pd.Timestamp) and not pd.isnull(fecha):
            fecha = fecha.strftime('%d/%m/%Y')
    # Si es string, intento parsear
        elif isinstance(fecha, str) and fecha.strip():
            try:
                fecha_dt = pd.to_datetime(fecha, dayfirst=True, errors='coerce')
                if not pd.isnull(fecha_dt):
                    fecha = fecha_dt.strftime('%d/%m/%Y')
                else:
                    fecha = ''
            except:
             fecha = ''
        else:
            fecha = ''
        ws2.write(i + 1, 0, fecha)
        ws2.write(i + 1, 1, r.get('NRO', ''))
        ws2.write(i + 1, 2, r.get('ORDEN', ''))
        ws2.write(i + 1, 3, r['TOTAL'], money)


    # Hoja Cheques Pendientes Diferidos ORDENADA por número de cheque
    ws3 = wb.add_worksheet("Cheques Pendientes Diferidos")
    ws3.write(0, 0, "Fecha de Cobro", bold)
    ws3.write(0, 1, LABEL_NRO_CHEQUE, bold)
    ws3.write(0, 2, "A la Orden De", bold)
    ws3.write(0, 3, "Monto", bold)
    diferidos_no_cobrados_ordenado = diferidos_no_cobrados.sort_values(by='NRO', key=lambda x: x.astype(str))
    for i, r in diferidos_no_cobrados_ordenado.reset_index(drop=True).iterrows():
        fecha = r['FECHA_COBRO']
        if isinstance(fecha, pd.Timestamp) and not pd.isnull(fecha):
            fecha = fecha.strftime('%d/%m/%Y')
        else:
            fecha = ''
        ws3.write(i + 1, 0, fecha)
        ws3.write(i + 1, 1, r['NRO'])
        ws3.write(i + 1, 2, r.get('ORDEN', ''))
        ws3.write(i + 1, 3, r['IMPORTE'], money)


    # Hoja Cheques No Registrados ORDENADA por número de cheque
    # Hoja Cheques No Registrados ORDENADA por número de cheque

    cheques_vista_nros = set(cheques_vista_df['NRO'].astype(str).str.strip())
    cheques_diferidos_nros = set(cheques_diferidos_df['NRO'].astype(str).str.strip())

# Usamos cheques_en_extracto que ya tiene la extracción limpia
    cheques_no_registrados = cheques_en_extracto[
    ~cheques_en_extracto['NRO_CHEQUE'].isin(cheques_vista_nros.union(cheques_diferidos_nros))
    ]

    ws4 = wb.add_worksheet("Cheques No Registrados")
    ws4.write(0, 0, "Fecha", bold)
    ws4.write(0, 1, "Descripción", bold)
    ws4.write(0, 2, "Monto", bold)
    ws4.write(0, 3, LABEL_NRO_CHEQUE, bold)

    cheques_no_registrados_ordenado = cheques_no_registrados.sort_values(by='NRO_CHEQUE', key=lambda x: x.astype(str))
    for i, r in cheques_no_registrados_ordenado.reset_index(drop=True).iterrows():
        fecha = r['FECHA']
        if isinstance(fecha, pd.Timestamp) and not pd.isnull(fecha):
            fecha = fecha.strftime('%d/%m/%Y')
        elif isinstance(fecha, str) and fecha.strip():
            try:
                fecha_dt = pd.to_datetime(fecha, dayfirst=True, errors='coerce')
                if not pd.isnull(fecha_dt):
                    fecha = fecha_dt.strftime('%d/%m/%Y')
                else:
                    fecha = ''
            except:
                fecha = ''
        else:
            fecha = ''
        ws4.write(i + 1, 0, fecha)
        ws4.write(i + 1, 1, r['DESCRIP'])
        ws4.write(i + 1, 2, r['MONTO'], money)
        ws4.write(i + 1, 3, r['NRO_CHEQUE'])

    wb.close()
