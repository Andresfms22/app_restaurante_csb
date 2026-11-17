import streamlit as st
import pandas as pd
import os
import io
import zipfile
from datetime import datetime, date

from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# ==================================================
# CONFIGURACIÓN GENERAL Y ESTILOS VISUALES (CSS)
# ==================================================

if not os.path.exists("data"):
    os.makedirs("data")

st.set_page_config(
    page_title="Restaurante Pro",
    page_icon="🍽️",
    layout="wide"
)

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    [data-testid="stSidebar"] { background-color: #1e293b; color: #ffffff; }
    [data-testid="stSidebar"] span, [data-testid="stSidebar"] label, [data-testid="stSidebar"] div { color: #f1f5f9 !important; }
    h1, h2, h3 { color: #2c3e50; font-family: 'Helvetica Neue', sans-serif; }
    .stButton>button { border-radius: 8px; font-weight: 600; padding: 0.5rem 1rem; }
    </style>
""", unsafe_allow_html=True)

st.title("🍽️ Sistema Administrativo - Restaurante")
st.markdown("---")

anio_actual = datetime.today().year
MIN_DATE = date(anio_actual - 80, 1, 1)
MAX_DATE = date(anio_actual + 10, 12, 31)

def formato_cop(numero):
    try:
        return f"${int(numero):,.0f}".replace(",", ".")
    except:
        return "$0"

def limpiar_numero_texto(texto):
    if not texto:
        return 0
    limpio = str(texto).replace("$", "").replace(".", "").replace(",", "").strip()
    if limpio.isdigit():
        return int(limpio)
    return 0

def guardar_en_excel(nuevo_df, ruta_archivo):
    if os.path.exists(ruta_archivo):
        df_existente = pd.read_excel(ruta_archivo)
        df_final = pd.concat([df_existente, nuevo_df], ignore_index=True)
    else:
        df_final = nuevo_df
    df_final.to_excel(ruta_archivo, index=False)

def convertir_df_a_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# ==================================================
# MENÚ LATERAL
# ==================================================

st.sidebar.title("📌 Menú de Navegación")
opcion = st.sidebar.radio(
    "Seleccione módulo",
    [
        "Dashboard",
        "Cierre POS",
        "Empleados",
        "Proveedores (Facturas y Saldos)",
        "Seguros Médicos (Compartidos)",
        "Control de Inventarios",
        "Consulta de Gastos",
        "Consulta de Vales",
        "Consulta de Retiros",
        "Consulta Otros Ingresos",
        "Respaldo y Seguridad"
    ]
)

st.subheader(f"📂 Módulo: {opcion}")

# ==================================================
# DASHBOARD
# ==================================================

if opcion == "Dashboard":
    st.markdown("### 📊 Resumen Financiero General")

    gastos = pd.DataFrame()
    total_ventas, total_transferencias, total_domicilios = 0, 0, 0
    total_mesas, total_proveedores, total_gastos = 0, 0, 0
    total_vales, total_retiros, total_ingresos_extra = 0, 0, 0

    if os.path.exists("data/cierre_pos.xlsx"):
        cierre_pos = pd.read_excel("data/cierre_pos.xlsx")
        if "Total Ventas" in cierre_pos.columns: total_ventas = cierre_pos["Total Ventas"].fillna(0).sum()
        if "Transferencia" in cierre_pos.columns: total_transferencias = cierre_pos["Transferencia"].fillna(0).sum()
        if "Ventas Domicilio" in cierre_pos.columns: total_domicilios = cierre_pos["Ventas Domicilio"].fillna(0).sum()
        if "Ventas Mesas" in cierre_pos.columns: total_mesas = cierre_pos["Ventas Mesas"].fillna(0).sum()

    if os.path.exists("data/gastos.xlsx"):
        gastos = pd.read_excel("data/gastos.xlsx")
        if "Valor" in gastos.columns: total_gastos = gastos["Valor"].fillna(0).sum()
        if "Categoria" in gastos.columns and "Valor" in gastos.columns:
            prov_df = gastos[gastos["Categoria"] == "Proveedores"]
            total_proveedores = prov_df["Valor"].fillna(0).sum()

    if os.path.exists("data/vales.xlsx"):
        vales = pd.read_excel("data/vales.xlsx")
        if "Valor" in vales.columns: total_vales = vales["Valor"].fillna(0).sum()

    if os.path.exists("data/retiros.xlsx"):
        retiros = pd.read_excel("data/retiros.xlsx")
        if "Valor" in retiros.columns: total_retiros = retiros["Valor"].fillna(0).sum()

    if os.path.exists("data/ingresos.xlsx"):
        df_ingresos = pd.read_excel("data/ingresos.xlsx")
        if "Valor" in df_ingresos.columns: total_ingresos_extra = df_ingresos["Valor"].fillna(0).sum()

    saldo_neto = (total_ventas + total_ingresos_extra - total_gastos - total_vales - total_retiros)

    c1, c2 = st.columns(2)
    with c1: st.metric("🍽️ Ventas Acumuladas", formato_cop(total_ventas))
    with c2: st.metric("💸 Gastos Totales", formato_cop(total_gastos))

    c3, c4, c5 = st.columns(3)
    c3.metric("🧾 Vales Históricos", formato_cop(total_vales))
    c4.metric("💰 Otros Ingresos", formato_cop(total_ingresos_extra))
    c5.metric("👤 Retiros", formato_cop(total_retiros))

    c6, c7, c8, c9 = st.columns(4)
    c6.metric("💳 Transferencias", formato_cop(total_transferencias))
    c7.metric("🛵 Ventas Domicilio", formato_cop(total_domicilios))
    c8.metric("🍽️ Ventas Mesas", formato_cop(total_mesas))
    c9.metric("🚚 Proveedores Pagados", formato_cop(total_proveedores))
    
    st.markdown("---")
    st.markdown(f"### 📈 Saldo Neto Histórico (Utilidad Real): **{formato_cop(saldo_neto)}**")

# ==================================================
# CIERRE POS
# ==================================================

elif opcion == "Cierre POS":
    st.markdown("### 🧾 Ingreso Centralizado de Turno (Cierre POS)")
    
    lista_empleados = []
    if os.path.exists("data/empleados.xlsx"):
        df_emp = pd.read_excel("data/empleados.xlsx")
        if "Nombre" in df_emp.columns:
            lista_empleados = df_emp["Nombre"].dropna().unique().tolist()
        
    lista_proveedores = []
    if os.path.exists("data/proveedores.xlsx"):
        df_prov = pd.read_excel("data/proveedores.xlsx")
        if "Nombre" in df_prov.columns:
            lista_proveedores = df_prov["Nombre"].dropna().unique().tolist()

    medios_pago_lista = ["Efectivo", "Nequi", "Daviplata", "Transferencia", "Tarjeta"]

    col_ap1, col_ap2, col_ap3 = st.columns(3)
    fecha_cierre = col_ap1.date_input("Fecha del Cierre", value=datetime.today().date(), min_value=MIN_DATE, max_value=MAX_DATE)
    apertura = col_ap2.selectbox("Responsable del Turno", lista_empleados) if lista_empleados else col_ap2.text_input("Responsable del Turno")
    txt_base = col_ap3.text_input("💵 Base Inicial en Caja", value="0")
    base_efectivo = limpiar_numero_texto(txt_base)

    st.markdown("---")
    st.markdown("#### 1. INGRESOS Y VENTAS")
    col_in1, col_in2, col_in3 = st.columns(3)
    ventas_realizadas = limpiar_numero_texto(col_in1.text_input("Cantidad Tickets", value="0"))
    ventas_mesas = limpiar_numero_texto(col_in2.text_input("Ventas Mesas", value="0"))
    ventas_domicilio = limpiar_numero_texto(col_in3.text_input("Ventas Domicilio", value="0"))

    col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns(5)
    efectivo_pos = limpiar_numero_texto(col_p1.text_input("Ventas Efectivo", value="0"))
    nequi_pos = limpiar_numero_texto(col_p2.text_input("Ventas Nequi", value="0"))
    daviplata_pos = limpiar_numero_texto(col_p3.text_input("Ventas Daviplata", value="0"))
    tarjeta_pos = limpiar_numero_texto(col_p4.text_input("Ventas Tarjeta", value="0"))
    transferencia_pos = limpiar_numero_texto(col_p5.text_input("Transferencias", value="0"))
    total_pos = efectivo_pos + nequi_pos + daviplata_pos + tarjeta_pos + transferencia_pos

    propinas_retiradas = limpiar_numero_texto(st.text_input("💡 Retiro de Propinas", value="0"))

    st.markdown("---")
    st.markdown("#### 2. SALIDAS DE DINERO")
    col_salida1, col_salida2 = st.columns(2)
    
    vales_edit, prov_edit = pd.DataFrame(), pd.DataFrame()
    total_vales, total_proveedores = 0, 0

    with col_salida1:
        st.markdown("**Vales de Empleados**")
        if lista_empleados:
            vales_edit = st.data_editor(pd.DataFrame([{"Empleado": None, "Concepto": "", "Valor": 0, "Medio Pago": "Efectivo"}]), num_rows="dynamic", key="t_vales", hide_index=True, column_config={"Empleado": st.column_config.SelectboxColumn("Empleado", options=lista_empleados), "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista), "Valor": st.column_config.NumberColumn("Valor ($)")})
            total_vales = vales_edit["Valor"].fillna(0).sum() if not vales_edit.empty else 0

    with col_salida2:
        st.markdown("**Pago a Proveedores**")
        if lista_proveedores:
            prov_edit = st.data_editor(pd.DataFrame([{"Proveedor": None, "Detalle": "", "Valor": 0, "Medio Pago": "Efectivo"}]), num_rows="dynamic", key="t_prov", hide_index=True, column_config={"Proveedor": st.column_config.SelectboxColumn("Proveedor", options=lista_proveedores), "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista), "Valor": st.column_config.NumberColumn("Valor ($)")})
            total_proveedores = prov_edit["Valor"].fillna(0).sum() if not prov_edit.empty else 0

    st.markdown("**Rodamientos de Domiciliarios**")
    domiciliarios_edit = pd.DataFrame()
    if lista_empleados:
        domiciliarios_edit = st.data_editor(
            pd.DataFrame([{"Domiciliario": None, "Valor": 0, "Medio Pago": "Efectivo"}]),
            num_rows="dynamic", key="tabla_domiciliarios", hide_index=True,
            column_config={
                "Domiciliario": st.column_config.SelectboxColumn("Domiciliario", options=lista_empleados),
                "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista),
                "Valor": st.column_config.NumberColumn("Valor ($)")
            }
        )

    st.markdown("**Turnos Extras / Apoyo**")
    turnos_edit = pd.DataFrame()
    if lista_empleados:
        turnos_edit = st.data_editor(
            pd.DataFrame([{"Persona": None, "Detalle / Turno": "", "Valor": 0, "Medio Pago": "Efectivo"}]),
            num_rows="dynamic", key="tabla_turnos", hide_index=True,
            column_config={
                "Persona": st.column_config.SelectboxColumn("Persona", options=lista_empleados),
                "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista),
                "Valor": st.column_config.NumberColumn("Valor ($)")
            }
        )

    st.markdown("---")
    st.markdown("#### Retenciones y Otros Gastos")
    col_ret1, col_ret2, col_ret3 = st.columns(3)
    with col_ret1:
        txt_ahorro = st.text_input("Ahorro Impuesto", value="0", key="ahorro_impuesto_pos")
        ahorro_impuesto = limpiar_numero_texto(txt_ahorro)
        pago_ahorro = st.selectbox("Pagado con (Ahorro)", medios_pago_lista, index=0, key="pago_ahorro_sel")
    with col_ret2:
        txt_cadena = st.text_input("Cadena (Domingos)", value="0", key="cadena_domingo_pos")
        cadena_domingo = limpiar_numero_texto(txt_cadena)
        pago_cadena = st.selectbox("Pagado con (Cadena)", medios_pago_lista, index=0, key="pago_cadena_sel")
    with col_ret3:
        txt_otros = st.text_input("Otros Imprevistos", value="0", key="otros_imprevistos_pos")
        otros_pagos_pos = limpiar_numero_texto(txt_otros)
        pago_otros = st.selectbox("Pagado con (Otros)", medios_pago_lista, index=0, key="pago_otros_sel")

    egresos_efectivo = 0
    if not vales_edit.empty: egresos_efectivo += vales_edit[vales_edit["Medio Pago"] == "Efectivo"]["Valor"].fillna(0).sum()
    if not prov_edit.empty: egresos_efectivo += prov_edit[prov_edit["Medio Pago"] == "Efectivo"]["Valor"].fillna(0).sum()
    if not domiciliarios_edit.empty: egresos_efectivo += domiciliarios_edit[domiciliarios_edit["Medio Pago"] == "Efectivo"]["Valor"].fillna(0).sum()
    if not turnos_edit.empty: egresos_efectivo += turnos_edit[turnos_edit["Medio Pago"] == "Efectivo"]["Valor"].fillna(0).sum()
    
    if pago_ahorro == "Efectivo": egresos_efectivo += ahorro_impuesto
    if pago_cadena == "Efectivo": egresos_efectivo += cadena_domingo
    if pago_otros == "Efectivo": egresos_efectivo += otros_pagos_pos

    egresos_efectivo_total = egresos_efectivo + propinas_retiradas
    efectivo_teorico = base_efectivo + efectivo_pos - egresos_efectivo_total

    st.markdown("---")
    st.markdown("#### 3. CUADRE DE CAJA FÍSICA")
    efectivo_fisico = limpiar_numero_texto(st.text_input("💰 Efectivo físico contado en el cajón", value="0", key="efectivo_fisico_cierre"))
    diferencia_caja = efectivo_fisico - efectivo_teorico

    c1, c2, c3 = st.columns(3)
    c1.metric("Esperado", formato_cop(efectivo_teorico))
    c2.metric("Contado", formato_cop(efectivo_fisico))
    c3.metric("Diferencia", formato_cop(diferencia_caja))

    if diferencia_caja == 0: st.success("✅ CAJA CUADRADA")
    elif diferencia_caja > 0: st.warning(f"🟢 SOBRANTE: {formato_cop(diferencia_caja)}")
    else: st.error(f"🔴 FALTANTE: {formato_cop(abs(diferencia_caja))}")

    observaciones_pos = st.text_area("Observaciones del turno", key="observaciones_turno_cierre")

    def generar_pdf_en_memoria():
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer)
        pdf.drawString(50, 800, "CIERRE DE CAJA - RESTAURANTE")
        pdf.drawString(50, 780, f"Fecha: {fecha_cierre.strftime('%d/%m/%Y')} | Apertura: {apertura}")
        pdf.drawString(50, 760, f"Base Efectivo: {formato_cop(base_efectivo)}")
        pdf.drawString(50, 730, "--------------------------------------------------")
        pdf.drawString(50, 710, f"TOTAL VENTAS: {formato_cop(total_pos)}")
        pdf.drawString(50, 690, f"Efectivo Físico: {formato_cop(efectivo_fisico)}")
        pdf.drawString(50, 670, f"Diferencia: {formato_cop(diferencia_caja)}")
        pdf.save()
        buffer.seek(0)
        return buffer

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        if st.button("💾 Guardar Todo el Turno", type="primary", key="btn_guardar_turno_cierre"):
            guardar_en_excel(pd.DataFrame({
                "Fecha": [fecha_cierre], "Responsable": [apertura], "Base": [base_efectivo],
                "Ventas Mesas": [ventas_mesas], "Ventas Domicilio": [ventas_domicilio],
                "Total Ventas": [total_pos], "Efectivo Físico": [efectivo_fisico], "Diferencia": [diferencia_caja], "Observaciones": [observaciones_pos]
            }), "data/cierre_pos.xlsx")
            
            if not prov_edit.empty:
                prov_Val = prov_edit[(prov_edit["Valor"].fillna(0) > 0) & (prov_edit["Proveedor"].notna())]
                if not prov_Val.empty:
                    guardar_en_excel(pd.DataFrame({
                        "Fecha": [fecha_cierre] * len(prov_Val),
                        "Categoria": ["Proveedores"] * len(prov_Val),
                        "Descripcion": prov_Val["Proveedor"] + " - " + prov_Val["Detalle"].fillna(""),
                        "Valor": prov_Val["Valor"], "Medio Pago": prov_Val["Medio Pago"]
                    }), "data/gastos.xlsx")

            if not vales_edit.empty:
                vales_validos = vales_edit[(vales_edit["Valor"].fillna(0) > 0) & (vales_edit["Empleado"].notna())]
                if not vales_validos.empty:
                    guardar_en_excel(pd.DataFrame({
                        "Fecha": [fecha_cierre] * len(vales_validos),
                        "Empleado": vales_validos["Empleado"], "Concepto": vales_validos["Concepto"],
                        "Valor": vales_validos["Valor"], "Medio Pago": vales_validos["Medio Pago"],
                        "Estado": ["Pagado"] * len(vales_validos)
                    }), "data/vales.xlsx")

            gastos_adicionales = []
            if propinas_retiradas > 0: 
                gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Operación", "Descripcion": "Retiro de Propinas", "Valor": propinas_retiradas, "Medio Pago": "Efectivo"})
            
            if not domiciliarios_edit.empty:
                dom_validos = domiciliarios_edit[(domiciliarios_edit["Valor"].fillna(0) > 0) & (domiciliarios_edit["Domiciliario"].notna())]
                for _, row in dom_validos.iterrows():
                    gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Operación", "Descripcion": f"Rodamiento Domiciliario - {row['Domiciliario']}", "Valor": row["Valor"], "Medio Pago": row["Medio Pago"]})

            if not turnos_edit.empty:
                turnos_validos = turnos_edit[(turnos_edit["Valor"].fillna(0) > 0) & (turnos_edit["Persona"].notna())]
                for _, row in turnos_validos.iterrows():
                    gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Operación", "Descripcion": f"Turno Extra - {row['Persona']} ({row['Detalle / Turno']})", "Valor": row["Valor"], "Medio Pago": row["Medio Pago"]})

            if ahorro_impuesto > 0: gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Ahorro Impuesto", "Descripcion": "Fijo Diario", "Valor": ahorro_impuesto, "Medio Pago": pago_ahorro})
            if cadena_domingo > 0: gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Cadena", "Descripcion": "Recogida Dominical", "Valor": cadena_domingo, "Medio Pago": pago_cadena})
            if otros_pagos_pos > 0: gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Otros", "Descripcion": "Imprevistos", "Valor": otros_pagos_pos, "Medio Pago": pago_otros})

            if gastos_adicionales:
                guardar_en_excel(pd.DataFrame(gastos_adicionales), "data/gastos.xlsx")

            st.success("✅ ¡Turno, rodamientos, vales, retenciones y salidas guardados exitosamente!")

    with col_btn2:
        st.download_button(label="🖨️ Descargar Tirilla (PDF)", data=generar_pdf_en_memoria(), file_name=f"Cierre_{fecha_cierre.strftime('%Y%m%d')}.pdf", mime="application/pdf", key="btn_pdf_tirilla")

# ==================================================
# EMPLEADOS
# ==================================================

elif opcion == "Empleados":
    st.markdown("### 👥 Gestión de Empleados")

    with st.form("form_empleados", clear_on_submit=True):
        st.markdown("#### Registrar Nuevo Empleado")
        nombre = st.text_input("Nombre Completo")
        documento = st.text_input("Documento")
        cargo = st.text_input("Cargo")
        
        if st.form_submit_button("💾 Crear Empleado") and nombre:
            nuevo_empleado = pd.DataFrame({"Nombre": [nombre], "Documento": [documento], "Cargo": [cargo]})
            guardar_en_excel(nuevo_empleado, "data/empleados.xlsx")
            st.success("✅ Empleado guardado correctamente")

    if os.path.exists("data/empleados.xlsx"):
        st.markdown("---")
        st.markdown("#### 📋 Personal Activo (Edición Directa)")
        df_emp = pd.read_excel("data/empleados.xlsx")
        df_emp_editado = st.data_editor(df_emp, use_container_width=True, num_rows="dynamic", key="editor_empleados")
        
        if st.button("💾 Guardar Cambios en la Tabla de Empleados"):
            df_emp_editado.to_excel("data/empleados.xlsx", index=False)
            st.success("✅ Cambios de empleados guardados permanentemente.")

# ==================================================
# PROVEEDORES (FACTURAS Y SALDO CORRIENTE)
# ==================================================

elif opcion == "Proveedores (Facturas y Saldos)":
    st.markdown("### 🚚 Gestión de Proveedores y Cuenta Corriente")
    
    tab1, tab2, tab3 = st.tabs(["📋 Directorio", "📥 Registrar Factura / Despacho", "📊 Estado de Cuenta y Saldo Corriente"])
    
    with tab1:
        with st.form("f_prov", clear_on_submit=True):
            nombre_p = st.text_input("Nombre del Proveedor (Ej: Centro de Acopio)")
            tel_p = st.text_input("Teléfono")
            if st.form_submit_button("Guardar Proveedor") and nombre_p:
                guardar_en_excel(pd.DataFrame({"Nombre": [nombre_p], "Teléfono": [tel_p]}), "data/proveedores.xlsx")
                st.success(f"Proveedor '{nombre_p}' registrado.")
        
        if os.path.exists("data/proveedores.xlsx"):
            df_p = pd.read_excel("data/proveedores.xlsx")
            st.data_editor(df_p, use_container_width=True, key="ed_prov")

    with tab2:
        st.markdown("#### Registrar Factura o Mercancía Recibida (Deuda)")
        if os.path.exists("data/proveedores.xlsx"):
            progs = pd.read_excel("data/proveedores.xlsx")["Nombre"].dropna().tolist()
            if progs:
                with st.form("f_fact_prov"):
                    f_fec = st.date_input("Fecha de Factura", value=datetime.today().date())
                    f_pro = st.selectbox("Proveedor", progs)
                    f_num = st.text_input("Número de Factura / Remisión")
                    f_val = st.text_input("Valor de la Mercancía Despachada", value="0")
                    f_not = st.text_area("Detalle de la mercancía")
                    
                    if st.form_submit_button("Registrar Factura"):
                        val_num = limpiar_numero_texto(f_val)
                        guardar_en_excel(pd.DataFrame({
                            "Fecha": [f_fec], "Proveedor": [f_pro], "Nro Factura": [f_num],
                            "Valor Facturado": [val_num], "Detalle": [f_not]
                        }), "data/facturas_proveedores.xlsx")
                        st.success(f"Factura registrada por {formato_cop(val_num)}")
            else: st.warning("Primero crea un proveedor.")

    with tab3:
        st.markdown("#### ⚖️ Balance y Saldo Vivo con Proveedores")
        if os.path.exists("data/proveedores.xlsx"):
            progs = pd.read_excel("data/proveedores.xlsx")["Nombre"].dropna().tolist()
            if progs:
                sel_p = st.selectbox("Seleccione Proveedor para auditoría", progs, key="aud_p_filtrado")
                
                total_facturado = 0
                total_pagado = 0
                
                df_facs = pd.DataFrame()
                if os.path.exists("data/facturas_proveedores.xlsx"):
                    df_facs = pd.read_excel("data/facturas_proveedores.xlsx")
                    df_facs = df_facs[df_facs["Proveedor"].astype(str).str.contains(sel_p, case=False, na=False)]
                    total_facturado = df_facs["Valor Facturado"].sum() if "Valor Facturado" in df_facs.columns else 0
                
                df_pags = pd.DataFrame()
                if os.path.exists("data/gastos.xlsx"):
                    df_pags = pd.read_excel("data/gastos.xlsx")
                    df_pags = df_pags[(df_pags["Categoria"] == "Proveedores") & (df_pags["Descripcion"].astype(str).str.contains(sel_p, case=False, na=False))]
                    total_pagado = df_pags["Valor"].sum() if "Valor" in df_pags.columns else 0
                
                saldo_cuenta = total_facturado - total_pagado
                
                col1, col2, col3 = st.columns(3)
                col1.metric("📦 Total Facturado (Deuda)", formato_cop(total_facturado))
                col2.metric("💸 Total Pagado (Abonos)", formato_cop(total_pagado))
                
                if saldo_cuenta > 0:
                    col3.metric("🔴 Saldo Pendiente por Pagar", formato_cop(saldo_cuenta))
                elif saldo_cuenta < 0:
                    col3.metric("🟢 Saldo a Favor / Dinero Abonado", formato_cop(abs(saldo_cuenta)))
                else:
                    col3.metric("✨ Cuenta en Cero (Al Día)", "$0")
                
                st.markdown("##### 📄 Detalle de Facturas Recibidas")
                st.dataframe(df_facs, use_container_width=True, hide_index=True)
                
                st.markdown("##### 💸 Detalle de Pagos Realizados")
                st.dataframe(df_pags, use_container_width=True, hide_index=True)

# ==================================================
# SEGUROS MÉDICOS (COMPARTIDOS - DINÁMICOS)
# ==================================================

elif opcion == "Seguros Médicos (Compartidos)":
    st.markdown("### 🏥 Control de Seguros Médicos Compartidos")
    st.markdown("Registra el pago mensual global del seguro y utiliza la tabla para agregar o quitar tantas personas/sedes como necesites este mes.")

    f_fecha = st.date_input("Fecha del Pago del Seguro", value=datetime.today().date())
    f_concepto = st.text_input("Concepto / Póliza (Ej: Seguro médico mes de Diciembre)")

    st.markdown("**Distribución por persona / sede (Agrega o elimina filas con los botones de la tabla):**")
    
    df_seguro_base = pd.DataFrame([
        {"Persona / Sede": "Sede Principal", "Valor Asumido": 333333},
        {"Persona / Sede": "Persona 2", "Valor Asumido": 333333},
        {"Persona / Sede": "Persona 3", "Valor Asumido": 333334}
    ])

    tabla_division = st.data_editor(
        df_seguro_base,
        num_rows="dynamic",
        key="editor_seguros_dinamico",
        use_container_width=True,
        column_config={
            "Persona / Sede": st.column_config.TextColumn("Persona / Sede", required=True),
            "Valor Asumido": st.column_config.NumberColumn("Valor Asumido ($)", min_value=0, step=1000, required=True)
        }
    )

    total_calculado = tabla_division["Valor Asumido"].fillna(0).sum() if not tabla_division.empty else 0
    st.info(f"💡 **Valor Total del Seguro (Calculado de la tabla):** {formato_cop(total_calculado)}")

    if st.button("💾 Guardar Registro de Seguro", type="primary"):
        if not tabla_division.empty:
            registro_dict = {
                "Fecha": [f_fecha],
                "Concepto": [f_concepto],
                "Valor Total": [total_calculado]
            }
            for _, row in tabla_division.iterrows():
                nombre_col = str(row["Persona / Sede"]).strip()
                if nombre_col:
                    registro_dict[nombre_col] = [row["Valor Asumido"]]

            guardar_en_excel(pd.DataFrame(registro_dict), "data/seguros_medicos.xlsx")
            st.success("✅ ¡Seguro médico registrado con su división flexible exitosamente!")
        else:
            st.warning("⚠️ La tabla de distribución está vacía.")

    if os.path.exists("data/seguros_medicos.xlsx"):
        st.markdown("---")
        st.markdown("#### 📋 Historial de Seguros Médicos Registrados")
        df_seg = pd.read_excel("data/seguros_medicos.xlsx")
        st.dataframe(df_seg, use_container_width=True, hide_index=True)
        st.download_button("📥 Exportar Historial de Seguros", convertir_df_a_csv(df_seg), "Seguros_Medicos.csv", "text/csv")

# ==================================================
# CONTROL DE INVENTARIOS
# ==================================================

elif opcion == "Control de Inventarios":
    st.markdown("### 📦 Control de Inventarios (Saldos Diarios)")
    st.markdown("Alimenta manualmente las entradas por facturas y registra los saldos o consumos diarios de tus productos específicos.")

    with st.form("f_inv", clear_on_submit=True):
        f_inv_fec = st.date_input("Fecha de Registro", value=datetime.today().date())
        prod_nombre = st.text_input("Nombre del Producto (Ej: Carne, Queso, Envases)")
        tipo_mov = st.selectbox("Tipo de Movimiento", ["Entrada por Factura", "Salida / Consumo Diario", "Inventario Físico (Saldo Final)"])
        cantidad = st.number_input("Cantidad", min_value=0.0, step=1.0)
        observacion_inv = st.text_input("Notas / Factura asociada")
        
        if st.form_submit_button("Registrar en Inventario"):
            guardar_en_excel(pd.DataFrame({
                "Fecha": [f_inv_fec], "Producto": [prod_nombre], "Movimiento": [tipo_mov],
                "Cantidad": [cantidad], "Notas": [observacion_inv]
            }), "data/inventario.xlsx")
            st.success("✅ Movimiento de inventario guardado.")

    if os.path.exists("data/inventario.xlsx"):
        st.markdown("---")
        st.markdown("#### 📊 Historial y Estado de Inventarios")
        df_inv = pd.read_excel("data/inventario.xlsx")
        st.dataframe(df_inv, use_container_width=True, hide_index=True)
        
        if not df_inv.empty and "Producto" in df_inv.columns:
            st.markdown("#### 📐 Resumen de Stock por Producto")
            resumen_inv = df_inv.groupby(["Producto", "Movimiento"])["Cantidad"].sum().unstack(fill_value=0)
            st.dataframe(resumen_inv, use_container_width=True)

# ==================================================
# MÓDULOS DE CONSULTA
# ==================================================

elif opcion == "Consulta de Gastos":
    st.markdown("### 📋 Consulta General de Gastos")
    if os.path.exists("data/gastos.xlsx"):
        df_g = pd.read_excel("data/gastos.xlsx")
        st.dataframe(df_g, use_container_width=True, hide_index=True)
        st.download_button("📥 Exportar Gastos", convertir_df_a_csv(df_g), "Gastos.csv", "text/csv")
    else: st.info("No hay gastos registrados.")

elif opcion == "Consulta de Vales":
    st.markdown("### 📋 Historial y Filtro de Vales por Empleado")
    if os.path.exists("data/vales.xlsx"):
        df_vales = pd.read_excel("data/vales.xlsx")
        
        col_f1, _ = st.columns(2)
        empleados_lista = ["Todos"] + df_vales["Empleado"].dropna().unique().tolist() if "Empleado" in df_vales.columns else ["Todos"]
        filtro_emp = col_f1.selectbox("Filtrar por Empleado", empleados_lista, key="filtro_vales_emp")
        
        df_filtrado = df_vales.copy()
        if filtro_emp != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Empleado"] == filtro_emp]
            
        total_filtrado = df_filtrado["Valor"].sum() if "Valor" in df_filtrado.columns and not df_filtrado.empty else 0
        st.metric(f"💰 Total Vales ({filtro_emp})", formato_cop(total_filtrado))
        
        st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
        st.download_button("📥 Exportar Vales Filtrados", convertir_df_a_csv(df_filtrado), "Vales_Filtrados.csv", "text/csv")
    else: 
        st.info("No hay vales registrados todavía.")

elif opcion == "Consulta de Retiros":
    st.markdown("### 📋 Consulta de Retiros")
    if os.path.exists("data/retiros.xlsx"):
        st.dataframe(pd.read_excel("data/retiros.xlsx"), use_container_width=True, hide_index=True)
    else: st.info("No hay retiros.")

elif opcion == "Consulta Otros Ingresos":
    st.markdown("### 📋 Consulta de Otros Ingresos")
    if os.path.exists("data/ingresos.xlsx"):
        st.dataframe(pd.read_excel("data/ingresos.xlsx"), use_container_width=True, hide_index=True)
    else: st.info("No hay ingresos.")

# ==================================================
# RESPALDO Y SEGURIDAD
# ==================================================

elif opcion == "Respaldo y Seguridad":
    st.markdown("### 🔒 Respaldo de Datos (Backup)")
    def crear_zip_datos():
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for carpeta, _, archivos in os.walk("data"):
                for archivo in archivos:
                    zip_file.write(os.path.join(carpeta, archivo), arcname=archivo)
        buffer.seek(0)
        return buffer

    if os.path.exists("data"):
        st.download_button(label="📦 Descargar Respaldo Completo (ZIP)", data=crear_zip_datos(), file_name=f"Backup_{datetime.today().strftime('%Y%m%d')}.zip", mime="application/zip")
    else: 
        st.warning("No hay datos para respaldar.")
