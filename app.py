import streamlit as st
import pandas as pd
import os
import io
import zipfile
from datetime import datetime

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

# Inyección de Estilos CSS corregidos (Sidebar oscuro para que el texto resalte perfecto)
st.markdown("""
    <style>
    /* Estilo global */
    .main {
        background-color: #f8f9fa;
    }
    
    /* Sidebar estilizado en tono oscuro moderno para contraste perfecto */
    [data-testid="stSidebar"] {
        background-color: #1e293b;
        color: #ffffff;
    }
    
    /* Forzar que los textos del sidebar sean blancos y legibles */
    [data-testid="stSidebar"] span, [data-testid="stSidebar"] label, [data-testid="stSidebar"] div {
        color: #f1f5f9 !important;
    }
    
    /* Encabezados y títulos */
    h1, h2, h3 {
        color: #2c3e50;
        font-family: 'Helvetica Neue', sans-serif;
    }
    
    /* Botones personalizados */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5rem 1rem;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🍽️ Sistema Administrativo - Restaurante")
st.markdown("---")

def formato_cop(numero):
    return f"${numero:,.0f}".replace(",", ".")

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
        "Proveedores",
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

    total_ventas, total_transferencias, total_domicilios = 0, 0, 0
    total_mesas, total_proveedores, total_gastos = 0, 0, 0
    total_vales, total_retiros, total_ingresos_extra = 0, 0, 0

    if os.path.exists("data/cierre_pos.xlsx"):
        cierre_pos = pd.read_excel("data/cierre_pos.xlsx")
        if "Total Ventas" in cierre_pos.columns: total_ventas = cierre_pos["Total Ventas"].sum()
        if "Transferencia" in cierre_pos.columns: total_transferencias = cierre_pos["Transferencia"].sum()
        if "Ventas Domicilio" in cierre_pos.columns: total_domicilios = cierre_pos["Ventas Domicilio"].sum()
        if "Ventas Mesas" in cierre_pos.columns: total_mesas = cierre_pos["Ventas Mesas"].sum()

    if os.path.exists("data/gastos.xlsx"):
        gastos = pd.read_excel("data/gastos.xlsx")
        if "Valor" in gastos.columns: total_gastos = gastos["Valor"].sum()
        if "Categoria" in gastos.columns:
            prov_df = gastos[gastos["Categoria"] == "Proveedores"]
            total_proveedores = prov_df["Valor"].sum()

    if os.path.exists("data/vales.xlsx"):
        vales = pd.read_excel("data/vales.xlsx")
        if "Valor" in vales.columns: total_vales = vales["Valor"].sum()

    if os.path.exists("data/retiros.xlsx"):
        retiros = pd.read_excel("data/retiros.xlsx")
        if "Valor" in retiros.columns: total_retiros = retiros["Valor"].sum()

    if os.path.exists("data/ingresos.xlsx"):
        df_ingresos = pd.read_excel("data/ingresos.xlsx")
        if "Valor" in df_ingresos.columns: total_ingresos_extra = df_ingresos["Valor"].sum()

    saldo_neto = (total_ventas + total_ingresos_extra - total_gastos - total_vales - total_retiros)

    c1, c2 = st.columns(2)
    with c1:
        st.metric("🍽️ Ventas Acumuladas", formato_cop(total_ventas))
    with c2:
        st.metric("💸 Gastos Totales", formato_cop(total_gastos))

    c3, c4, c5 = st.columns(3)
    c3.metric("🧾 Vales Históricos", formato_cop(total_vales))
    c4.metric("💰 Otros Ingresos", formato_cop(total_ingresos_extra))
    c5.metric("👤 Retiros", formato_cop(total_retiros))

    c6, c7, c8, c9 = st.columns(4)
    c6.metric("💳 Transferencias", formato_cop(total_transferencias))
    c7.metric("🛵 Ventas Domicilio", formato_cop(total_domicilios))
    c8.metric("🍽️ Ventas Mesas", formato_cop(total_mesas))
    c9.metric("🚚 Proveedores", formato_cop(total_proveedores))
    
    st.markdown("---")
    st.markdown(f"### 📈 Saldo Neto Histórico: **{formato_cop(saldo_neto)}**")

    st.markdown("### 📉 Distribución de Gastos por Categoría")
    if os.path.exists("data/gastos.xlsx") and not gastos.empty:
        if "Categoria" in gastos.columns and "Valor" in gastos.columns:
            gastos_por_cat = gastos.groupby("Categoria")["Valor"].sum()
            st.bar_chart(gastos_por_cat)
    else:
        st.info("Aún no hay suficientes registros de gastos para graficar.")


# ==================================================
# CIERRE POS (INGRESO CENTRALIZADO)
# ==================================================

elif opcion == "Cierre POS":
    st.markdown("### 🧾 Ingreso Centralizado de Turno (Cierre POS)")
    
    lista_empleados = []
    if os.path.exists("data/empleados.xlsx"):
        lista_empleados = pd.read_excel("data/empleados.xlsx")["Nombre"].dropna().unique().tolist()
        
    lista_proveedores = []
    if os.path.exists("data/proveedores.xlsx"):
        lista_proveedores = pd.read_excel("data/proveedores.xlsx")["Nombre"].dropna().unique().tolist()

    medios_pago_lista = ["Efectivo", "Nequi", "Daviplata", "Transferencia", "Tarjeta"]

    col_ap1, col_ap2, col_ap3 = st.columns(3)
    fecha_cierre = col_ap1.date_input("Fecha del Cierre")
    apertura = col_ap2.selectbox("Responsable del Turno", lista_empleados) if lista_empleados else col_ap2.text_input("Responsable del Turno")
    base_efectivo = col_ap3.number_input("💵 Base Inicial en Caja", min_value=0, step=1000)

    st.markdown("---")
    st.markdown("#### 1. INGRESOS Y VENTAS")
    col_in1, col_in2, col_in3 = st.columns(3)
    ventas_realizadas = col_in1.number_input("Cantidad Tickets", min_value=0)
    ventas_mesas = col_in2.number_input("Ventas Mesas", min_value=0, step=1000)
    ventas_domicilio = col_in3.number_input("Ventas Domicilio", min_value=0, step=1000)

    st.markdown("**Desglose por Método de Pago**")
    col_p1, col_p2, col_p3, col_p4, col_p5 = st.columns(5)
    efectivo_pos = col_p1.number_input("Ventas Efectivo", min_value=0, step=1000)
    nequi_pos = col_p2.number_input("Ventas Nequi", min_value=0, step=1000)
    daviplata_pos = col_p3.number_input("Ventas Daviplata", min_value=0, step=1000)
    tarjeta_pos = col_p4.number_input("Ventas Tarjeta", min_value=0, step=1000)
    transferencia_pos = col_p5.number_input("Transferencias", min_value=0, step=1000)
    
    total_pos = efectivo_pos + nequi_pos + daviplata_pos + tarjeta_pos + transferencia_pos

    st.markdown("#### 💡 PROPINAS DEL TURNO")
    propinas_retiradas = st.number_input("Retiro de Propinas (Dinero entregado del cajón)", min_value=0, step=1000)

    st.markdown("---")
    st.markdown("#### 2. SALIDAS DE DINERO (Egresos y Retenciones)")
    
    col_salida1, col_salida2 = st.columns(2)
    
    total_vales = 0
    vales_edit = pd.DataFrame()
    with col_salida1:
        st.markdown("**Vales de Empleados**")
        if not lista_empleados:
            st.error("⚠️ Registra personal en el módulo 'Empleados'.")
        else:
            df_vales_base = pd.DataFrame([{"Empleado": None, "Concepto": "", "Valor": 0, "Medio Pago": "Efectivo"}])
            vales_edit = st.data_editor(
                df_vales_base, num_rows="dynamic", key="tabla_vales", hide_index=True,
                column_config={
                    "Empleado": st.column_config.SelectboxColumn("Empleado", options=lista_empleados, required=True),
                    "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista, required=True)
                }
            )
            total_vales = vales_edit["Valor"].sum()
            st.caption(f"Total Vales: {formato_cop(total_vales)}")

    total_proveedores = 0
    prov_edit = pd.DataFrame()
    with col_salida2:
        st.markdown("**Pago a Proveedores**")
        if not lista_proveedores:
            st.error("⚠️ Registra proveedores en su módulo.")
        else:
            df_prov_base = pd.DataFrame([{"Proveedor": None, "Detalle": "", "Valor": 0, "Medio Pago": "Efectivo"}])
            prov_edit = st.data_editor(
                df_prov_base, num_rows="dynamic", key="tabla_prov", hide_index=True,
                column_config={
                    "Proveedor": st.column_config.SelectboxColumn("Proveedor", options=lista_proveedores, required=True),
                    "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista, required=True)
                }
            )
            total_proveedores = prov_edit["Valor"].sum()
            st.caption(f"Total Proveedores: {formato_cop(total_proveedores)}")

    st.markdown("**Rodamientos de Domiciliarios**")
    if not lista_empleados:
        domiciliarios_edit = pd.DataFrame()
        total_rodamientos = 0
    else:
        df_dom_base = pd.DataFrame([{"Domiciliario": None, "Valor": 0, "Medio Pago": "Efectivo"}])
        domiciliarios_edit = st.data_editor(
            df_dom_base, num_rows="dynamic", key="tabla_domiciliarios", hide_index=True,
            column_config={
                "Domiciliario": st.column_config.SelectboxColumn("Domiciliario", options=lista_empleados, required=True),
                "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista, required=True)
            }
        )
        total_rodamientos = domiciliarios_edit["Valor"].sum()
        st.caption(f"Total Rodamientos: {formato_cop(total_rodamientos)}")

    st.markdown("**Turnos Extras / Apoyo (Múltiples personas)**")
    if not lista_empleados:
        turnos_edit = pd.DataFrame()
        total_turnos = 0
    else:
        df_turnos_base = pd.DataFrame([{"Persona": None, "Detalle / Turno": "", "Valor": 0, "Medio Pago": "Efectivo"}])
        turnos_edit = st.data_editor(
            df_turnos_base, num_rows="dynamic", key="tabla_turnos", hide_index=True,
            column_config={
                "Persona": st.column_config.SelectboxColumn("Persona", options=lista_empleados, required=True),
                "Medio Pago": st.column_config.SelectboxColumn("Medio Pago", options=medios_pago_lista, required=True)
            }
        )
        total_turnos = turnos_edit["Valor"].sum()
        st.caption(f"Total Turnos Extras: {formato_cop(total_turnos)}")

    st.markdown("**Retenciones y Otros Gastos**")
    col_ret1, col_ret2, col_ret3 = st.columns(3)
    with col_ret1:
        ahorro_impuesto = st.number_input("Ahorro Impuesto (Sugerido 100.000)", value=0, step=1000)
        pago_ahorro = st.selectbox("Pagado con (Ahorro)", medios_pago_lista, index=0)
    with col_ret2:
        cadena_domingo = st.number_input("Cadena (Domingos)", value=0, step=1000)
        pago_cadena = st.selectbox("Pagado con (Cadena)", medios_pago_lista, index=0)
    with col_ret3:
        otros_pagos_pos = st.number_input("Otros Imprevistos", min_value=0, step=1000)
        pago_otros = st.selectbox("Pagado con (Otros)", medios_pago_lista, index=0)

    egresos_efectivo = 0
    if not vales_edit.empty:
        egresos_efectivo += vales_edit[vales_edit["Medio Pago"] == "Efectivo"]["Valor"].sum()
    if not prov_edit.empty:
        egresos_efectivo += prov_edit[prov_edit["Medio Pago"] == "Efectivo"]["Valor"].sum()
    if not domiciliarios_edit.empty:
        egresos_efectivo += domiciliarios_edit[domiciliarios_edit["Medio Pago"] == "Efectivo"]["Valor"].sum()
    if not turnos_edit.empty:
        egresos_efectivo += turnos_edit[turnos_edit["Medio Pago"] == "Efectivo"]["Valor"].sum()
        
    if pago_ahorro == "Efectivo": egresos_efectivo += ahorro_impuesto
    if pago_cadena == "Efectivo": egresos_efectivo += cadena_domingo
    if pago_otros == "Efectivo": egresos_efectivo += otros_pagos_pos

    egresos_efectivo_total = egresos_efectivo + propinas_retiradas
    total_egresos_generales = total_vales + total_proveedores + total_rodamientos + total_turnos + ahorro_impuesto + cadena_domingo + otros_pagos_pos + propinas_retiradas
    
    efectivo_teorico = base_efectivo + efectivo_pos - egresos_efectivo_total

    st.markdown("---")
    st.markdown("#### 3. CUADRE DE CAJA FÍSICA")
    st.info(f"💡 De los {formato_cop(total_egresos_generales)} en salidas totales, el sistema descontó **{formato_cop(egresos_efectivo_total)}** del cajón físico de efectivo.")
    
    efectivo_fisico = st.number_input("💰 ¿Efectivo físico contado en el cajón ahora mismo?", min_value=0, step=1000)
    diferencia_caja = efectivo_fisico - efectivo_teorico

    col_res1, col_res2, col_res3 = st.columns(3)
    col_res1.metric("Cajón Esperado (Teórico)", formato_cop(efectivo_teorico))
    col_res2.metric("Cajón Contado (Físico)", formato_cop(efectivo_fisico))
    col_res3.metric("Diferencia de Caja", formato_cop(diferencia_caja))

    if diferencia_caja == 0:
        st.success("✅ **CAJA CUADRADA EXACTAMENTE**")
    elif diferencia_caja > 0:
        st.warning(f"🟢 **SOBRANTE EN CAJÓN:** {formato_cop(diferencia_caja)}")
    else:
        st.error(f"🔴 **FALTANTE EN CAJÓN:** {formato_cop(abs(diferencia_caja))}")

    observaciones_pos = st.text_area("Observaciones del turno")

    def generar_pdf_en_memoria():
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer)
        
        pdf.drawString(50, 800, "CIERRE DE CAJA - RESTAURANTE")
        pdf.drawString(50, 780, f"Fecha: {fecha_cierre.strftime('%d/%m/%Y')} | Apertura: {apertura}")
        pdf.drawString(50, 760, f"Base Efectivo: {formato_cop(base_efectivo)}")
        pdf.drawString(50, 730, "--------------------------------------------------")
        pdf.drawString(50, 710, "INGRESOS DEL DIA")
        pdf.drawString(50, 690, f"Efectivo: {formato_cop(efectivo_pos)}")
        pdf.drawString(50, 670, f"Nequi: {formato_cop(nequi_pos)}")
        pdf.drawString(50, 650, f"Daviplata: {formato_cop(daviplata_pos)}")
        pdf.drawString(50, 630, f"Tarjeta: {formato_cop(tarjeta_pos)}")
        pdf.drawString(50, 610, f"Transferencia: {formato_cop(transferencia_pos)}")
        pdf.drawString(50, 580, f"TOTAL VENTAS: {formato_cop(total_pos)}")
        pdf.drawString(50, 550, "--------------------------------------------------")
        pdf.drawString(50, 530, "SALIDAS Y EGRESOS")
        pdf.drawString(50, 510, f"Retiro de Propinas: {formato_cop(propinas_retiradas)}")
        pdf.drawString(50, 490, f"Total Vales: {formato_cop(total_vales)}")
        pdf.drawString(50, 470, f"Total Proveedores: {formato_cop(total_proveedores)}")
        pdf.drawString(50, 450, f"Total Rodamientos: {formato_cop(total_rodamientos)}")
        pdf.drawString(50, 430, f"Total Turnos Extras: {formato_cop(total_turnos)}")
        pdf.drawString(50, 410, f"Impuestos/Cadena/Otros: {formato_cop(ahorro_impuesto + cadena_domingo + otros_pagos_pos)}")
        pdf.drawString(50, 380, "--------------------------------------------------")
        pdf.drawString(50, 360, "CUADRE DE CAJA FISICA (EFECTIVO)")
        pdf.drawString(50, 340, f"Gastado/Retirado en Efectivo: {formato_cop(egresos_efectivo_total)}")
        pdf.drawString(50, 320, f"Cajón Esperado: {formato_cop(efectivo_teorico)}")
        pdf.drawString(50, 300, f"Cajón Contado: {formato_cop(efectivo_fisico)}")
        pdf.drawString(50, 280, f"Diferencia: {formato_cop(diferencia_caja)}")
        
        pdf.save()
        buffer.seek(0)
        return buffer

    st.markdown("---")
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("💾 Guardar Todo el Turno", type="primary"):
            nuevo_cierre = pd.DataFrame({
                "Fecha": [fecha_cierre], "Responsable": [apertura], "Base": [base_efectivo],
                "Efectivo": [efectivo_pos], "Nequi": [nequi_pos], "Daviplata": [daviplata_pos],
                "Tarjeta": [tarjeta_pos], "Transferencia": [transferencia_pos],
                "Total Ventas": [total_pos], "Gastos Efectivo": [egresos_efectivo_total],
                "Total Egresos": [total_egresos_generales], "Efectivo Físico": [efectivo_fisico],
                "Diferencia": [diferencia_caja], "Observaciones": [observaciones_pos]
            })
            guardar_en_excel(nuevo_cierre, "data/cierre_pos.xlsx")

            if not vales_edit.empty:
                vales_validos = vales_edit[(vales_edit["Valor"] > 0) & (vales_edit["Empleado"].notna())]
                if not vales_validos.empty:
                    vales_a_guardar = pd.DataFrame({
                        "Fecha": [fecha_cierre] * len(vales_validos),
                        "Empleado": vales_validos["Empleado"],
                        "Concepto": vales_validos["Concepto"],
                        "Valor": vales_validos["Valor"],
                        "Medio Pago": vales_validos["Medio Pago"],
                        "Estado": ["Pagado"] * len(vales_validos)
                    })
                    guardar_en_excel(vales_a_guardar, "data/vales.xlsx")

            if not prov_edit.empty:
                prov_validos = prov_edit[(prov_edit["Valor"] > 0) & (prov_edit["Proveedor"].notna())]
                if not prov_validos.empty:
                    prov_a_guardar = pd.DataFrame({
                        "Fecha": [fecha_cierre] * len(prov_validos),
                        "Categoria": ["Proveedores"] * len(prov_validos),
                        "Descripcion": prov_validos["Proveedor"] + " - " + prov_validos["Detalle"],
                        "Valor": prov_validos["Valor"],
                        "Medio Pago": prov_validos["Medio Pago"]
                    })
                    guardar_en_excel(prov_a_guardar, "data/gastos.xlsx")

            gastos_adicionales = []
            if propinas_retiradas > 0:
                gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Operación", "Descripcion": "Retiro de Propinas", "Valor": propinas_retiradas, "Medio Pago": "Efectivo"})
            
            if not domiciliarios_edit.empty:
                dom_validos = domiciliarios_edit[(domiciliarios_edit["Valor"] > 0) & (domiciliarios_edit["Domiciliario"].notna())]
                for _, row in dom_validos.iterrows():
                    gastos_adicionales.append({
                        "Fecha": fecha_cierre, 
                        "Categoria": "Operación", 
                        "Descripcion": f"Rodamiento Domiciliario - {row['Domiciliario']}", 
                        "Valor": row["Valor"], 
                        "Medio Pago": row["Medio Pago"]
                    })

            if not turnos_edit.empty:
                turnos_validos = turnos_edit[(turnos_edit["Valor"] > 0) & (turnos_edit["Persona"].notna())]
                for _, row in turnos_validos.iterrows():
                    gastos_adicionales.append({
                        "Fecha": fecha_cierre, 
                        "Categoria": "Operación", 
                        "Descripcion": f"Turno Extra - {row['Persona']} ({row['Detalle / Turno']})", 
                        "Valor": row["Valor"], 
                        "Medio Pago": row["Medio Pago"]
                    })

            if ahorro_impuesto > 0:
                gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Ahorro Impuesto", "Descripcion": "Fijo Diario", "Valor": ahorro_impuesto, "Medio Pago": pago_ahorro})
            if cadena_domingo > 0:
                gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Cadena", "Descripcion": "Recogida Dominical", "Valor": cadena_domingo, "Medio Pago": pago_cadena})
            if otros_pagos_pos > 0:
                gastos_adicionales.append({"Fecha": fecha_cierre, "Categoria": "Otros", "Descripcion": "Imprevistos", "Valor": otros_pagos_pos, "Medio Pago": pago_otros})
            
            if gastos_adicionales:
                guardar_en_excel(pd.DataFrame(gastos_adicionales), "data/gastos.xlsx")

            st.success("✅ ¡Turno guardado exitosamente!")

    with col_btn2:
        st.download_button(
            label="🖨️ Descargar Tirilla (PDF)", data=generar_pdf_en_memoria(),
            file_name=f"Cierre_{fecha_cierre.strftime('%Y%m%d')}.pdf", mime="application/pdf"
        )


# ==================================================
# PROVEEDORES
# ==================================================

elif opcion == "Proveedores":
    st.markdown("### 🚚 Gestión de Proveedores y Estado de Cuenta")

    with st.form("form_proveedores", clear_on_submit=True):
        st.markdown("#### Registrar Nuevo Proveedor")
        nombre_prov = st.text_input("Nombre de la Empresa o Persona (Ej: Centro de Acopio)")
        tel_prov = st.text_input("Teléfono de Contacto")
        cuenta_pagos = st.text_input("Cuenta para Pagos (Ej: Nequi 300123...)")
        notas_pedidos = st.text_area("Notas / Días de Pedido (Ej: Piden los martes)")
        
        guardar_prov = st.form_submit_button("💾 Guardar Proveedor")

        if guardar_prov and nombre_prov:
            nuevo_proveedor = pd.DataFrame({
                "Nombre": [nombre_prov], 
                "Teléfono": [tel_prov], 
                "Cuenta Pagos": [cuenta_pagos],
                "Notas / Pedidos": [notas_pedidos]
            })
            guardar_en_excel(nuevo_proveedor, "data/proveedores.xlsx")
            st.success(f"✅ Proveedor '{nombre_prov}' creado correctamente.")

    if os.path.exists("data/proveedores.xlsx"):
        st.markdown("---")
        st.markdown("#### 📋 Directorio de Proveedores")
        df_provs = pd.read_excel("data/proveedores.xlsx")
        st.dataframe(df_provs, use_container_width=True)

        st.markdown("---")
        st.markdown("#### 📊 Estado de Cuenta Histórico")
        prov_seleccionado = st.selectbox("Seleccione Proveedor a Consultar", df_provs["Nombre"].tolist())
        
        if os.path.exists("data/gastos.xlsx"):
            df_gastos = pd.read_excel("data/gastos.xlsx")
            prov_pagos = df_gastos[df_gastos["Categoria"] == "Proveedores"].copy()
            if not prov_pagos.empty:
                prov_pagos_filtrados = prov_pagos[prov_pagos["Descripcion"].str.contains(prov_seleccionado, na=False, case=False)]
                if not prov_pagos_filtrados.empty:
                    total_pagado_prov = prov_pagos_filtrados["Valor"].sum()
                    st.metric(f"Total Pagado a {prov_seleccionado}", formato_cop(total_pagado_prov))
                    st.markdown("**Historial de Abonos:**")
                    st.dataframe(prov_pagos_filtrados, hide_index=True, use_container_width=True)
                    
                    csv_prov = convertir_df_a_csv(prov_pagos_filtrados)
                    st.download_button(
                        label=f"📥 Exportar Estado de Cuenta ({prov_seleccionado})",
                        data=csv_prov,
                        file_name=f"Estado_Cuenta_{prov_seleccionado.replace(' ', '_')}.csv",
                        mime="text/csv"
                    )
                else:
                    st.info(f"No hay pagos registrados aún para {prov_seleccionado}.")
            else:
                st.info("No hay registros de pagos a proveedores.")


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
        
        guardar_empleado = st.form_submit_button("💾 Guardar Empleado")

        if guardar_empleado and nombre:
            nuevo_empleado = pd.DataFrame({
                "Nombre": [nombre], "Documento": [documento], "Cargo": [cargo]
            })
            guardar_en_excel(nuevo_empleado, "data/empleados.xlsx")
            st.success("✅ Empleado guardado correctamente")

    if os.path.exists("data/empleados.xlsx"):
        st.markdown("---")
        st.markdown("#### 📋 Personal Activo")
        st.dataframe(pd.read_excel("data/empleados.xlsx"), use_container_width=True)

    if os.path.exists("data/vales.xlsx"):
        st.markdown("---")
        st.markdown("#### 📊 Reporte y Estado de Cuenta de Vales")
        df_vales_reporte = pd.read_excel("data/vales.xlsx")
        df_vales_reporte["Fecha"] = pd.to_datetime(df_vales_reporte["Fecha"]).dt.date
        
        col_f1, col_f2 = st.columns(2)
        fecha_inicio = col_f1.date_input("Fecha Inicio", value=df_vales_reporte["Fecha"].min())
        fecha_fin = col_f2.date_input("Fecha Fin", value=datetime.today())
        
        mask = (df_vales_reporte["Fecha"] >= fecha_inicio) & (df_vales_reporte["Fecha"] <= fecha_fin)
        vales_filtrados = df_vales_reporte.loc[mask]
        
        if not vales_filtrados.empty:
            st.markdown("##### Resumen Acumulado por Empleado")
            resumen_vales = vales_filtrados.groupby("Empleado")["Valor"].sum().reset_index()
            
            resumen_vales_display = resumen_vales.copy()
            resumen_vales_display["Valor Total"] = resumen_vales_display["Valor"].apply(formato_cop)
            resumen_vales_display = resumen_vales_display.drop(columns=["Valor"])
            st.dataframe(resumen_vales_display, hide_index=True, use_container_width=True)
            
            st.download_button(
                label="📥 Exportar Resumen de Vales por Empleado (CSV)",
                data=convertir_df_a_csv(resumen_vales),
                file_name=f"Vales_Por_Empleado_{fecha_inicio}_a_{fecha_fin}.csv",
                mime="text/csv"
            )

            st.markdown("##### Detalle Completo de Vales")
            st.dataframe(vales_filtrados, hide_index=True, use_container_width=True)
            
            st.download_button(
                label="📥 Exportar Detalle Total de Vales por Fechas (CSV)",
                data=convertir_df_a_csv(vales_filtrados),
                file_name=f"Detalle_Vales_{fecha_inicio}_a_{fecha_fin}.csv",
                mime="text/csv"
            )
        else:
            st.info("No hay vales registrados en este rango de fechas.")


# ==================================================
# MÓDULOS DE CONSULTA
# ==================================================

elif opcion == "Consulta de Gastos":
    st.markdown("### 📋 Consulta General de Gastos")
    if os.path.exists("data/gastos.xlsx"):
        df_g = pd.read_excel("data/gastos.xlsx")
        df_g["Fecha"] = pd.to_datetime(df_g["Fecha"]).dt.date
        
        col_f1, col_f2 = st.columns(2)
        f_ini = col_f1.date_input("Fecha Inicio", value=df_g["Fecha"].min())
        f_fin = col_f2.date_input("Fecha Fin", value=datetime.today())
        
        mask = (df_g["Fecha"] >= f_ini) & (df_g["Fecha"] <= f_fin)
        df_g_filtrado = df_g.loc[mask]
        
        st.metric("Total Gastos en Rango", formato_cop(df_g_filtrado["Valor"].sum() if not df_g_filtrado.empty else 0))
        st.dataframe(df_g_filtrado, use_container_width=True, hide_index=True)
        
        if not df_g_filtrado.empty:
            st.download_button("📥 Exportar Gastos Filtrados", convertir_df_a_csv(df_g_filtrado), "Gastos.csv", "text/csv")
    else:
        st.info("Aún no hay gastos registrados.")

elif opcion == "Consulta de Vales":
    st.markdown("### 📋 Historial Detallado de Vales")
    if os.path.exists("data/vales.xlsx"):
        df_v = pd.read_excel("data/vales.xlsx")
        df_v["Fecha"] = pd.to_datetime(df_v["Fecha"]).dt.date
        
        col_f1, col_f2 = st.columns(2)
        f_ini = col_f1.date_input("Fecha Inicio", value=df_v["Fecha"].min())
        f_fin = col_f2.date_input("Fecha Fin", value=datetime.today())
        
        mask = (df_v["Fecha"] >= f_ini) & (df_v["Fecha"] <= f_fin)
        df_v_filtrado = df_v.loc[mask]
        
        st.metric("Total Vales en Rango", formato_cop(df_v_filtrado["Valor"].sum() if not df_v_filtrado.empty else 0))
        st.dataframe(df_v_filtrado, use_container_width=True, hide_index=True)
        
        if not df_v_filtrado.empty:
            st.download_button("📥 Exportar Vales Filtrados", convertir_df_a_csv(df_v_filtrado), "Vales.csv", "text/csv")
    else:
        st.info("Aún no hay vales registrados.")

elif opcion == "Consulta de Retiros":
    st.markdown("### 📋 Consulta de Retiros")
    if os.path.exists("data/retiros.xlsx"):
        df_r = pd.read_excel("data/retiros.xlsx")
        if "Fecha" in df_r.columns:
            df_r["Fecha"] = pd.to_datetime(df_r["Fecha"]).dt.date
            col_f1, col_f2 = st.columns(2)
            f_ini = col_f1.date_input("Fecha Inicio", value=df_r["Fecha"].min())
            f_fin = col_f2.date_input("Fecha Fin", value=datetime.today())
            mask = (df_r["Fecha"] >= f_ini) & (df_r["Fecha"] <= f_fin)
            df_r = df_r.loc[mask]
        st.dataframe(df_r, use_container_width=True, hide_index=True)
        if not df_r.empty:
            st.download_button("📥 Exportar Retiros", convertir_df_a_csv(df_r), "Retiros.csv", "text/csv")
    else:
        st.info("Aún no hay retiros registrados.")

elif opcion == "Consulta Otros Ingresos":
    st.markdown("### 📋 Consulta de Otros Ingresos")
    if os.path.exists("data/ingresos.xlsx"):
        df_i = pd.read_excel("data/ingresos.xlsx")
        if "Fecha" in df_i.columns:
            df_i["Fecha"] = pd.to_datetime(df_i["Fecha"]).dt.date
            col_f1, col_f2 = st.columns(2)
            f_ini = col_f1.date_input("Fecha Inicio", value=df_i["Fecha"].min())
            f_fin = col_f2.date_input("Fecha Fin", value=datetime.today())
            mask = (df_i["Fecha"] >= f_ini) & (df_i["Fecha"] <= f_fin)
            df_i = df_i.loc[mask]
        st.dataframe(df_i, use_container_width=True, hide_index=True)
        if not df_i.empty:
            st.download_button("📥 Exportar Ingresos", convertir_df_a_csv(df_i), "Ingresos.csv", "text/csv")
    else:
        st.info("Aún no hay ingresos extra registrados.")


# ==================================================
# RESPALDO Y SEGURIDAD
# ==================================================

elif opcion == "Respaldo y Seguridad":
    st.markdown("### 🔒 Respaldo de Datos (Backup)")
    st.markdown("Descarga un archivo comprimido `.zip` con todos los archivos de Excel de tu base de datos.")

    def crear_zip_datos():
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for carpeta, _, archivos in os.walk("data"):
                for archivo in archivos:
                    ruta_completa = os.path.join(carpeta, archivo)
                    zip_file.write(ruta_completa, arcname=archivo)
        buffer.seek(0)
        return buffer

    if os.path.exists("data"):
        st.download_button(
            label="📦 Descargar Respaldo Completo (ZIP)",
            data=crear_zip_datos(),
            file_name=f"Backup_Restaurante_{datetime.today().strftime('%Y%m%d')}.zip",
            mime="application/zip"
        )
    else:
        st.warning("No se encontró la carpeta 'data' con archivos para respaldar.")