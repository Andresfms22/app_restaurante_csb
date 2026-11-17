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
    else: st.warning("No hay datos para respaldar.")

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
    else: st.warning("No hay datos para respaldar.")

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
