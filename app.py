import pandas as pd
import streamlit as st
import io

st.set_page_config(page_title="Visor de Excel", layout="wide")
st.title("Aplicación Web - Visor de Excel")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx o .xlsm)", type=["xlsx", "xlsm"])

# Estado de la vista
if "mostrar_visor" not in st.session_state:
    st.session_state.mostrar_visor = False

if "lineas_marcadas" not in st.session_state:
    st.session_state.lineas_marcadas = set()

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")

        if 'Hoja1' in excel_file.sheet_names:
            hoja1_df = excel_file.parse('Hoja1')

            if not st.session_state.mostrar_visor:
                st.subheader("Vista previa de Hoja1")

                # Crear una copia para agregar columna de marcado
                hoja1_df_display = hoja1_df.copy()
                hoja1_df_display.insert(0, "ESC.", ["✅" if i in st.session_state.lineas_marcadas else "" for i in hoja1_df.index])

                # Mostrar tabla con sombreado
                def highlight_marked(row):
                    return ["background-color: #d4edda" if row.name in st.session_state.lineas_marcadas else "" for _ in row]

                st.dataframe(hoja1_df_display.style.apply(highlight_marked, axis=1), use_container_width=True)

                # Casillas para marcar tareas realizadas
                st.subheader("Marcar tareas realizadas")
                for i in hoja1_df.index:
                    marcado = i in st.session_state.lineas_marcadas
                    if st.checkbox(f"Fila {i+1}", key=f"chk_hoja1_{i}", value=marcado):
                        st.session_state.lineas_marcadas.add(i)
                    else:
                        st.session_state.lineas_marcadas.discard(i)

                if st.button("Ampliar - Mostrar Visor"):
                    st.session_state.mostrar_visor = True

            else:
                if 'Visor' in excel_file.sheet_names:
                    visor_df = excel_file.parse('Visor', header=None)
                    st.subheader("Contenido del Visor")

                    marcados = []
                    for i in range(len(visor_df)):
                        item = visor_df.iloc[i, 0]
                        marcado = i in st.session_state.lineas_marcadas

                        container_style = (
                            "background-color:#d4edda; border:2px solid #28a745;"
                            if marcado else "background-color:#f9f9f9; border:1px solid #ccc;"
                        )

                        with st.container():
                            cols = st.columns([0.05, 0.95])
                            with cols[0]:
                                if st.checkbox("", key=f"chk_{i}", value=marcado):
                                    st.session_state.lineas_marcadas.add(i)
                                else:
                                    st.session_state.lineas_marcadas.discard(i)
                            with cols[1]:
                                st.markdown(f"""
                                    <div style='{container_style} border-radius:10px; padding:10px; margin:5px;'>
                                        <strong style='font-size:18px;'>{item}</strong>
                                    </div>
                                """, unsafe_allow_html=True)
                        marcados.append("Sí" if i in st.session_state.lineas_marcadas else "No")

                    # Exportar Excel con marcados
                    export_df = visor_df.copy()
                    export_df["Realizado"] = marcados

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        export_df.to_excel(writer, sheet_name='Visor Marcado', index=False, header=False)

                    st.download_button(
                        label="Exportar marcados a Excel",
                        data=output.getvalue(),
                        file_name="visor_marcado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.button("Volver", on_click=lambda: st.session_state.update({"mostrar_visor": False}))
                else:
                    st.warning("La hoja 'Visor' no está presente en el archivo.")

        else:
            st.error("La hoja 'Hoja1' no se encuentra en el archivo.")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
