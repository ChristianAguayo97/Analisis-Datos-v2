import streamlit as st
import pandas as pd
import gspread
from ydata_profiling import ProfileReport 
from streamlit.components.v1 import html
from google.oauth2.service_account import Credentials

st.set_page_config(layout="wide", page_title="Analizador de Hojas Google", page_icon="📊")
st.title("📊 Analizador de Hojas Google")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly"
]


@st.cache_resource
def get_google_client():
    try:
        if 'gcp_service_account' not in st.secrets:
            st.error("Error: No se encontraron credenciales en secrets.toml")
            st.stop()
            
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=SCOPES
        )
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Error de autenticación: {e}")
        st.stop()


@st.cache_data(show_spinner="Buscando tus hojas de cálculo...")
def get_all_spreadsheets():
    try:
        gc = get_google_client()
        spreadsheets = []
        for spreadsheet in gc.openall():
            spreadsheets.append({
                "title": spreadsheet.title,
                "id": spreadsheet.id,
                "url": f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
            })
        return spreadsheets
    except Exception as e:
        st.error(f"Error obteniendo hojas: {str(e)}")
        return []

@st.cache_data(show_spinner="Cargando estructura del documento...")
def get_worksheets(spreadsheet_id):
    try:
        gc = get_google_client()
        spreadsheet = gc.open_by_key(spreadsheet_id)
        worksheets = []
        for worksheet in spreadsheet.worksheets():
            worksheets.append({
                "title": worksheet.title,
                "id": worksheet.id,
                "row_count": worksheet.row_count,
                "col_count": worksheet.col_count
            })
        return worksheets
    except Exception as e:
        st.error(f"Error obteniendo hojas: {str(e)}")
        return []

@st.cache_data(show_spinner="Cargando datos...")
def cargar_datos(spreadsheet_id, worksheet_id):
    try:
        gc = get_google_client()
        spreadsheet = gc.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.get_worksheet_by_id(worksheet_id)
        data = worksheet.get_all_values()
        
        if not data:
            return pd.DataFrame()
        df = pd.DataFrame(data)
        header_row = encontrar_mejor_header(df)
        
        if header_row is not None:
            new_header = df.iloc[header_row]
            df = df[header_row+1:]
            df.columns = new_header
            df.columns = df.columns.astype(str)  
            df.reset_index(drop=True, inplace=True)

        
        return df
    except Exception as e:
        st.error(f"Error cargando datos: {str(e)}")
        return pd.DataFrame()


def encontrar_mejor_header(df, max_filas=5):
    mejor_puntaje = 0
    mejor_fila = None
    
    for i in range(min(max_filas, len(df))):
        fila = df.iloc[i]
        non_empty = sum(1 for cell in fila if str(cell).strip() != "")
        avg_len = sum(len(str(cell)) for cell in fila) / len(fila)
        puntaje = non_empty * avg_len
        
        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_fila = i
    
    return mejor_fila

@st.cache_data(show_spinner="Generando reporte...")
def generar_reporte(df, nombre_doc, nombre_hoja):
    if df.empty:
        return None
    columnas_validas = [
        col for col in df.columns
        if not df[col].apply(lambda x: isinstance(x, (list, dict, tuple, set))).any()
    ]
    df_filtrado = df[columnas_validas].copy()

    # Opcional: podés mostrar las columnas excluidas
    columnas_excluidas = set(df.columns) - set(columnas_validas)
    if columnas_excluidas:
        st.warning(f"Se excluyeron columnas no compatibles: {', '.join(columnas_excluidas)}")
    perfil = ProfileReport(
        df_filtrado,
        title=f"Reporte: {nombre_doc} - {nombre_hoja}",
        explorative=True,
        html={
            "style": {"full_width": True},
            "navbar_show": False
        }
    )
    return perfil.to_html()



st.sidebar.header("Configuración de Google Sheets")
gc = get_google_client()


spreadsheets = get_all_spreadsheets()

if not spreadsheets:
    st.warning("No se encontraron hojas de cálculo. Comparte tus hojas con la cuenta de servicio.")
    st.stop()

spreadsheet_titles = [f"{s['title']} ({s['id']})" for s in spreadsheets]
selected_title = st.sidebar.selectbox(
    "Selecciona un documento:", 
    options=spreadsheet_titles,
    index=0
)

selected_spreadsheet = next((s for s in spreadsheets if f"{s['title']} ({s['id']})" == selected_title), None)

if selected_spreadsheet:
    worksheets = get_worksheets(selected_spreadsheet["id"])
    
    if not worksheets:
        st.warning("Este documento no tiene hojas visibles.")
        st.stop()
    
    worksheet_names = [f"{ws['title']} ({ws['row_count']}x{ws['col_count']})" for ws in worksheets]
    selected_ws_name = st.sidebar.selectbox(
        "Selecciona una hoja:", 
        options=worksheet_names,
        index=0
    )
    

    selected_worksheet = next((ws for ws in worksheets if f"{ws['title']} ({ws['row_count']}x{ws['col_count']})" == selected_ws_name), None)
    

    if selected_worksheet and st.sidebar.button("Generar Reporte", type="primary"):
        with st.spinner(f"Cargando {selected_worksheet['title']}..."):
            df = cargar_datos(selected_spreadsheet["id"], selected_worksheet["id"])
            
        if df.empty:
            st.error("No se pudieron cargar datos de esta hoja")
        else:
            st.success(f"Datos cargados: {df.shape[0]} filas, {df.shape[1]} columnas")
            
            with st.spinner("Creando reporte de calidad..."):
                html_report = generar_reporte(
                    df, 
                    selected_spreadsheet["title"], 
                    selected_worksheet["title"]
                )
            
            if html_report:
                st.subheader(f"Reporte de Calidad: {selected_spreadsheet['title']} - {selected_worksheet['title']}")
                html(html_report, height=1000, scrolling=True)
            else:
                st.error("Error generando el reporte")
else:
    st.warning("Selecciona un documento válido")