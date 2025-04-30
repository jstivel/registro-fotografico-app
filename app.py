import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.utils.units import cm_to_EMU
from io import BytesIO
from PIL import Image as PILImage
from openpyxl.drawing.xdr import XDRPositiveSize2D

# --- Configuración de las dimensiones del área de la imagen (en cm) ---
AREA_HEIGHT_CM = 6.8
AREA_WIDTH_CM = 9.42

# --- Función para redimensionar la imagen manteniendo la relación de aspecto ---
def redimensionar_imagen(imagen_pil, max_ancho_cm, max_alto_cm, dpi=96):
    """Redimensiona una imagen de Pillow manteniendo su relación de aspecto."""
    max_ancho_pixels = max_ancho_cm * dpi / 2.54
    max_alto_pixels = max_alto_cm * dpi / 2.54
    ancho, alto = imagen_pil.size

    ratio_ancho = max_ancho_pixels / ancho
    ratio_alto = max_alto_pixels / alto

    if ratio_ancho < 1 or ratio_alto < 1:  # Solo redimensionar si es más grande
        ratio = min(ratio_ancho, ratio_alto)
        nuevo_ancho = int(ancho * ratio)
        nuevo_alto = int(alto * ratio)
        imagen_redimensionada = imagen_pil.resize((nuevo_ancho, nuevo_alto))
        return imagen_redimensionada
    return imagen_pil

# --- Función para convertir cm a EMUs ---
c2e = cm_to_EMU

# --- Función para calcular el offset en EMUs para centrar la imagen ---
def calcular_offset(area_cm, img_cm):
    delta_cm = ((area_cm - img_cm) / 2)+0.1
    return c2e(delta_cm)

# --- Interfaz de Usuario en Streamlit ---
st.title("Registro fotografico")

formato = ["Preventivo", "Recorredor", "clientes interno", "clientes externo"]
formato_seleccionado = st.radio("Selecciona el formato:", formato)

opciones = ["BRAYAN STIVEN SALAMANCA CASTAÑEDA", "JUAN CARLOS RODRIGUEZ CHIQUILLO", "YOVANNI GOMEZ PEÑA", "MICHAEL ESTEBAN URQUIJO LOPEZ", 
            "FABIAN DAVID CIFUENTES GRUESO","JOHN EZEQUIEL TANGARIFE ARENAS"]
map_telefono = {
            "BRAYAN STIVEN SALAMANCA CASTAÑEDA":"3133977853", 
            "JUAN CARLOS RODRIGUEZ CHIQUILLO":"3214522373", 
            "YOVANNI GOMEZ PEÑA":"3112973928", 
            "MICHAEL ESTEBAN URQUIJO LOPEZ":"3114669376", 
            "FABIAN DAVID CIFUENTES GRUESO":"3112042566",
            "JOHN EZEQUIEL TANGARIFE ARENAS":"3118859551"}

ejecutor = st.selectbox("Ejecutor:", opciones)

if formato_seleccionado == "clientes interno" or formato_seleccionado == "clientes externo":
    cliente = st.text_input("Nombre del sitio:")

telefono = map_telefono.get(ejecutor, "")
direccion = st.text_input("DIRECCIÓN:")
fecha_visita = st.date_input("FECHA DE LA VISITA:")
telefono = st.text_input("TELEFONO:", value=telefono, disabled=True)

uploaded_files = st.file_uploader("Subir Registros Fotográficos", accept_multiple_files=True, type=["png", "jpg", "jpeg"])
descripciones = [""] * len(uploaded_files)

# --- Sección de Vista Previa con Miniaturas y Campos de Descripción Adyacentes ---
if ejecutor or direccion or fecha_visita or telefono or uploaded_files:
    st.subheader("Vista Previa de los Datos:")
    if ejecutor:
        st.write(f"**Ejecutor:** {ejecutor}")
    if direccion:
        st.write(f"**Dirección:** {direccion}")
    if fecha_visita:
        st.write(f"**Fecha de la Visita:** {fecha_visita.strftime('%Y-%m-%d')}")
    if telefono:
        st.write(f"**Teléfono:** {telefono}")
    if uploaded_files:
        st.write("**Registros Fotográficos:**")
        for i, file in enumerate(uploaded_files):
            col_preview, col_descripcion = st.columns([1, 2])
            with col_preview:
                st.image(file, caption=f"Foto {i+1}", width=100)
            with col_descripcion:
                descripciones[i] = st.text_input(f"Descripción para la Foto {i+1}:", key=f"descripcion_{i}")

if st.button("Generar Excel"):
    if ejecutor and direccion and fecha_visita and telefono and uploaded_files:
        try:
            # --- Cargar el archivo Excel existente ---
            ruta_excel = ""
            fila_foto_inicio = 8
            columna_foto_inicio = 1  # Columna A
            celda_ejecutor = 'G7'
            celda_dirección = 'C5'
            celda_fecha = 'C6'
            celda_tel = 'H7'

            if formato_seleccionado == "Preventivo" or formato_seleccionado == "Recorredor":
                ruta_excel = 'RF_PREVENTIVO.XLSX'  
                         

            elif formato_seleccionado == "clientes interno":
                ruta_excel = 'RF_CLIENTE_INTERNO.xlsx'
                fila_foto_inicio = 10     # Fila 10
                celda_ejecutor = 'G8'
                celda_dirección = 'C6'
                celda_fecha = 'C7'
                celda_tel = 'H8'

            elif formato_seleccionado == "clientes externo":
                ruta_excel = 'RF_CLIENTE_EXTERNO.xlsx'
                fila_foto_inicio = 10     # Fila 10 
                celda_ejecutor = 'G8'
                celda_dirección = 'C6'
                celda_fecha = 'C7'
                celda_tel = 'H8'                        
            
            libro = load_workbook(ruta_excel)
            hoja = libro.active

            # --- Llenar los campos de texto ---
            if formato_seleccionado == "Recorredor":
                hoja['A4'] = 'REGISTRO FOTOGRÁFICO RECORREDOR'
                hoja['D7'] = 'RECORREDOR'
            
            hoja[celda_ejecutor] = ejecutor
            hoja[celda_dirección] = direccion
            hoja[celda_fecha] = fecha_visita.strftime("%d-%m-%Y")
            hoja[celda_tel] = telefono            
            if formato_seleccionado != "Preventivo":
                hoja['C5'] = cliente  
            

            for i, archivo_subido in enumerate(uploaded_files):
                # --- Redimensionar la imagen con Pillow ---
                img_pil = PILImage.open(archivo_subido)
                img_redimensionada = redimensionar_imagen(img_pil, AREA_WIDTH_CM, AREA_HEIGHT_CM)
                img_width_cm, img_height_cm = img_redimensionada.size[0] * 2.54 / 96, img_redimensionada.size[1] * 2.54 / 96

                # --- Convertir la imagen redimensionada para openpyxl ---
                img_buffer = BytesIO()
                img_redimensionada.save(img_buffer, format=img_pil.format)
                img_buffer.seek(0)
                img = Image(img_buffer)

                # --- Calcular la celda de anclaje y el offset para centrar ---
                if (i + 1) % 2 != 0:  # Foto en columna A
                    col_idx = columna_foto_inicio - 1 # openpyxl usa índice base 0
                    row_idx = fila_foto_inicio - 1
                else:  # Foto en columna E
                    col_idx = columna_foto_inicio + 4 - 1 # openpyxl usa índice base 0
                    row_idx = fila_foto_inicio - 1

                x_offset_emu = calcular_offset(AREA_WIDTH_CM, img_width_cm)
                y_offset_emu = calcular_offset(AREA_HEIGHT_CM, img_height_cm)

                # --- Definir el marcador de anclaje ---
                marker = AnchorMarker(
                    col=col_idx,
                    colOff=x_offset_emu,
                    row=row_idx,
                    rowOff=y_offset_emu
                )

                # --- Definir el tamaño de la imagen (en EMUs) como un objeto XDRPositiveSize2D ---
                
                size = XDRPositiveSize2D(cx=c2e(img_width_cm), cy=c2e(img_height_cm))

                # --- Crear el anclaje de una celda ---
                img.anchor = OneCellAnchor(_from=marker, ext=size)

                # --- Añadir la imagen a la hoja ---
                hoja.add_image(img)

                # --- Calcular la celda unificada para la descripción ---
                if (i + 1) % 2 != 0:
                    celda_descripcion_inicio = f"B{fila_foto_inicio + 15}"
                    celda_descripcion_fin = f"D{fila_foto_inicio + 16}"
                else:
                    celda_descripcion_inicio = f"F{fila_foto_inicio + 15}"
                    celda_descripcion_fin = f"H{fila_foto_inicio + 16}"
                    fila_foto_inicio += 17 # Saltar a la siguiente fila para la siguiente pareja

                hoja.merge_cells(f"{celda_descripcion_inicio}:{celda_descripcion_fin}")
                hoja[celda_descripcion_inicio] = descripciones[i]

            # --- Guardar el archivo modificado en memoria ---
            buffer = BytesIO()
            libro.save(buffer)
            buffer.seek(0)

            # --- Ofrecer la descarga ---
            st.download_button(
                label="Descargar Excel Generado",
                data=buffer,
                file_name=f"Registro_fotografico_{fecha_visita.strftime('%d-%m-%Y')} {direccion}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.success("¡El archivo Excel ha sido generado exitosamente!")

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
    else:
        st.warning("Por favor, completa todos los campos y sube al menos una foto.")