import pandas as pd
import os
import matplotlib.pyplot as plt
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import scrolledtext

# Configurar Matplotlib para trabajar en segundo plano (Evita conflictos con la GUI de Tkinter)
import matplotlib
matplotlib.use('Agg')

# =========================================================================
# LÓGICA DE PROCESAMIENTO CRÍTICO (Motor de análisis dinámico)
# =========================================================================
def procesar_asistencia(ruta_archivo, txt_output, horas_config):
    try:
        txt_output.insert(tk.END, "1. Leyendo archivo de asistencia...\n")
        txt_output.see(tk.END)
        txt_output.update_idletasks()
        
        # Cargar el archivo .xls (Formato antiguo de Excel usado por los biométricos)
        df = pd.read_excel(ruta_archivo, engine='xlrd')
        txt_output.insert(tk.END, f"   -> ¡Archivo cargado! Total registros encontrados: {len(df)}\n")
        
        txt_output.insert(tk.END, "2. Limpiando formatos de hora (a. m. / p. m.)...\n")
        df['Tiempo'] = df['Tiempo'].astype(str).str.replace('a. m.', 'AM', regex=False)
        df['Tiempo'] = df['Tiempo'].astype(str).str.replace('p. m.', 'PM', regex=False)
        
        txt_output.insert(tk.END, "3. Convirtiendo texto a fechas reales de Python...\n")
        df['Tiempo'] = pd.to_datetime(df['Tiempo'], errors='coerce')
        df = df.dropna(subset=['Tiempo'])
        
        df['Fecha'] = df['Tiempo'].dt.date
        df['Hora_Marcacion'] = df['Tiempo'].dt.time
        
        txt_output.insert(tk.END, "4. Aplicando los horarios configurados para el análisis diario...\n")
        txt_output.see(tk.END)
        txt_output.update_idletasks()
        
        # Helper para convertir los textos de la interfaz "HH:MM:SS" a minutos totales del día
        def texto_a_minutos(texto_hora):
            t = pd.to_datetime(texto_hora.strip()).time()
            return t.hour * 60 + t.minute

        # Cargar las horas dinámicas modificadas por el usuario en la GUI
        MINS_ENTRADA_OFICIAL = texto_a_minutos(horas_config['entrada'].get())
        MINS_INICIO_ALMUERZO = texto_a_minutos(horas_config['inicio_almuerzo'].get())
        MINS_FIN_ALMUERZO    = texto_a_minutos(horas_config['fin_almuerzo'].get())
        MINS_SALIDA_OFICIAL  = texto_a_minutos(horas_config['salida'].get())
        
        # Punto de corte intermedio (Medio día relativo)
        CORTE_MEDIO_DIA = 12 * 60 

        registro_infracciones = []

        # Agrupar datos por empleado y día para evaluar su comportamiento
        for (empleado, fecha), grupo in df.groupby(['Nombre', 'Fecha']):
            # Convertir todas las marcas del empleado en ese día a minutos del día y ordenarlas
            marcas_minutos = sorted([h.hour * 60 + h.minute for h in grupo['Hora_Marcacion'].tolist()])
            
            es_tarde = False
            es_temprano = False
            
            # --- Evaluar Entrada Mañana ---
            marcas_manana = [m for m in marcas_minutos if m < 600] # Antes de las 10:00 AM
            if marcas_manana:
                primera_entrada = marcas_manana[0]
                if primera_entrada > MINS_ENTRADA_OFICIAL:
                    es_tarde = True
            elif marcas_minutos:
                if marcas_minutos[0] > MINS_ENTRADA_OFICIAL:
                    es_tarde = True

            # --- Evaluar Salida Temprana Tarde ---
            marcas_tarde = [m for m in marcas_minutos if m >= CORTE_MEDIO_DIA]
            if marcas_tarde:
                ultima_marca_tarde = marcas_tarde[-1]
                
                # Si su última marca es antes de las 3:00 PM (15:00)
                if ultima_marca_tarde < 15 * 60:
                    if ultima_marca_tarde < MINS_INICIO_ALMUERZO:
                        es_temprano = True
                    elif ultima_marca_tarde < MINS_SALIDA_OFICIAL and len(marcas_tarde) == 1:
                        es_temprano = True
                else:
                    if ultima_marca_tarde < MINS_SALIDA_OFICIAL:
                        es_temprano = True

            registro_infracciones.append({
                'Nombre': empleado,
                'Fecha': fecha,
                'Llegada_Tarde': 1 if es_tarde else 0,
                'Salida_Temprana': 1 if es_temprano else 0
            })

        df_infracciones_diarias = pd.DataFrame(registro_infracciones)

        txt_output.insert(tk.END, "5. Consolidando totales acumulados y ordenando por ID...\n")
        reporte = df_infracciones_diarias.groupby('Nombre').agg(
            Llegadas_Tarde=('Llegada_Tarde', 'sum'),
            Salidas_Tempranas=('Salida_Temprana', 'sum')
        ).reset_index()

        # Extraer ID alfanumérico para ordenamiento numérico estricto
        def extraer_numero(nombre):
            numeros = re.findall(r'\d+', str(nombre))
            return int(numeros[0]) if numeros else 999

        reporte['ID_Numerico'] = reporte['Nombre'].apply(extraer_numero)
        reporte = reporte.sort_values(by=['ID_Numerico', 'Nombre'], ascending=[True, True])
        reporte = reporte.drop(columns=['ID_Numerico'])

        # Renombrar columnas dinámicamente en el reporte visual según las horas de la GUI
        ent_label = horas_config['entrada'].get().strip()[:5]
        sal_label = horas_config['salida'].get().strip()[:5]
        col_tarde = f'Llegadas Tarde (> {ent_label})'
        col_temprano = f'Salidas Tempranas (< {sal_label})'
        
        reporte.columns = ['Nombre del Empleado', col_tarde, col_temprano]

        # Definir rutas de guardado automáticas basadas en la ubicación del archivo original
        carpeta_origen = os.path.dirname(ruta_archivo)
        ruta_excel_salida = os.path.join(carpeta_origen, 'Reporte_Infracciones_Asistencia.xlsx')
        ruta_img_salida = os.path.join(carpeta_origen, 'Reporte_Asistencia.png')

        # Guardar en Excel
        reporte.to_excel(ruta_excel_salida, index=False)
        
        # Generar Tabla Estilizada Gráfica en PNG
        alto_imagen = max(6, len(reporte) * 0.35)
        fig, ax = plt.subplots(figsize=(12, alto_imagen))
        ax.axis('off')

        tabla_visual = ax.table(
            cellText=reporte.values, 
            colLabels=reporte.columns, 
            cellLoc='center', 
            loc='center'
        )
        tabla_visual.auto_set_font_size(False)
        tabla_visual.set_fontsize(10)
        tabla_visual.scale(1.2, 1.3)

        # Pintar cabecera y filas alternas
        for (row, col), cell in tabla_visual.get_celld().items():
            if row == 0:
                cell.set_text_props(weight='bold', color='white')
                cell.set_facecolor('#4F5B66')
            elif row % 2 == 0:
                cell.set_facecolor('#F9F9F9')

        plt.savefig(ruta_img_salida, bbox_inches='tight', dpi=300)
        plt.close()

        # Imprimir resumen de texto en la consola integrada de la app
        txt_output.insert(tk.END, "\n================ REPORTES DE ASISTENCIA ================\n")
        txt_output.insert(tk.END, reporte.to_string(index=False))
        txt_output.insert(tk.END, "\n========================================================\n")
        txt_output.insert(tk.END, f"\n¡Éxito! Reporte generado usando las horas establecidas.\n")
        txt_output.insert(tk.END, f"-> Archivo Excel: {ruta_excel_salida}\n")
        txt_output.insert(tk.END, f"-> Imagen PNG: {ruta_img_salida}\n")
        txt_output.see(tk.END)
        
        messagebox.showinfo("¡Completado!", "El reporte dinámico se procesó y guardó exitosamente.")

    except PermissionError:
        # Captura específica de si el archivo de salida está abierto en Microsoft Excel
        messagebox.showerror(
            "Archivo Abierto o Bloqueado", 
            "El archivo 'Reporte_Infracciones_Asistencia.xlsx' está abierto en Excel.\n\nPor favor, cierra la ventana de Excel y vuelve a presionar el botón."
        )
    except Exception as e:
        messagebox.showerror(
            "Error de Formato", 
            f"Asegúrate de escribir las horas en formato HH:MM:SS (Ej: 06:30:00).\nDetalle: {str(e)}"
        )

# =========================================================================
# INTERFAZ GRÁFICA DE ESCRITORIO CON CONFIGURACIÓN DE HORARIOS
# =========================================================================
ventana = tk.Tk()
ventana.title("Analizador de Asistencia Configurable - iEDGEED")
ventana.geometry("780x680")
ventana.configure(bg="#F4F6F7")

# Título Principal
lbl_titulo = tk.Label(ventana, text="Control Manual y Configurable de Asistencia", font=("Arial", 15, "bold"), bg="#F4F6F7", fg="#2C3E50")
lbl_titulo.pack(pady=10)

# ----------------- PANEL DE CONFIGURACIÓN DE HORAS -----------------
frame_horas = tk.LabelFrame(ventana, text=" Configurar Horarios del Turno (Formato 24h -> HH:MM:SS) ", font=("Arial", 10, "bold"), bg="#F4F6F7", fg="#34495E", padx=10, pady=10)
frame_horas.pack(pady=10, fill="x", padx=30)

# Variables reactivas vinculadas a los campos de texto
dict_horas_var = {
    'entrada': tk.StringVar(value="06:30:00"),
    'inicio_almuerzo': tk.StringVar(value="13:45:00"),
    'fin_almuerzo': tk.StringVar(value="14:45:00"),
    'salida': tk.StringVar(value="15:30:00")
}

# Grid para estructurar los inputs de texto de los horarios en una sola línea
tk.Label(frame_horas, text="Hora Entrada:", bg="#F4F6F7", font=("Arial", 9, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Entry(frame_horas, textvariable=dict_horas_var['entrada'], width=12, justify="center").grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_horas, text="Inicio Almuerzo:", bg="#F4F6F7", font=("Arial", 9, "bold")).grid(row=0, column=2, padx=5, pady=5, sticky="e")
tk.Entry(frame_horas, textvariable=dict_horas_var['inicio_almuerzo'], width=12, justify="center").grid(row=0, column=3, padx=5, pady=5)

tk.Label(frame_horas, text="Fin Almuerzo:", bg="#F4F6F7", font=("Arial", 9, "bold")).grid(row=0, column=4, padx=5, pady=5, sticky="e")
tk.Entry(frame_horas, textvariable=dict_horas_var['fin_almuerzo'], width=12, justify="center").grid(row=0, column=5, padx=5, pady=5)

tk.Label(frame_horas, text="Hora Salida:", bg="#F4F6F7", font=("Arial", 9, "bold")).grid(row=0, column=6, padx=5, pady=5, sticky="e")
tk.Entry(frame_horas, textvariable=dict_horas_var['salida'], width=12, justify="center").grid(row=0, column=7, padx=5, pady=5)

# ----------------- PANEL SELECCIÓN DE ARCHIVO -----------------
frame_archivo = tk.Frame(ventana, bg="#F4F6F7")
frame_archivo.pack(pady=10)

def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo de asistencia",
        filetypes=[("Archivos de Excel antiguos", "*.xls"), ("Todos los archivos", "*.*")]
    )
    if ruta:
        lbl_archivo.config(text=f"Archivo cargado: {os.path.basename(ruta)}", fg="#27AE60")
        btn_analizar.config(state=tk.NORMAL, command=lambda: empezar_analisis(ruta))

btn_buscar = tk.Button(frame_archivo, text="📂 Buscar Archivo de Asistencia (.xls)", font=("Arial", 10, "bold"), bg="#34495E", fg="white", padx=10, command=seleccionar_archivo)
btn_buscar.pack(side=tk.LEFT, padx=10)

lbl_archivo = tk.Label(ventana, text="Ningún archivo seleccionado aún", font=("Arial", 9, "italic"), bg="#F4F6F7", fg="#7F8C8D")
lbl_archivo.pack(pady=2)

# Botón de Procesamiento Principal
btn_analizar = tk.Button(ventana, text="🚀 Procesar con Horarios Actuales", font=("Arial", 12, "bold"), bg="#2ECC71", fg="white", state=tk.DISABLED, pady=5)
btn_analizar.pack(pady=10)

# Consola Integrada de Visualización de Logs
lbl_consola = tk.Label(ventana, text="Resultados del Análisis:", font=("Arial", 10, "bold"), bg="#F4F6F7", fg="#2C3E50")
lbl_consola.pack(anchor="w", padx=35, pady=2)

txt_consola = scrolledtext.ScrolledText(ventana, width=95, height=18, font=("Consolas", 9), bg="#1E1E1E", fg="#FFFFFF")
txt_consola.pack(padx=25, pady=5)

def empezar_analisis(ruta):
    # Limpiar logs de ejecuciones previas
    txt_consola.delete('1.0', tk.END)
    btn_buscar.config(state=tk.DISABLED)
    btn_analizar.config(state=tk.DISABLED)
    
    # Lanzar el backend core del software
    procesar_asistencia(ruta, txt_consola, dict_horas_var)
    
    # Restablecer los botones al terminar
    btn_buscar.config(state=tk.NORMAL)
    btn_analizar.config(state=tk.NORMAL)

ventana.mainloop()