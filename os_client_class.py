import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import csv
from unidecode import unidecode
import matplotlib.pyplot as plt

# Funcion para cargar de los datos provenientes del archivo csv
def cargar_csv(archivo_ventas):
    with open(archivo_ventas, mode='r', encoding='latin1') as archivo:
        lector_csv = csv.reader(archivo)
        datos = list(lector_csv)
    return datos

# *********************FUNCIONES DE LIMPIEZA************************

# Funciones para limpiar los datos y eliminar columnas que no se ocupan, quitar acentos y espacios
def limpiar_datos(lista_datos):
    indices_a_eliminar = [0, 2, 3,5,7,8,10,11,12,13,14,16]
    datos_limpiados = []
    for fila in lista_datos:
        if len(fila) > 1:
            fila_filtrada = [dato for i, dato in enumerate(fila) if i not in indices_a_eliminar]
            datos_limpiados.append(fila_filtrada)
    return datos_limpiados

def limpiar_acentos(columna):
    columna = unidecode(columna)
    return columna.strip().lower()

def lista_a_dataframe(ventas_datos_limpiados):
    columnas = ventas_datos_limpiados[0]
    datos = ventas_datos_limpiados[1:]
    columnas_limpias = [limpiar_acentos(columna) for columna in columnas]
    df = pd.DataFrame(datos, columns=columnas_limpias)
    return df

def limpiar_columnas_numericas(df, columnas_numericas):
    for col in columnas_numericas:
        if col in df.columns:
            df[col] = df[col].replace(",", "", regex=True).astype(float)
    return df

# *******************************FUNCIONES DE CLASIFICACIÓN**********************************

# Funcion que clasifique las horas, tomando como base el index de horas reales,
# haciendo que se pueda ver la clasificacion del cliente como numero de horas totales en un rago historico
def clasificar_por_horas(df):
    try:
        # Verificar que la columna hrs reales exista ya que hay otra columna de horas
        if 'hrs reales' not in df.columns:
            messagebox.showerror("Error", "La columna 'hrs reales' no existe en el archivo CSV.")
            return None

        # creacion de las categorias base para el numero de horas
        condiciones = [
            (df['hrs reales'] <= 1),
            (df['hrs reales'] > 1) & (df['hrs reales'] <= 5),
            (df['hrs reales'] > 5) & (df['hrs reales'] <= 10),
            (df['hrs reales'] > 10)
        ]
        categorias = ['Bajas horas', 'Horas medias', 'Altas horas', 'Muy altas horas']

        # Asignar categoría a cada fila según los resultados de la clasificacin
        df['Clasificación por Horas'] = pd.cut(df['hrs reales'], bins=[-float('inf'), 1, 5, 10, float('inf')], labels=categorias)

        return df
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema al clasificar las horas: {str(e)}")
        return None

# Funcion que clasifica por tipo de servicio-ciudad, los tipos de servicio remoto o presencial y la ciudad
def clasificar_por_tipodeserv_ciudad(df):
    try:
        # Verificar que la columna exista, debido a los varios index que hay
        if 'tipodeserv-ciudad' not in df.columns:
            messagebox.showerror("Error", "La columna 'tipodeserv-ciudad' no existe en el archivo CSV.")
            return None

        # Limpieza la columna tipodeserv-ciudad: eliminar espacios al inicio y al final, y reemplazar múltiples espacios por uno solo
        df['tipodeserv-ciudad'] = df['tipodeserv-ciudad'].str.strip()  
        df['tipodeserv-ciudad'] = df['tipodeserv-ciudad'].str.replace(r'\s+', ' ', regex=True)  

        # Agrupar por razon social y tipodeserv-ciudad y contar las incidencias
        incidencias = df.groupby(['razon social', 'tipodeserv-ciudad']).size().reset_index(name='Incidencias')

        # Obtener la combinacion con la mayor incidencia para posteriormente desplegarlo en la grafica
        max_incidencia = incidencias.loc[incidencias['Incidencias'].idxmax()]

        return incidencias, max_incidencia
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema al clasificar por tipo de servicio y ciudad: {str(e)}")
        return None

# ***************************LAS GRÁFICAS DE LAS FUNCIONES DE CLASIFICACIÓN**************
def graficar_incidencias_tipodeserv_ciudad(incidencias):
    try:
        # obtener el conteno de las incedencias y agruparlas
        incidencia_counts = incidencias.groupby('tipodeserv-ciudad')['Incidencias'].sum().sort_values(ascending=False)

        # Graficas
        incidencia_counts.plot(kind='bar', color='lightcoral', edgecolor='black')
        plt.title("Incidencias por Tipo de Servicio y Ciudad")
        plt.xlabel("Tipo de Servicio y Ciudad")
        plt.ylabel("Número de Incidencias")
        plt.tight_layout()
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema al graficar las incidencias: {str(e)}")

def graficar_clientes_con_mas_horas(df):
    try:
        df_horas_totales = df.groupby("razon social")["hrs reales"].sum().sort_values(ascending=False)
        df_horas_totales.head(10).plot(kind='bar', color='lightblue', edgecolor='black')
        plt.title("Clientes con más Horas Reales")
        plt.xlabel("Cliente")
        plt.ylabel("Total de Horas Reales")
        plt.tight_layout()
    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema al graficar los clientes con más horas: {str(e)}")

#*********FUNCION DE CARGA DE ARCHIVO  CSV PARA INTERFAZ GRAFICA Y GUARDA ARCHIVO DE EXCEL PARA APOYAR EL ANALISIS****************
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Seleccione un archivo CSV",
        filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
    )
    if archivo:
        try:
            # Cargar, limpiar y procesar datos
            datos = cargar_csv(archivo)
            ventas_datos_limpiados = limpiar_datos(datos)
            df_limpio = lista_a_dataframe(ventas_datos_limpiados)
            df_limpio = limpiar_columnas_numericas(df_limpio, ["hrs reales"])

            # Clasificar por horas
            df_clasificado_horas = clasificar_por_horas(df_limpio)

            # Clasificar por tipo de servicio y ciudad
            incidencias_tipodeserv, max_incidencia = clasificar_por_tipodeserv_ciudad(df_limpio)

            # Guardar el archivo con los datos
            incidencias_tipodeserv.to_excel("clasificacion_por_tipodeserv_ciudad.xlsx", index=False)

            # Mostrar ambas gráficas en una sola ventana 
            fig, axes = plt.subplots(1, 2, figsize=(14, 6)) 

            # Graficar clientes con más horas
            axes[0].bar(df_limpio.groupby("razon social")["hrs reales"].sum().sort_values(ascending=False).head(10).index,
                        df_limpio.groupby("razon social")["hrs reales"].sum().sort_values(ascending=False).head(10).values,
                        color='lightblue', edgecolor='black')
            axes[0].set_title("Clientes con más Horas Reales")
            axes[0].set_xlabel("Cliente")
            axes[0].set_ylabel("Total de Horas Reales")
            axes[0].tick_params(axis='x', rotation=45)

            # Graficar incidencias por tipo de servicio y ciudad
            incidencias_counts = incidencias_tipodeserv.groupby('tipodeserv-ciudad')['Incidencias'].sum().sort_values(ascending=False)
            axes[1].bar(incidencias_counts.index, incidencias_counts.values, color='lightcoral', edgecolor='black')
            axes[1].set_title("Incidencias por Tipo de Servicio y Ciudad")
            axes[1].set_xlabel("Tipo de Servicio y Ciudad")
            axes[1].set_ylabel("Número de Incidencias")
            axes[1].tick_params(axis='x', rotation=90)

            plt.tight_layout()
            plt.show()

            # Mensaje final
            messagebox.showinfo("Éxito", f"Clasificación por tipo de servicio y ciudad guardada en 'clasificacion_por_tipodeserv_ciudad.xlsx'. Las gráficas también fueron generadas.")
        except Exception as e:
            messagebox.showerror("Error", f"Hubo un problema: {str(e)}")
    else:
        messagebox.showwarning("Advertencia", "No seleccionó ningún archivo.")

# Interfaz gráfica para cargar el archivo csv y no seleccionarlo directamente de una raiz
ventana = tk.Tk()
ventana.title("Análisis de Clientes apartir de las OS")
ventana.geometry("400x200")

btn_seleccionar = tk.Button(
    ventana,
    text="Selecciona Archivo CSV",
    command=seleccionar_archivo,
    font=("Arial", 12),
    bg="lightgreen"
)
btn_seleccionar.pack(pady=50)

ventana.mainloop()
