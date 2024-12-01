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
    indices_a_eliminar = [1, 3, 12]
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
# Funcion especifica para crear el excel coin clasificacion
def generar_excel(df, nombre_archivo_excel):
    resumen = df.describe()
    ventas_por_cliente = df.groupby("razon social")["total"].sum().sort_values(ascending=False)
    ventas_por_fecha = df.groupby("fecha")["total"].sum()
    ventas_pendientes = df[df["pendiente"] > 0]
    
    with pd.ExcelWriter(nombre_archivo_excel, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Datos Limpiados', index=False)
        resumen.to_excel(writer, sheet_name='Resumen Estadístico')
        ventas_por_cliente.to_excel(writer, sheet_name='Ventas por Cliente')
        ventas_por_fecha.to_excel(writer, sheet_name='Ventas por Fecha')
        ventas_pendientes.to_excel(writer, sheet_name='Ventas Pendientes')

    print(f"Archivo Excel '{nombre_archivo_excel}' creado exitosamente!")

# ******************Funciones para las graficas separadas************
def graficar_ventas_totales(df):
    plt.figure(figsize=(10, 6))
    df['total'] = df['total'].replace(",", "", regex=True).astype(float)
    plt.hist(df['total'], bins=20, color='skyblue', edgecolor='black')
    plt.title('Distribución de Ventas Totales')
    plt.xlabel('Ventas Totales ($)')
    plt.ylabel('Frecuencia')
    plt.show()

def graficar_ventas_por_cliente(df):
    plt.figure(figsize=(10, 6))
    ventas_por_cliente = df.groupby("razon social")["total"].sum().sort_values(ascending=False)
    ventas_por_cliente.head(10).plot(kind='barh', color='lightgreen', edgecolor='black')
    plt.title('Ventas por Cliente')
    plt.xlabel('Total de Ventas ($)')
    plt.ylabel('Razón Social')
    plt.show()

def graficar_ventas_por_fecha(df):
    plt.figure(figsize=(10, 6))
    df['fecha'] = pd.to_datetime(df['fecha'], format='%d/%m/%Y')
    ventas_por_fecha = df.groupby("fecha")["total"].sum()
    ventas_por_fecha.plot(kind='line', color='purple', marker='o')
    plt.title('Ventas por Fecha')
    plt.xlabel('Fecha')
    plt.ylabel('Total de Ventas ($)')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

def graficar_ventas_pendientes(df):
    plt.figure(figsize=(10, 6))
    pendientes = df[df['pendiente'] > 0]
    plt.bar(pendientes['razon social'], pendientes['pendiente'], color='salmon')
    plt.title('Ventas con Saldo Pendiente')
    plt.xlabel('Razón Social')
    plt.ylabel('Saldo Pendiente ($)')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

# Columnas numericas separadas
columnas_numericas = ["neto", "descuento", "i.v.a", "total", "pendiente"]

# Interfaz gráfica con funcion para la carga de los archivos y mande a llamar la func del guardado del excel
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Seleccione un archivo CSV",
        filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
    )
    if archivo:
        try:
            datos = cargar_csv(archivo)
            ventas_datos_limpiados = limpiar_datos(datos)
            df_limpio = lista_a_dataframe(ventas_datos_limpiados)
            df_limpio = limpiar_columnas_numericas(df_limpio, columnas_numericas)
            
            # Generar Excel y mostrar gráficas
            generar_excel(df_limpio, "analisis_ventas.xlsx")
            graficar_ventas_totales(df_limpio)
            graficar_ventas_por_cliente(df_limpio)
            graficar_ventas_por_fecha(df_limpio)
            graficar_ventas_pendientes(df_limpio)

            messagebox.showinfo("Éxito", "Análisis completado, gráficos contruidos y Excel de apoyo generado.")
        except Exception as e:
            messagebox.showerror("Error", f"Hubo un problema: {str(e)}")
    else:
        messagebox.showwarning("Advertencia", "No seleccionó ningún archivo.")

# Ventana emergente para carga del archivo 
ventana = tk.Tk()
ventana.title("Procesador de Ventas")
ventana.geometry("400x200")

btn_seleccionar = tk.Button(
    ventana,
    text="Seleccionar Archivo CSV",
    command=seleccionar_archivo,
    font=("Arial", 12),
    bg="lightblue"
)
btn_seleccionar.pack(pady=50)

ventana.mainloop()
