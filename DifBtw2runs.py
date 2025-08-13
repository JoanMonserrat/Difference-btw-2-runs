import pandas as pd
import uuid
import tkinter as tk
import os
import openpyxl 
import pyodbc 


#Botón con toda la lógica del script, si se usa el script por un tiempo largo sería buena práctica dividir en funciones más pequeñas para testing
def on_button_click():
    try:
        # Verificar pandas
        print("Verificando pandas...")
        import pandas as pd
        print("pandas está disponible.")
    except ImportError as e:
        print(f"Error al importar pandas: {e}")

    try:
        # Verificar openpyxl
        print("Verificando openpyxl...")
        import openpyxl
        print("openpyxl está disponible.")
    except ImportError as e:
        print(f"Error al importar openpyxl: {e}")

    try:
        # Verificar pyodbc
        print("Verificando pyodbc...")
        import pyodbc
        print("pyodbc está disponible.")
    except ImportError as e:
        print(f"Error al importar pyodbc: {e}")
        
    file_path = r"C:\\Users\\Dif2runs.xlsx"

    if os.path.exists(file_path):
        print("Directorio actual:", os.getcwd())
        print("Archivo encontrado.Procesando.")
        df = pd.read_excel(file_path)
        print(df)
    else:
        print("No está el archivo en el path con el nombre correcto.")







# Lógica para la preparación del dataframe a insertar:
# Step 1 : Creo que lo primero, sería modificar el dataframe para que cada test tenga tantas líneas como valores de threshold introducidos y a cada línea darle un ranglelistID. Insertamos este dataframe en:
# Primera aproximación para tomar según el número de tresholds que tenga cada test, que se generen 1, 2 o 3 líneas con ese test y ese valor de threshold para insertar:

data = []

df = pd.read_excel('Dif2runs.xlsx')
for idx, row in df.iterrows():
    nombre = row[0]  # Primera columna
    valores = row[1:]  # El resto de columnas
    for v in valores:
        data.append({'nombre': nombre, 'valor': int(v)})

# Crear nuevo DataFrame con la estructura deseada
df_new = pd.DataFrame(data)

print(df_new)

# Añadir un número creciente a cada una de las filas generadas que serán el RangedList:

df_new['numero'] = range(1, len(df_new) + 1)

print(df_new)

#def_new sería el dataframe a insertar en la primera tabla, ya debería tener todo lo correcto

# Guardar el nuevo DataFrame a Excel para revisar
df_new.to_excel(r"C:\\Users\\1TableDif2Runs.xlsx", index=False, engine='openpyxl')
  
# Step 2:  A partir del dataframe generado previamente, generamos uno nuevo, juntando en una linea los un test ID y los diferentes RangeListID identificados. Insertamos este dataframe en: 



# Agrupar por 'nombre' y juntar los números en una cadena separada por coma. Revisar si así me va bien

df_grouped = df_new.groupby('nombre').agg({
    'valor': 'first',  
    'numero': lambda x: ', '.join(map(str, x))
}).reset_index()

print(df_grouped)

# Con esto tendríamos el segundo dataframe, "df_grouped" que se insertaría en la segunda tabla.

  
#Una vez generado el segundo dataframe, lo guardamos por si fuera necesario revisar:
df_grouped.to_excel(r"C:\\Users\\2TableDif2Runs.xlsx", index=False, engine='openpyxl')







#Conexión a la DB. Usaremos pyodbc.
    df = pd.read_excel("C:\\Users\\dif2runsmodified.xlsx") 


#Conexión ODBC a caché (se podría hacer con DSN, pero se tendría que configurar, ver primero si podemos hacerlo sin DSN, haciendo toda la lista de parámetros asi:
    
    dsn = 'Script connection'  # Aquí reemplazar 'NombreDSN' con el nombre del DSN configurado
    user = '_SYSTEM' 
    password = 'INFINITY'

    try:
        conn = pyodbc.connect(f'DSN={dsn};UID={user};PWD={password}')
        print("Conexión exitosa a la base de datos")
    
        cursor = conn.cursor()

#Insert primera tabla. Hay que modificar valores tanto de la tabla SQL como del dataframe

        for index, row in df_new.iterrows():
            cursor.execute(
                "INSERT INTO core_test_config_dm_tables.Diff2Range(columna1SQL, columna2SQL, columna3SQL, columna4SQL, columna5SQL, columna6SQL, columna7SQL) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                row ['valor 1 dataframe'], 0, row ['valor 2 dataframe'], row ['valor 3 dataframe'], row ['valor 4 dataframe']
    )
        conn.commit()
        
        print("✅ Datos insertados en tabla Diff2Range") #Debería juntar todo esto en un bloque try y hacer otro except para detectar errores.

#Insert segunda tabla SQL en nLO
        for _, row in df_grouped.iterrows():
            cursor.execute(
                "INSERT INTO core_test_config_dm_tables.Diff2RunsRangeVersions "
                "(columna1SQL, columna2SQL, columna3SQL) VALUES (?, ?, ?)",
                row['valor 1 dataframe'], row['valor 2 dataframe'], row['valor 3 dataframe']
            )
        conn.commit()
        cursor.close()
        conn.close()
        print("✅ Datos insertados en tabla Diff2RunsRangeVersions")
       
    except pyodbc.Error as e:
        print(f"Error de conexión: {e}")

# GUI muy simple con tkinter   
root = tk.Tk()
root.title("Dif. Btw 2 runs script")

label = tk.Label(root, text="\u26A0\u26A0\u26A0 Before running the script, please follow the documentation.\n Resumed: \n First step, take the .xml from the MPL DB, open an excel, click on Data>From othe sources and use the.xml.\n Second step, place the .xls file in the folder C:\\Users\\    with the following name: Dif2runs.xlsx\u26A0\u26A0\u26A0")
label.pack(pady=20)

button = tk.Button(root, text="Script!", command=on_button_click)
button.pack(pady=10)

root.mainloop()
