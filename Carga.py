
print("Iniciar carga")

import pyodbc
import json
import random
import string

# Datos de conexión
servidor = "172.19.3.20"
usuario = "DSNet_APP"
contrasena = "7b*ZfrR2W"
base_datos = "DaltSoftTGDL"
rutaDeJson = r"C:\Users\Oscar Omar\Desktop\CargaDalton\Articulos.json"

cadena_conexion = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={servidor};DATABASE={base_datos};UID={usuario};PWD={contrasena}"

# Conexión a la base de datos
try:
    conexion = pyodbc.connect(cadena_conexion)
    print("Conexión exitosa.")
    cursor = conexion.cursor()
    
    with open(rutaDeJson) as archivo:
        data = json.load(archivo)  
        Articulos = data["Articulos"]
    consultaSelect =""
    for articulo in Articulos: 
         descripcion = articulo['DESCRIPCION'] 
         descripcionSinEspacios = descripcion.replace(" ","")
         
         codigo = descripcionSinEspacios[0:3] + descripcionSinEspacios[-3:]
         centro_index = len(descripcionSinEspacios) // 2

         if len(descripcionSinEspacios) % 2 == 1:
            codigo += descripcionSinEspacios[centro_index]
         else:
            codigo += descripcionSinEspacios[centro_index - 1]
        
         codigo += descripcionSinEspacios[5:6]
         letra_aleatoria = random.choice(string.ascii_uppercase)  # Genera una letra aleatoria de A a Z
         codigo += letra_aleatoria
         params = (
            1,                                      # Primer parámetro
            codigo,                              # Segundo parámetro
            articulo['DESCRIPCION'],                # Tercer parámetro (DESCRIPCION)
            'A',                                    # Cuarto parámetro
            None,                                   # Quinto parámetro
            'CA',                      # Canal
            None,                                   # Séptimo parámetro
            None,                                   # Octavo parámetro
            None,                                   # Noveno parámetro
            None,                                   # Décimo parámetro
            'N/A',                                  # Undécimo parámetro
            'MN',                                   # Duodécimo parámetro
            'MN',                                   # Decimotercer parámetro
            'IVA',                                  # Decimocuarto parámetro
            0,                                      # Decimoquinto parámetro
            0,                                      # Decimosexto parámetro
            'E',                                     # Decimoséptimo parámetro
            None,                                   # Decimoctavo parámetro
            None,                                   # Decimonoveno parámetro
            None,                                   # Vigésimo parámetro
            0,                                      # Vigésimo primero
            'N',                                    # Vigésimo segundo
            0,                                      # Vigésimo tercero
            0,                                      # Vigésimo cuarto
            'N',                                    # Vigésimo quinto
            0,                                      # Vigésimo sexto
            None,                                   # Vigésimo séptimo
            None,                                   # Vigésimo octavo
            None,                                   # Vigésimo noveno
            0,                                      # Trigésimo
            None,                                   # Trigésimo primero
            None,                                   # Trigésimo segundo
            None,                                   # Trigésimo tercero
            None,                                   # Trigésimo cuarto
            0,                                      # Trigésimo quinto
            None,                                   # Trigésimo sexto
            'ADMON',                                # Trigésimo séptimo
            'ADMON',                                # Trigésimo octavo
            None,                                   # Trigésimo noveno
            None,                                   # Cuadragésimo
            'N',                                    # Cuadragésimo primero
            1,                                      # Cuadragésimo segundo
            'O',                                    # Cuadragésimo tercero
            0,                                      # Cuadragésimo cuarto
            None,                                   # Cuadragésimo quinto
            0,                                      # Cuadragésimo sexto
            None,                                   # Cuadragésimo séptimo
            None,                                   # Cuadragésimo octavo
            0,                                      # Cuadragésimo noveno
            None,                                   # Quincuagésimo
            None,                                   # Quincuagésimo primero
            None,                                   # Quincuagésimo segundo
            'False',                                # Quincuagésimo tercero
            '0',                                    # Quincuagésimo cuarto
            False,                                  # Quincuagésimo quinto
            0,                                      # Quincuagésimo sexto
            None,                                   # Quincuagésimo séptimo
            articulo['PRODUCTO'],             # Valor de btpproductart
            articulo['SUBPRODUCTO'],             # Valor de btpsubrproduct
            None,                                   # Quincuagésimo noveno
            None,                                   # Sexagésimo
            'False',                                # Sexagésimo primero
            0,                                      # Sexagésimo segundo
            0,                                      # Sexagésimo tercero
            '80101500',                             # Clave SAT
            'IVA',                                  # Valor para IVA
            0,                                      
            0,
            0                                      
        )
        
         cursor.execute("{CALL SP_DSInvArt(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)}", params)
         cursor.execute(f"UPDATE DSInvArt SET BPTpProd_Art = '{articulo['PRODUCTO']}', BPTpSubProd_Art = '{articulo['SUBPRODUCTO']}', BPCanal_Art = '{articulo['CANAL']}' where Cod_Art = '{codigo}';")
         cursor.execute(f"insert into DSCxpCptosXArt (Cod_Art, Cod_CptoCxp,Cod_CC, FyHReg_RGas, IdAsig_Conce) values('{codigo}','{articulo['CODIGOCXP']}',06,'2024-11-01 15:58:00','DT')")
        
         #Inicio de bucle
         consultaSql = f"'{codigo}',  "
         consultaSelect += consultaSql
         print(consultaSelect)
         
    # Ejecutar el procedimiento almacenado
    
    conexion.commit()
    

except Exception as e:
    print(f"Error al conectar: {e}")

finally:
    conexion.close()
