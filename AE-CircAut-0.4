#========= Módulos =============

import pandas as pd
import shutil
import os

#========= Rutas de Archivos =============

año = '2023'

# Definir la ruta base para los archivos
base_ruta = f'/Users/luislopeztrejo/Library/CloudStorage/OneDrive-Personal/Universidad/CodeAcademy/Machine Learning and AI Fundamentals Skill Path/Python Fundamentals/Exercises/Contador escaños por circunscripcion/4.0/Excel/{año}/'

# Definir los nombres de los archivos
archivo_votos = f'Datos/Resultados {año}.xlsx'
archivo_poblacion = f'Datos/Poblacion {año} C.xlsx'

archivo_resultados = f'Resultados/Resultados {año} Circ CCAA - FASE I.xlsx'
archivo_resultados_FASE_II = f'Resultados/Resultados {año} Circ CCAA - FASE II.xlsx'
archivo_resultados_FASE_III = f'Resultados/Resultados {año} Circ CCAA - FINAL.xlsx'

# Actualizar las rutas usando la ruta base
ruta_archivo_votos = base_ruta + archivo_votos
ruta_archivo_poblacion = base_ruta + archivo_poblacion
ruta_archivo_resultados = base_ruta + archivo_resultados
ruta_archivo_resultados_FASE_II = base_ruta + archivo_resultados_FASE_II
ruta_archivo_resultados_FASE_III = base_ruta + archivo_resultados_FASE_III

#========= Tamaño Circunscripciones =============

min_escanos = 4
ceut_mel_escanos = 2
max_escanos = 350

print('Mínimo de escaños: {}, Total escaños: {}'.format(min_escanos, max_escanos))

#================================================================================================================================================



#========= Funciones Excel =============


# Función para cargar tabla Excel
def cargar_tabla_excel(ruta_archivo):
    excel_file = pd.ExcelFile(ruta_archivo)
    datos_por_hoja = {}
    for hoja in excel_file.sheet_names:
        df = excel_file.parse(hoja)
        columna1 = df.iloc[:, 0]
        columna2 = df.iloc[:, 1]
        comb_colum1_colum_2 = list(zip(columna1, columna2))
        datos_por_hoja[hoja] = comb_colum1_colum_2
    return datos_por_hoja


#========= Funcion D'Hondt =============


# Función para calcular escaños según método D'Hondt y guardarlos en un archivo Excel


def dhondt(n_escanos, votos):
    """
    Implementación del método D'Hondt para distribuir escaños.

    Args:
        n_escanos (int): Número total de escaños a repartir.
        votos (dict): Diccionario con los votos de cada partido.

    Returns:
        dict: Diccionario con los escaños asignados a cada partido.
    """

    # Crea una copia de los votos y un diccionario para rastrear los escaños asignados

    t_votos = votos.copy()
    escanos = {}

    # Inicializa todos los partidos con 0 escaños
    for voto in votos:
        escanos[voto] = 0

    # Mientras no se alcance el número de escaños deseados
    while sum(escanos.values()) < n_escanos:
        # Encuentra el partido con más votos
        max_v = max(t_votos.values())
        sig_escano = list(t_votos.keys())[list(t_votos.values()).index(max_v)]

        # Asigna un escaño al partido
        if sig_escano in escanos:
            escanos[sig_escano] += 1
        else:
            escanos[sig_escano] = 1

        # Actualiza los votos del partido para el próximo ciclo
        t_votos[sig_escano] = votos[sig_escano] / (escanos[sig_escano] + 1)

    # Filtra y devuelve solo los partidos que obtienen escaño
    partidos_escano = {voto: escanos[voto] for voto in escanos if escanos[voto] > 0}
    return partidos_escano

#================================================================================================================================================

#========= Cálculo Resultados =============


# Función para calcular los resultados y guardarlos en un archivo Excel
def calcular_resultados():
    """
    Calcula los resultados y los guarda en un archivo Excel.

    Utiliza el método D'Hondt para distribuir los escaños basados en los votos.

    También genera un resumen de los resultados y lo guarda en otro archivo Excel.
    """

    # Cargar datos de votos desde un archivo Excel
    partidos_votos = cargar_tabla_excel(ruta_archivo_votos)

    # Crear una lista para almacenar las tuplas formateadas de cada hoja
    partidos_votos_formated_all = []

    # Almacenar los datos de cada hoja formateada en la lista de tuplas
    for datos in partidos_votos.values():
        partidos_votos_formated_all.append({partido: votos for partido, votos in datos})

    # Importar datos de población desde un archivo Excel
    xls_poblacion_ccaa = pd.read_excel(ruta_archivo_poblacion)

    # Crear un DataFrame con los datos de población
    df_poblacion_ccaa = pd.DataFrame(xls_poblacion_ccaa)

    # Crear una lista de tuplas con el nombre de la CCAA y su población
    poblacion_ccaa = [(row['ccaa'], row['poblacion']) for _, row in df_poblacion_ccaa.iterrows()]

    # Calcular el total de la población
    total_poblacion = sum(poblacion for _, poblacion in poblacion_ccaa)

    # Calcular tamaño de las circunscripciones
    escanos_total = []

    #====================================================== 

    #========= Asignación escaños por población ========= 


    # Cargar los datos de población
    xls_poblacion_ccaa = pd.read_excel(ruta_archivo_poblacion)
    df_poblacion_ccaa = pd.DataFrame(xls_poblacion_ccaa)


    #========= Ceuta y Melilla ========= 

    # Obtener la población de Ceuta y Melilla
    poblacion_ceuta = df_poblacion_ccaa.loc[df_poblacion_ccaa['ccaa'] == 'Ceuta']['poblacion'].values[0]
    poblacion_melilla = df_poblacion_ccaa.loc[df_poblacion_ccaa['ccaa'] == 'Melilla']['poblacion'].values[0]

    # Calcular la población total excluyendo la de Ceuta y Melilla
    poblacion_total_resto = total_poblacion - (poblacion_ceuta + poblacion_melilla)

    # Asignar 2 escaños para Ceuta y Melilla
    escanos_total.append(('Ceuta', ceut_mel_escanos))
    escanos_total.append(('Melilla', ceut_mel_escanos))

    #========= Resto de Circunscripciones ========= 

    # Calcular el número total de escaños disponibles para el resto de las circunscripciones
    total_escanos_resto = max_escanos - ceut_mel_escanos * 2  # Restar 4 escaños de Ceuta y Melilla

    # Calcular la población total excluyendo la de Ceuta y Melilla
    poblacion_total_resto = total_poblacion - (poblacion_ceuta + poblacion_melilla)

    # Crear un diccionario para rastrear los escaños asignados a cada circunscripción
    escaños_asignados = {nombre: 0 for nombre, _ in poblacion_ccaa if 'Ceuta' not in nombre and 'Melilla' not in nombre}

    # Asignar 4 escaños al resto de circunscripciones
    for nombre, _ in poblacion_ccaa:
        if 'Ceuta' not in nombre and 'Melilla' not in nombre:
            escaños_asignados[nombre] = min_escanos
            
            #Restamos los escaños ya dados del total de escaños
            total_escanos_resto -= min_escanos

    # D'Hondt para asignar escaños adicionales basados en la población
    while total_escanos_resto > 0:
        
        #Máximo de población por escaño en las circunscripcion
        max_ratio = 0

        #Circunscripción a la que asignamos escaño
        circunscripcion_ganadora = None

        for nombre, poblacion in poblacion_ccaa:
            if nombre in escaños_asignados:
                ratio = poblacion / (escaños_asignados[nombre] + 1)
                if ratio > max_ratio:
                    max_ratio = ratio
                    circunscripcion_ganadora = nombre

        if circunscripcion_ganadora:
            escaños_asignados[circunscripcion_ganadora] += 1
            total_escanos_resto -= 1
        else:
            # Si no se puede asignar más escaños, salimos del bucle
            break

    # Asignar los escaños calculados a las circunscripciones
    for nombre in escaños_asignados:
        escanos_total.append((nombre, escaños_asignados[nombre]))

    #=====================================================

    #========= Votos por circunscripción ========= 

    # Utilizar D'Hondt para distribuir los escaños basados en los votos **************
    resultados = []

    for ccaa, votos_ccaa in zip(escanos_total, partidos_votos_formated_all):
        nombre, n_escanos = ccaa
        votos = votos_ccaa

        escanos = dhondt(n_escanos, votos)

        if escanos:
            resultados.append((nombre, escanos))


    # Crear un DataFrame con los resultados y guardarlos en un archivo Excel
    df_resultados = pd.DataFrame(resultados, columns=['CCAA', 'Resultados'])
    df_resultados['Resultados'] = df_resultados['Resultados'].apply(lambda x: str(x).strip('{}'))

    print(df_resultados.to_string(index=False, header=None))

    df_resultados.to_excel(ruta_archivo_resultados, index=False)

    # Calcular totales de escaños por partido
    totales_escaños = {}

    for _, escaños in resultados:
        for partido, escaños_partido in escaños.items():
            if partido in totales_escaños:
                totales_escaños[partido] += escaños_partido
            else:
                totales_escaños[partido] = escaños_partido

    # Imprimir el total de escaños por partido
    for partido, total_escaños in totales_escaños.items():
        print(f"{partido}: {total_escaños}")


    #========= DataFrames =============
    

    # Crear un DataFrame con los totales de escaños por partido
    data_totales_escaños = {'Partido': list(totales_escaños.keys()), 'Total Escaños': list(totales_escaños.values())}
    df_totales_escaños = pd.DataFrame(data_totales_escaños)

    # Leer un archivo Excel existente
    df_existente_temp = pd.read_excel(ruta_archivo_resultados)

    # Escribir en un nuevo archivo Excel con múltiples hojas
    with pd.ExcelWriter(ruta_archivo_resultados_FASE_II, engine='openpyxl') as writer:
        df_existente_temp.to_excel(writer, sheet_name=f'Circunscripciones {año}', index=False)
        df_totales_escaños.to_excel(writer, sheet_name=f'Total {año}', index=False)

    # Copiar un archivo a otro destino
    shutil.copy(ruta_archivo_resultados_FASE_II, ruta_archivo_resultados_FASE_III)

    # Crear un DataFrame con la distribución de escaños por CCAA y guardarlos en un archivo Excel
    df_escanos_ccaa = pd.DataFrame(escanos_total, columns=['CCAA', 'Número de Escaños'])

    with pd.ExcelWriter(ruta_archivo_resultados_FASE_III, engine='openpyxl', mode='a') as writer:
        df_escanos_ccaa.to_excel(writer, sheet_name=f'Distribución de Escaños {año}', index=False)

    # Imprimir el total de escaños
    print("Total:", sum(totales_escaños.values()))
    print(df_escanos_ccaa)

    #≠≠≠≠≠≠≠≠≠≠≠≠≠≠ Ñapa ≠≠≠≠≠≠≠≠≠≠≠≠≠≠

    # Eliminar los archivos temporales
    os.remove(ruta_archivo_resultados)
    os.remove(ruta_archivo_resultados_FASE_II)


#========= Let's Go! =============

calcular_resultados()
