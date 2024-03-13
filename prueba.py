import pandas as pd
import plotly.express as px
import webbrowser

# Cargar el archivo Excel con múltiples páginas
xls = pd.ExcelFile("sexo_pais_random.xlsx")

# Crear una lista para almacenar los nombres de archivo HTML generados
html_files = []

# Crear un diccionario para almacenar los DataFrames de cada página
dfs = {}

# Iterar sobre cada página del archivo Excel
for sheet_name in xls.sheet_names:
    # Leer la página del archivo Excel y almacenarla en el diccionario
    dfs[sheet_name] = pd.read_excel("sexo_pais_random.xlsx", sheet_name=sheet_name)

    # Eliminar las filas y columnas vacías
    dfs[sheet_name] = dfs[sheet_name].dropna(axis=0, how='all').dropna(axis=1, how='all')

    # Transponer el DataFrame para tener las nacionalidades como columnas
    dfs[sheet_name] = dfs[sheet_name].T

    # Tomar la primera fila como nombres de columnas
    dfs[sheet_name].columns = dfs[sheet_name].iloc[0]

    # Eliminar la primera fila
    dfs[sheet_name] = dfs[sheet_name][1:]

    # Restablecer el índice
    dfs[sheet_name].reset_index(inplace=True)

    # Renombrar la columna de índice
    dfs[sheet_name].rename(columns={'index': 'Nacionalidad'}, inplace=True)

    # Melt para convertir las columnas en filas
    dfs[sheet_name] = dfs[sheet_name].melt(id_vars='Nacionalidad', var_name='Pais', value_name='Cantidad')

    # Crear el gráfico de barras con la nacionalidad
    fig = px.bar(dfs[sheet_name], x='Nacionalidad', y='Cantidad', color='Pais', barmode='group', title='Nacionalidades')

    # Guardar el gráfico como archivo HTML
    file_name = f"grafico_nacionalidades_{sheet_name}.html"
    fig.write_html(file_name)
    html_files.append(file_name)

    # Crear el gráfico de población por sexo en círculo
    fig_pie_sexo = px.pie(dfs[sheet_name], names='Nacionalidad', values='Cantidad', title='Población por Sexo')
    fig_pie_sexo.update_traces(textposition='inside', textinfo='percent+label')

    # Guardar el gráfico como archivo HTML
    file_name = f"grafico_poblacion_sexo_{sheet_name}.html"
    fig_pie_sexo.write_html(file_name)
    html_files.append(file_name)

    # Filtrar los datos para obtener solo los valores numéricos en la columna 'Cantidad'
    df_numeric = dfs[sheet_name][dfs[sheet_name]['Cantidad'].apply(lambda x: str(x).isdigit())]

    # Calcular la población chilena y extranjera para el gráfico de población por nacionalidad
    poblacion_chilena = df_numeric.loc[df_numeric['Nacionalidad'] == 'Chilena', 'Cantidad'].astype(int).sum()
    poblacion_extranjera = df_numeric.loc[df_numeric['Nacionalidad'] != 'Chilena', 'Cantidad'].astype(int).sum()

    # Crear el gráfico de población por nacionalidad en círculo
    fig_pie_poblacion = px.pie(names=['Chilena', 'Extranjera'], values=[poblacion_chilena, poblacion_extranjera], title='Población por Nacionalidad')
    fig_pie_poblacion.update_traces(textposition='inside', textinfo='percent+label')

    # Guardar el gráfico como archivo HTML
    file_name = f"grafico_poblacion_nacionalidad_{sheet_name}.html"
    fig_pie_poblacion.write_html(file_name)
    html_files.append(file_name)

# Abrir todos los archivos HTML generados automáticamente
for file_name in html_files:
    webbrowser.open(file_name)

