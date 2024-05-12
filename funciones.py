import logging
import os
import pandas as pd
from FuncionesGenerales import conectar2SQL, enviarCorreo
import matplotlib.pyplot as plt
import io
import base64

import utils.constantes as const


dbriesgo, cursor = conectar2SQL()

logger = logging.getLogger(const.PROCESS_NAME)


def convertir_nombre_propio(texto):
    return texto.title()


def obtener_datos():
    try:
        remotos = pd.read_sql(const.QUERY_REMOTOS, dbriesgo)
        remotos_agrupado = pd.read_sql(const.QUERY_REMOTOS_AGRUPADO, dbriesgo)
        remotos_5dias = pd.read_sql(const.QUERY_REMOTOS_5DIAS, dbriesgo)
        remotos_5dias_agrupado = pd.read_sql(
            const.QUERY_REMOTOS_5DIAS_AGRUPADO, dbriesgo
        )

        remotos["Nombre"] = remotos["Nombre"].apply(convertir_nombre_propio)
        remotos = remotos.sort_values(by=["Gerencia", "Unidad", "Rut"])
        remotos.reset_index(inplace=True, drop=True)

        logger.info(
            f"Datos de usuarios remotos obtenidos exitosamente - {len(remotos)} usuarios remotos el dia de hoy"
        )
        return remotos, remotos_agrupado, remotos_5dias, remotos_5dias_agrupado
    except Exception as e:
        raise ValueError(
            f"Ocurrio un error al obtener los datos de usuarios remotos: {e}"
        )


def generar_grafico_remotos_area(remotos_agrupado):
    data = remotos_agrupado.copy()
    data = data.sort_values(by=[const.KEY_PORCENTAJE_REMOTOS])
    fig, ax = plt.subplots(figsize=(10, 5))

    colors = ["#1f77b4" for _ in range(len(data["Gerencia"]))]
    bars = ax.bar(
        x=data["Gerencia"],
        height=data["% Remotos"],
        tick_label=data["Gerencia"],
        color="#1f77b4",
    )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#DDDDDD")
    ax.tick_params(bottom=False, left=False)
    ax.set_axisbelow(True)
    ax.yaxis.grid(True, color="#EEEEEE")
    ax.xaxis.grid(False)
    ax.set_xticklabels(data["Gerencia"], rotation=35, fontsize=8)

    for bar, color in zip(bars, colors):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.1,
            str(int(round(bar.get_height(), 0))) + "%",
            horizontalalignment="center",
            color=color,
            weight="bold",
            fontsize=8,
        )

    # Añadir etiquetas y título
    # ax0.set_xlabel('Gerencia', size=8,labelpad=15, color='#333333')
    ax.set_ylabel("% de Empleados Remoto", size=8, labelpad=15, color="#333333")
    ax.set_title(
        "Porcentaje de Empleados Remotos por Área",
        size=12,
        color="#1f77b4",
        weight="bold",
    )
    # Alinear las etiquetas del eje x a la derecha
    plt.setp(ax.xaxis.get_majorticklabels(), ha="right")

    # Agregar una línea horizontal al 8%
    ax.axhline(y=8, color="#c73636", linestyle="-")

    # Agregar un texto al lado de la línea
    ax.text(len(data["Gerencia"]) + 0.5, 8, "8%", color="#c73636", va="center")

    fig.tight_layout()

    # Crear un objeto BytesIO para almacenar el gráfico en memoria
    buffer = io.BytesIO()
    # Guardar el gráfico en el buffer con el formato y dpi deseados
    fig.savefig(buffer, format="png", dpi=const.DPI_FIGURAS)
    # Codificar el buffer en base64
    b64 = base64.b64encode(buffer.getvalue())

    # Retornar el gráfico en base64
    return b64.decode()


def generar_tabla_remotos_area(remotos_agrupado):
    data = preparar_datos(remotos_agrupado)
    fig, ax = inicializar_grafico(data)
    agregar_texto_tabla(ax, data)
    agregar_nombres_columnas(ax, data)
    agregar_lineas_divisorias(ax, data)
    ax.set_axis_off()
    fig.tight_layout()
    return guardar_grafico(fig)


def preparar_datos(remotos_agrupado):
    data = remotos_agrupado.copy()
    return data.sort_values(by=[const.KEY_PORCENTAJE_REMOTOS])


def inicializar_grafico(data):
    fig, ax = plt.subplots(figsize=(10, 8))
    ncols = data.shape[1]
    nrows = data.shape[0]
    ax.set_xlim(0, ncols + 1)
    ax.set_ylim(0, nrows + 1)
    return fig, ax


def agregar_texto_tabla(ax, data):
    positions = [0.25, 2.5, 3.5, 4.5]
    columns = data.columns
    nrows = data.shape[0]
    for i in range(nrows):
        for j, column in enumerate(columns):
            font_color, ha, text_label, weight = obtener_propiedades_celda(
                i, j, column, data
            )
            ax.annotate(
                xy=(positions[j], i + 0.5),
                text=text_label,
                ha=ha,
                va="center",
                weight=weight,
                color=font_color,
            )


def obtener_propiedades_celda(i, j, column, data):
    font_color = "#4f4e4b"
    ha = "left" if j == 0 else "center"
    if column == const.KEY_PORCENTAJE_REMOTOS:
        text_label = f"{data[column].iloc[i]}%"
        weight = "bold"
        if float(data[column].iloc[i]) >= const.UMBRAL_PORCENTAJE_REMOTO:
            font_color = "#c73636"
    else:
        text_label = f"{data[column].iloc[i]}"
        weight = "normal"
    return font_color, ha, text_label, weight


def agregar_nombres_columnas(ax, data):
    positions = [0.25, 2.5, 3.5, 4.5]
    column_names = data.columns.to_list()
    nrows = data.shape[0]
    for index, c in enumerate(column_names):
        ha = "left" if index == 0 else "center"
        ax.annotate(
            xy=(positions[index], nrows + 0.25),
            text=column_names[index],
            ha=ha,
            va="bottom",
            weight="bold",
            color="#1f77b4",
        )


def agregar_lineas_divisorias(ax, data):
    nrows = data.shape[0]
    ax.plot(
        [ax.get_xlim()[0], ax.get_xlim()[1]],
        [nrows, nrows],
        lw=1.5,
        color="#96948d",
        marker="",
        zorder=4,
    )
    ax.plot(
        [ax.get_xlim()[0], ax.get_xlim()[1]],
        [0, 0],
        lw=1.5,
        color="#96948d",
        marker="",
        zorder=4,
    )
    for x in range(1, nrows):
        ax.plot(
            [ax.get_xlim()[0], ax.get_xlim()[1]],
            [x, x],
            lw=1.15,
            color="gray",
            ls=":",
            zorder=3,
            marker="",
        )


def guardar_grafico(fig):
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=const.DPI_FIGURAS)
    b64 = base64.b64encode(buffer.getvalue())
    return b64.decode()


def generar_grafico_5dias(remotos_5dias_agrupado):
    data = remotos_5dias_agrupado.copy()
    data = data.sort_values(by=["Fecha"])
    fig, ax = plt.subplots(figsize=(10, 5))

    colors = ["#1f77b4" for _ in range(len(data["Fecha"]))]
    bars = ax.bar(
        x=data["Fecha"],
        height=data["CantidadRemotos"],
        tick_label=data["Fecha"],
        color="#1f77b4",
        width=0.5,
    )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#DDDDDD")
    ax.tick_params(bottom=False, left=False)
    ax.set_axisbelow(True)
    ax.yaxis.grid(True, color="#EEEEEE")
    ax.xaxis.grid(False)

    # Obtener los nombres de los días
    dias = pd.to_datetime(data["Fecha"]).dt.day_name(locale="es_CL")

    fechas_nombres = [
        dia + "\n" + data["Fecha"].iloc[idx] for idx, dia in enumerate(dias)
    ]

    ax.set_xticklabels(fechas_nombres, rotation=0, fontsize=8)

    for bar, color in zip(bars, colors):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.1,
            str(int(round(bar.get_height(), 0))),
            horizontalalignment="center",
            color=color,
            weight="bold",
            fontsize=8,
        )

    # Añadir etiquetas y título
    # ax0.set_xlabel('Gerencia', size=8,labelpad=15, color='#333333')
    ax.set_ylabel("Cantidad de Empleados Remotos", size=8, labelpad=15, color="#333333")
    ax.set_title(
        "Empleados Remotos en los Últimos 5 Días",
        size=12,
        color="#1f77b4",
        weight="bold",
    )
    # Alinear las etiquetas del eje x a la derecha
    plt.setp(ax.xaxis.get_majorticklabels(), ha="center")

    fig.tight_layout()

    # Crear un objeto BytesIO para almacenar el gráfico en memoria
    buffer = io.BytesIO()
    # Guardar el gráfico en el buffer con el formato y dpi deseados
    fig.savefig(buffer, format="png", dpi=const.DPI_FIGURAS)
    # Codificar el buffer en base64
    b64 = base64.b64encode(buffer.getvalue())

    # Retornar el gráfico en base64
    return b64.decode()


def calcular_remotos_5dias(rut, remotos_5dias):
    data = remotos_5dias.copy()
    data = data[data.Rut.astype(str) == str(rut)]
    return len(data)


def generar_tabla_remotos_hoy(remotos, remotos_5dias):
    data = remotos.copy()
    data[const.KEY_REMOTO_5DIAS] = data.apply(
        lambda row: calcular_remotos_5dias(row["Rut"], remotos_5dias), axis=1
    )
    data = data.sort_values(by=[const.KEY_REMOTO_5DIAS])

    fig, ax = plt.subplots(figsize=(10, 10))

    ncols = data.shape[1]
    nrows = data.shape[0]

    ax.set_xlim(0, ncols + 1)
    ax.set_ylim(0, nrows + 1)

    positions = [0.4, 1.55, 3, 4.5, 5.65]
    columns = data.columns

    # Add table's main text
    for i in range(nrows):
        for j, column in enumerate(columns):
            font_color = "#4f4e4b"
            font_size = 8
            ha = "center"
            text_label = f"{data[column].iloc[i]}"
            if column == const.KEY_REMOTO_5DIAS:
                weight = "bold"
            elif column == "Rut":
                weight = "bold"
            else:
                weight = "normal"
            ax.annotate(
                xy=(positions[j], i + 0.5),
                text=text_label,
                ha=ha,
                va="center",
                weight=weight,
                color=font_color,
                size=font_size,
            )

    # Add column names
    column_names = data.columns.to_list()
    for index, c in enumerate(column_names):
        font_size = 10
        ha = "center"
        ax.annotate(
            xy=(positions[index], nrows + 0.25),
            text=column_names[index],
            ha=ha,
            va="bottom",
            weight="bold",
            color="#1f77b4",
            size=font_size,
        )

    # Add dividing lines
    ax.plot(
        [ax.get_xlim()[0], ax.get_xlim()[1]],
        [nrows, nrows],
        lw=1.5,
        color="#96948d",
        marker="",
        zorder=4,
    )
    ax.plot(
        [ax.get_xlim()[0], ax.get_xlim()[1]],
        [0, 0],
        lw=1.5,
        color="#96948d",
        marker="",
        zorder=4,
    )
    for x in range(1, nrows):
        ax.plot(
            [ax.get_xlim()[0], ax.get_xlim()[1]],
            [x, x],
            lw=1.15,
            color="gray",
            ls=":",
            zorder=3,
            marker="",
        )

    ax.set_axis_off()

    fig.tight_layout()

    # Crear un objeto BytesIO para almacenar el gráfico en memoria
    buffer = io.BytesIO()
    # Guardar el gráfico en el buffer con el formato y dpi deseados
    fig.savefig(buffer, format="png", dpi=const.DPI_FIGURAS)
    # Codificar el buffer en base64
    b64 = base64.b64encode(buffer.getvalue())
    # Retornar el gráfico en base64
    return b64.decode()


def generar_correo(remotos, remotos_agrupado, remotos_5dias, remotos_5dias_agrupado):

    grafico_remotos_area = generar_grafico_remotos_area(remotos_agrupado)
    detalle_remotos_area = generar_tabla_remotos_area(remotos_agrupado)
    grafico_5dias = generar_grafico_5dias(remotos_5dias_agrupado)
    detalle_remotos_hoy = generar_tabla_remotos_hoy(remotos, remotos_5dias)

    mail_body = f"""

    <p> Estimados(as),</p>

    <p>Les comparto el reporte generado con la información de los empleados que trabajaron de forma remota el día de hoy. Adjunto a este 
    correo encontrarán un archivo de Excel que contiene una lista detallada de todos los empleados que se conectaron de forma remota. 
    En este archivo, podrán ver información como el nombre del empleado y su área.</p>

    <p>Asimismo, esta información la encontrarán en las tablas que se encuentran al final de este correo, además de algunos gráficos que resumen esta información. </p>

    <p>Saludos,</p>
    

    <img src="data:image/png;base64,{grafico_remotos_area}" style="width: 80%; height: auto;" />

    <img src="data:image/png;base64,{detalle_remotos_area}" style="width: 80%; height: auto;" />

    <img src="data:image/png;base64,{grafico_5dias}" style="width: 80%; height: auto;" />
    
    <img src="data:image/png;base64,{detalle_remotos_hoy}" style="width: 80%; height: auto;" />




    <br>
    <br>
    <i>Este mensaje fue enviado de manera automática. Si tiene algún comentario, duda o sugerencia favor enviar correo a lfiguero@banchile.cl o lsalazar@banchile.cl.</i>


    """
    return mail_body


def generar_excel(dict_df, nombre_archivo):
    # Crear un escritor de Excel usando pandas
    with pd.ExcelWriter(nombre_archivo, engine="xlsxwriter") as writer:
        # Obtener el objeto workbook de xlsxwriter
        workbook = writer.book

        # Crear un formato para las celdas de encabezado
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "fg_color": "#1f77b4",
                "font_color": "white",
                "border": 1,
            }
        )

        # Crear formatos para las celdas de datos
        cell_format1 = workbook.add_format({"bg_color": "white", "border": 0})
        cell_format2 = workbook.add_format({"bg_color": "#d1e3f0", "border": 0})

        # Iterar sobre el diccionario de DataFrames
        for sheetname, df in dict_df.items():
            # Escribir el DataFrame a su hoja correspondiente en el archivo de Excel
            df.to_excel(writer, sheet_name=sheetname, index=False)

            worksheet = writer.sheets[sheetname]
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            for row in range(1, len(df) + 1):
                for col in range(len(df.columns)):
                    if row % 2 == 0:
                        worksheet.write(row, col, df.iloc[row - 1, col], cell_format1)
                    else:
                        worksheet.write(row, col, df.iloc[row - 1, col], cell_format2)
            # Ajustar el ancho de las columnas
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            worksheet.set_column(0, len(df.columns) - 1, width=15)

            for column in df:
                column_length = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                worksheet.set_column(col_idx, col_idx, column_length)


def enviar_reporte(remotos, remotos_agrupado, remotos_5dias, remotos_5dias_agrupado):

    ruta_excel = os.path.join(
        os.getcwd(), "files", f"usuarios-remotos-{const.FECHA_HOY_STRING_GUIONES}.xlsx"
    )

    # Supongamos que ya tienes tus DataFrames: remotos, remotos_agrupado, remotos_5dias, remotos_5dias_agrupado

    # Crear un diccionario donde las claves son los nombres de las hojas y los valores son los DataFrames correspondientes
    tablas_excel = {
        "Usuarios Remotos Hoy": remotos,
        "Remotos por Gerencia": remotos_agrupado,
        "Detalle 5 Dias": remotos_5dias,
        "Remotos 5 Dias": remotos_5dias_agrupado,
    }

    try:
        mail_body = generar_correo(
            remotos, remotos_agrupado, remotos_5dias, remotos_5dias_agrupado
        )

        generar_excel(tablas_excel, ruta_excel)

        mail_status = enviarCorreo(
            destinatarios=const.DESTINATARIOS_REPORTE,
            conCopia=const.COPIA_REPORTE,
            asunto=f"Banchile - Reporte de Usuarios Remotos - {const.FECHA_HOY_STRING}",
            mensaje=mail_body,
            adjuntarArchivo=ruta_excel,
            mostrarSinEnviar=False,
        )
        if mail_status:
            logger.info("Reporte enviado por correo exitosamente")
        else:
            raise ValueError(
                f"Ocurrio un error al enviar el correo - {const.FECHA_HOY_STRING}"
            )
    except Exception as e:
        raise ValueError(
            f"Ocurrio un error al enviar el correo - {const.FECHA_HOY_STRING} :{e}"
        )
