"""
-------------------------------------------------------------------------------
_Titulo: reporte-horas-trabajadas-remoto
_Lugar: Banchile - Gerencia de Riesgo
_@autor: LFIGUERO
_Fecha de creacion: 2024-04-23
_Ultimos cambios: 
-------------------------------------------------------------------------------
"""

import logging
import os
import colorlog
from FuncionesGenerales import (
    detenerSiNoEsDiaHabil,
    conectar2SQL,
    registrarInicioProceso,
    registrarFinProceso,
)
from utils import constantes
from application.funciones import test

# -----------------------------------------------------------------------------
# Logging setup: Se definen las configuraciones basicas para el logging.
# -----------------------------------------------------------------------------
logger = logging.getLogger("reporte-horas-trabajadas-remoto")
logger.setLevel(logging.DEBUG)

# Handler para escribir logs en un archivo
f_handler = logging.FileHandler(
    os.path.join(os.getcwd(), "reporte-horas-trabajadas-remoto.log")
)
f_format = logging.Formatter(
    "%(asctime)s [%(name)s][%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
)
f_handler.setFormatter(f_format)
logger.addHandler(f_handler)

# Handler para mostrar logs en la consola
log_colors = {
    "DEBUG": "cyan",
    "INFO": "green",
    "WARNING": "yellow",
    "ERROR": "red",
    "CRITICAL": "red,bg_white",
}

color_formatter = colorlog.ColoredFormatter(
    constantes.COLOR_FORMAT, datefmt="%Y-%m-%d %H:%M:%S", log_colors=log_colors
)

console_handler = logging.StreamHandler()
console_handler.setFormatter(color_formatter)
logger.addHandler(console_handler)

# -----------------------------------------------------------------------------


dbriesgo, cursor = conectar2SQL()

os.chdir(constantes.WORKSPACE_DIR)


def main():
    """
    Ejecucion del proceso principal.
    """
    logger.info("-" * 10)
    logger.info("INICIANDO PROCESO")
    logger.info("-" * 10)

    detenerSiNoEsDiaHabil(dbriesgo)

    # registrarInicioProceso(constantes.PROCESS_NAME, "NOMBRE-BBDD", cursor)

    test()

    # registrarFinProceso(constantes.PROCESS_NAME, "NOMBRE-BBDD", cursor)

    logger.info("-" * 10)
    logger.info("FINALIZANDO PROCESO")
    logger.info("-" * 10)


if __name__ == "__main__":

    try:
        main()
        # logger.info('Proceso en Mantenimiento')
    except Exception as e:
        logger.exception("Ocurrio un error en el proceso %s", str(e))
        raise e

    finally:
        logger.handlers.clear()
