import CargaDeDatos
import ctypes
import sys


def main():
    # Mayab.cargaDatosMayab()
    # Adelaida.cargaDatosAdelaida()
    # Enfermeria.cargaDatosEnfermeria() ESTOS PUEDEN SER FUNCIONES DE OTROS SCRIPTS
    CargaDeDatos.cargaDatos()


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if is_admin():
    print("Ejecutando con privilegios administrativos.")
    if __name__ == "__main__":
        main()
else:
    # Re-lanza el script como administrador
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, __file__, None, 1)
print("Presiona cualquier tecla para salir...")
