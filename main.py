from win32com.client import Dispatch
import pandas as pd
import pywintypes
import getpass
import warnings
import os

import mysql.connector
from mysql.connector import errorcode


AC, AH, D, I, N, P = 'RF_SB2', 'RF_400', 'C_100', 'C_GR', 'P_G', 'P_NA'
SHEETNAME = "ALTOS EJECUTIVOS"
DATABASE_NAME = "test_altos_ejecutivos"
TEMPFILE_PREFIX = "temp_"


def remove_xlsx_password(filename, password):
    """ Crea una copia sin clave del archivo excel original
    
    Args:
        filename (str): Nombre del archivo excel
        password (str): Contraseña del archivo excel
    
    Return:
        newfilepath (str): Ruta del archivo excel generado sin contraseña.
    """

    cwd = os.getcwd()
    filepath = os.path.join(cwd, filename)
    newfilepath = os.path.join(cwd, f"{TEMPFILE_PREFIX}{filename}")

    xcl = Dispatch("Excel.Application")
    wb = xcl.Workbooks.Open(filepath, False, False, None, password)
    xcl.DisplayAlerts = False

    try:
        wb.SaveAs(newfilepath, None, '', '')
    except KeyboardInterrupt:
        xcl.Quit()
        raise KeyboardInterrupt

    xcl.Quit()

    return newfilepath



def get_df_from_secured_xlsx(password, *args, **kwargs):
    """ Crea una copia temporal sin clave del archivo excel original,
    extrae los datos y luego elimina el archivo temporal. Y retorna
    los datos obtenidos en un DataFrame.

    Args:
        password (str): La clave del archivo xlsx protegido
        *args, **kwargs: Argumentos que irían en pd.read_excel(...)
    Return:
        Pandas DataFrame.
    """

    newfilepath = remove_xlsx_password(args[0], password)

    # Desactivar las advertencias temporalmente
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning)
        df = pd.read_excel(newfilepath, *args[1:], **kwargs)

    os.remove(newfilepath)

    return df



def remove_temp_xlsx(filename):
    """ Elimina el archivo xlsx temporal que deja el método
    remove_xlsx_password cuando el proceso es interrumpido. """

    cwd = os.getcwd()
    tempfilepath = os.path.join(cwd, f"{TEMPFILE_PREFIX}{filename}")

    if os.path.exists(tempfilepath):
        os.remove(tempfilepath)



def search_for_an_excel():
    """ Busca archivos excel en el current working directory y retorna
    el nombre del archivo escogido por el usuario. """

    files = os.listdir('.')

    excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]

    if not excel_files:
        raise Exception("No se encontraron archivos de Excel en el directorio actual.")

    # Enlistar los archivos de Excel disponibles
    print("Archivos Excel:")
    for i, file in enumerate(excel_files):
        print(f"\t[{i}]: {file}")

    while True:
        choice = input("\nNúmero del archivo a leer [0]: ")
        choice = "0" if choice == "" else choice

        try:
            choice = int(choice)
            if 0 <= choice < len(excel_files):
                return excel_files[choice]
            else:
                print("Por favor, ingrese un número válido.")
        except ValueError:
            print("Por favor, ingrese un número válido.")




class DBManager:
    def __init__(self, app):
        self.app = app

        self.conn = None
        self.cursor = None


    def connect_to_database(self):
        print("\n\nConectar con la base de datos MySQL:")
        host = input("Host [localhost]: ")
        user = input("User [root]: ")
        password = getpass.getpass()

        self.conn = mysql.connector.connect(
            host='localhost' if host == '' else host,
            user='root' if user == '' else user,
            password=password,
            database=DATABASE_NAME
        )

        self.cursor = self.conn.cursor()


    def close(self):
        if self.conn is not None:
            print("Closing connection to database.")
            self.conn.close()



class App:
    def __init__(self):
        self.dbm = DBManager(self)
        self.df = None
        self.header_row = 5
        self.start_row = 2


    def calc_renta_bruta(self):
        # Sumar todo desde AC a AH, asumiendo que RF_SB2 y RF_400 siempre
        # corresponderán a los mismos valores.
        # 'RF_SB2' = Sueldo Base Mensual (columna AC)
        # 'RF_400' = Asignación Movilización Mensual (columna AH)

        print("Calculando Renta Bruta...")

        return self.df.loc[self.start_row:, AC:AH].sum(axis=1)


    def get_cargos_data(self):
        """ Retorna una lista de tuplas que contienen los datos de las columnas 
        D, I, N y P, para ser agregados a la tabla de Cargos. """

        data = self.df.iloc[self.start_row:].dropna(subset=[D, I, N, P]).apply(
            lambda row: (row[D], row[I], row[N], row[P]), axis=1
        ).tolist()[:-1]

        return data


    def get_rentas_data(self, cargos_data):
        """ Retorna una lista que contiene (id, renta_bruta) para ser agregados
        a la tabla de Rentas. """

        from_to = slice(self.start_row, len(cargos_data) + self.start_row)
        data = list(enumerate(
            self.df['Renta Bruta'][from_to], start=1
        ))

        return data




    def try_to_open_xlsx(self, filename):
        try:
            self.df = pd.read_excel(filename,
                sheet_name=SHEETNAME, skiprows=self.header_row-1
            )
        except Exception as e:

            # Si hay OLE2 en el error podría significar que el archivo
            # está protegido por contraseña.
            if "OLE2" in str(e):

                print("\nContraseña del archivo excel: ")
                excel_password = getpass.getpass()

                self.df = get_df_from_secured_xlsx(excel_password,
                    filename, sheet_name=SHEETNAME, skiprows=self.header_row-1
                )



    def start(self):
        self.excel_filename = search_for_an_excel()
        self.try_to_open_xlsx(self.excel_filename)

        self.dbm.connect_to_database()

        conn = self.dbm.conn
        cursor = self.dbm.cursor


        self.df['Renta Bruta'] = self.calc_renta_bruta()

        cargos_data = self.get_cargos_data()
        rentas_data = self.get_rentas_data(cargos_data)

        # Insertar datos en la tabla Cargos
        print("Insertando datos en la tabla Cargos...")
        cursor.executemany("""
            INSERT INTO Cargos (nombre, grado, genero, nacionalidad)
            VALUES (%s, %s, %s, %s)
        """, cargos_data)

        # Insertar datos en la tabla Rentas
        print("Insertando datos en la tabla Rentas...")
        cursor.executemany("""
            INSERT INTO Rentas (cargo_id, renta_bruta)
            VALUES (%s, %s)
        """, rentas_data)

        conn.commit()
        conn.close()

        print("Proceso terminado exitosamente.")


    def close(self):
        self.dbm.close()
        remove_temp_xlsx(self.excel_filename)



if __name__ == '__main__':
    app = App()

    try:
        app.start()

    except pywintypes.com_error as err:
        print(f"Error abriendo archivo excel: {err.args[2][2]}")

    except mysql.connector.errors.ProgrammingError as err:
        print("mysql.connector error:", end=' ')

        if err.errno == errorcode.ER_BAD_DB_ERROR:
            print(f"Aún no existe la base de datos '{DATABASE}'.", end=' ')
            print("Primero debe ejecutar el archivo 'create_database.sql'.")

        else:
            print("Ha ocurrido un error.")

        print(f"Detalle del error: {err}")

    except mysql.connector.errors.DatabaseError as err:
        print("mysql.connector error:", end=' ')

        if err.errno == errorcode.CR_UNKNOWN_HOST:
            print("No se pudo reconocer el host proporcionado.")
        else:
            print("Ha ocurrido un error.")

        print(f"Detalle del error: {err}")

    except Exception as err:
        print(f"Error: {err}")

    except KeyboardInterrupt:
        print("\nSaliendo del sistema.")
        app.close()




