# pyexcel

## Instalación:
1) Clonar repositorio: `git clone https://github.com/DeusCL/pyexcel.git`
2) Ir al directorio del proyecto: `cd pyexcel`
3) Instalar dependencias: `pip install -r requirements.txt`

### Dependencias:
- pandas
- pywin32
- mysql-connector-python

## Requisitos:
Antes de ejecutar main.py se requiere tener una base de datos llamada `test_altos_ejecutivos`. La estructura de la base de datos debe ser creada utilizando el script SQL proporcionado:
```bach
mysql -u <username> -p < create_database.sql
```
## Uso:
- El archivo excel que se desea leer debe estar en el current working directory de `main.py`.
- Teniendo el archivo excel preparado, ejecuta el script: `main.py`.
- Al ejecutar el script, se pedirá por teclado el archivo excel que se desea leer. Deja este campo vacío para seleccionar el primero que aparece en la lista.
- Si el archivo excel está protegido por contraseña, el programa pedirá ingresar la contraseña por teclado.
- Luego, el programa solicitará credenciales de usuario MySQL para conectarse con la base de datos `test_altos_ejecutivos`.
- Una vez establecida la conexión, el programa realizará las inserciones de datos correspondiente a la base de datos.

