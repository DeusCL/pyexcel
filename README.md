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
- El archivo excel a leer debe estar en el current working directory de main.py
- Al ejecutar el script se pedirá por teclado el archivo excel que se desea leer.
- Si el archivo excel está protegido por contraseña, el programa pedirá ingresar la contraseña por teclado.
- Luego, el programa solicitará credenciales de usuario MySQL para conectarse con la base de datos.
- Una vez que esté la base conectada y todo ingresado correctamente, el programa realizará las inserciones de datos a la base de datos.

