<p align="center">
    <img src="http://dev.comunidadargentum.com/ao_logo.png" />
</p>


<h1 align="center">
    <span>Argentum Online GameServer - VB6 </span>
</h1>
<p align="center">
    <span>También disponible en los siguientes idiomas</span>
</p>
<p align="center">
    <a href="#">Spanish</a> - <a href="#">English</a> - <a href="#">Portuguese</a>
</p>


---

[Argentum Online](https://www.argentumonline.com.ar) es un juego de rol, multijugador y masivo open source creado en Visual Basic 6.  
Este repositorio contiene el código y herramientas necesarias para compilar y ejecutar el SERVIDOR de Argentum Online (VB6).

El branch [`master`](../../../tree//master) contiene el código exacto de la última versión implementada en el [Servidor Oficial](https://www.argentumonline.com.ar).  
El branch [`development`](../../../tree/development) contiene las features, correcciones de errores y cambios que están planeados en el [milestone actual](/milestones).

## ➥ Componentes requeridos ##
El servidor, para su funcionamiento, requiere tener instalados los siguientes paquetes:

* [Visual Basic 6 SP6](https://www.microsoft.com/en-us/download/details.aspx?id=5721)
    * Librerías de tiempo de ejecución de VB6 SP6
	* Para instalarlo, descargar y descomprimir el archivo, y luego ejecutar el archivo **setupsp6.exe**
* [Visual Basic 6 SP6 Cumulative Updates](http://www.microsoft.com/downloads/details.aspx?familyid=cb824e35-0403-45c4-9e41-459f0eb89e36&displaylang=en)
    * Parche acumulativo para VB6 SP6 con upgrades a varias librerías requeridas por el proyecto.
* [MariaDB >=10.x](https://downloads.mariadb.org/)
    * Base de datos para el servidor
* [ODBC Connector x86 >= 8.0.0 ](https://dev.mysql.com/downloads/connector/odbc/)
    * Conector VB6->Base de datos

## ➥ Setup inicial ##
* Copiar las carpetas de recursos (Maps y Dats) en la carpeta del servidor. Los recursos se pueden descargar desde los repositorios de ambientación correspondientes
  * DAT: https://github.com/argentumonline/resources.dats
  * Maps: https://github.com/argentumonline/resources.maps
* Crear una base de datos MySQL e importar los archivos `.sql` que se encuentran en la carpeta `SQL` en orden
  * `01-dump_structure.sql` - Estructura de la base de datos
  * `02-dump_procedures.sql` - Stored procedures
  * `03-dump_data.sql` - Datos básicos necesarios para el funcionamiento del servidor.
* Crear una copia del archivo `server.ini.example` y nombrarlo `server.ini`
* Registrar la librería Aurora_IO.dll ejecutando el comando `regsvr32 Aurora_IO.dll`
* Modificar la sección `[DATABASE]` en el archivo `server.ini` tal y como se explica a continuación.

```
[DATABASE]
Driver=MySQL ODBC 8.0 Unicode Driver    # Nombre del driver ODBC Connector instalado
Server=127.0.0.1                        # IP del servidor de base de datos.
Database=ao-db                          # Nombre de la base de datos
UID=root                                # Usuario de la base de datos
Password=password                       # Password del usuario de la base de datos
```


## ➥ Arquitectura del servidor
<p align="center">
    <img src="http://dev.comunidadargentum.com/hlo_server_architecture.png" />
</p>

Debido a la naturaleza del lenguaje utilizado para el GameServer, el mismo requiere de otros componentes para realizar tareas pesadas y asi evitar el procesamiento excesivo que puede afectar negativamente la velocidad en la que se procesan los paquetes.  
Estos componentes están descritos en la imagen anterior, y la documentación para cada uno de ellos puede encontrarse en sus respectivos repositorios
* StateServer (Servidor de estados) - [Repositorio](https://github.com/argentumonline/state-server)
    * Encargado de controlar eventos basados en tiempo (Ej. la duración de las fogatas, las "condiciones sobre tiempo" o CoT, etc).
* Message Queue Proxy - [Repositorio](https://github.com/argentumonline/tool-queue-proxy)
    * Encargado de encolar mensajes en RabbitMQ para que el/los consumers de cada tipo de mensaje puedan procesarlos.
* MQConsumer - [Repositorio](https://github.com/argentumonline/queue-consumer)
    * Actualmente es el único consumer, encargado de leer los mensajes encolados en RabbitMQ y enviarlos a la web para iniciar el proceso de creación de cuentas. 


## ➥ Entorno de desarrollo basado en Docker
Este repositorio contiene un archivo `docker-compose.yml` que nos permite crear los servicios necesarios para el funcionamiento y administración del servidor durante su desarrollo.
Es necesario tener instalado [Docker Desktop](https://www.docker.com/products/docker-desktop).

Los servicios que serán creados mediante este procedimiento son:
* `MariaDB 10.x` - Servidor debase de datos
* `RabbitMQ` - Gestor de colas para procesar eventos de forma asincrónica.
* `Adminer` - Administrador de bases de datos web en caso de necesitar inspeccionar la base de datos

## ➥ Levantando el entorno usando docker-compose 

Antes de ejecutar el comando docker-compose, tenemos que asegurarnos que:
* Docker está funcionando y corriendo como administrador
* Está seleccionado el uso de Linux Containers
* La unidad (disco) donde se encuentra el servidor de Argentum Online está correctamente compartida con Docker ([ver el siguiente enlace](https://token2shell.com/howto/docker/sharing-windows-folders-with-containers/)).  


Para levantar el entorno de desarrollo, se deberá abrir una consola y posicionarse en el directorio del repositorio, y luego ejecutar el siguiente comando 

```
docker-compose up -d
```
Este comando descargará imagenes de docker para cada uno de los servicios mencionados en [Entorno de desarrollo basado en docker](#entorno-de-desarrollo-basado-en-docker), e iniciará un container basado en las configuraciones del archivo [docker-compose.yml](docker-compose.yml). 
El contenedor de la base de datos, por defecto, importará los 3 archivos SQL mencionados en la sección `Setup Inicial` y la base de datos estaría lista para ser usada.

## ➥ Cómo contribuir
Por favor, sigue nuestra [guía de contribución](https://github.com/argentumonline/.github/blob/master/.github/CONTRIBUTING.md) para saber cómo contribuir al código de Argentum Online.  


---
<p align="center">
    <a href="https://www.argentumonline.com.ar">Argentum Online</a> - <a href="https://wiki.comunidadargentum.com">Manual del juego</a> - <a href="https://soporte.comunidadargentum.com">Soporte</a>
</p>