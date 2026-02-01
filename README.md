# Sistema de Gestión Poblacional

Version de Go: 1.13.15
<br> <br>

## Instalación de Go:

### 1. 🛜 Descarga la versión de Go que este proyecto necesita:

Link para Windows 64 bits: https://go.dev/dl/go1.13.15.windows-amd64.msi

Link para Windows 32 bits: https://go.dev/dl/go1.13.15.windows-386.msi (No estoy seguro si es compatible)

### 2. ▶️ Ejecuta el archivo descargado y sigue los pasos del asistente.

### 3. ⚙️ Configura las variables de entorno

* Asegúrate de que el instalador haya agregado Go al `PATH`.
* Puedes verificarlo abriendo una terminal (CMD o PowerShell) y escribiendo:
  `go version`
* Si ves la versión instalada, ¡todo está bien

### 4. 🧑‍💻 Crear tu primer proyecto

* Crea una carpeta de proyecto y dentro un archivo `main.go` con:

  ```
  package main

  import "fmt"

  func main(){
      fmt.Println("¡Hola, Go!")
  }
  ```
* Ejecuta en terminal con:

  ```
  go run main.go
  ```
* Si corre el programa todo esta listo.

<br>

## Cómo agregar librerias al programa

* Abre tu terminal o línea de comandos en la carpeta del proyecto.
* Ejecuta el siguiente comando `go get linkdelalibreria aqui`
* Ejemplo:

```
go get github.com/360EntSecGroup-Skylar/excelize/v2@v2.3.2
```

<br>

## **Librerias utilizadas:**


**Excelize:** Se utiliza para trabajar con archivos de Excel.

`go get github.com/360EntSecGroup-Skylar/excelize/v2@v2.3.2.`

<br>

**wkhtmltopdf:** Para que el go-wkhtmltopdf funcione, wkhtmltopdf debe estar instalada en tu sistema operativo.

* **Ve al sitio web oficial de wkhtmltopdf: [https://wkhtmltopdf.org/downloads.html](https://www.google.com/url?sa=E&q=https%3A%2F%2Fwkhtmltopdf.org%2Fdownloads.html)**

* **Descarga el instalador .exe apropiado para tu versión de Windows (32-bit o 64-bit).**
  
* **Ejecuta el instalador.** Asegúrate de marcar la opción para "Add wkhtmltopdf to PATH" (o similar) durante la instalación. Si no la hay, o si el problema persiste, tendrás que añadir la ruta donde se instalówkhtmltopdf.exe(por ejemplo,C:\Program Files\wkhtmltopdf\bin) a la variable de entorno PATHde forma manual.
  
* **Para añadir a PATH manualmente en Windows:**

  * Busca "Editar las variables de entorno del sistema" en el menú de inicio y ábrelo.
  * Haz clic en "Variables de entorno...".
  * En la sección "Variables del sistema", busca la variable Path y selecciónala.
  * Haz clic en "Editar...".
  * Haz clic en "Nuevo" y añade la ruta a la carpeta bin de wkhtmltopdf (ej: C:\Program Files\wkhtmltopdf\bin).
  * Haz clic en Aceptar en todas las ventanas.
    
* **Después de la instalación:**

  * **Verifica la instalación:** **Abre una nueva terminal y ejecuta el siguiente comando:**

```
  wkhtmltopdf --version
```

<br>

**go-wkhtmltopdf:** Es un wrapper o envoltorio de línea de comandos para la herramienta externa wkhtmltopdf. La cual servira para hacer reportes en PDF.

`go get github.com/SebastiaanKlippert/go-wkhtmltopdf@v1.7.1`
