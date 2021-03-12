# Generador de Certificados para Cursos

Este proyecto es un generador / enviador automático de certificados para cursos.

## Instalación

Este proyecto usa Python 3, cualquier versión de Python `3.X.X` debería funcionar. En el momento se usó `3.8.7`.

Para instalar (estando en el root del proyecto):
```bash
pip install -r requirements.txt
```

En caso de expandir en el futuro, por favor agregar las dependencias nuevas al `requirements.txt` haciendo:
```bash
pip freeze > requirements.txt
```

## Configuración

### Excel

Los datos de las personas para enviar tienen que estar en un archivo `input.xlsx` dentro de la carpeta `input`. Debe tener al menos **email**, **nombre** y **apellido**.

Los nombres de estos campos son configurables dentro de `generator.py` en las variables:
* **FIELD_NAME** --> Nombre
* **FIELD_SURNAME** --> Apellido
* **FIELD_EMAIL** --> Email

### Template

La template debe estar en formato `SVG` en un archivo `template.svg` dentro de la carpeta `templates` (se incluyen ejemplos, conviene subir las templates que usan adentro de ahí para la posteridad).

La template tiene que tener un campo de texto con el texto `%name`, esto es lo que se va a reemplazar por el nombre del participante.

#### ¿Cómo crear una template?

Hay una serie de pasos a seguir para generar una template:
1. Tener un diseño en formato `PNG` con un espacio en blanco para los nombres (ejemplos en la carpeta `templates_png`).
2. Usar la [siguiente página](https://vectr.com/ghirsch/h96EeSoyV) para crear un archivo SVG.
3. Ponerle al archivo el mismo tamaño que el archivo del punto **1** (en píxeles).
4. Importar la imagen del paso **1** y ponerla de fondo.
5. Agregar un cuadro de texto centrado (ponerle el `%name` adentro) y del tamaño final.
6. Exportar como un SVG con la página.
7. Abrir el SVG en un editor y reemplazar lo siguiente:
    * Sacar el contenido del tag `<style>` que define la font, tener eso rompe a la template (experimentar con otras si quieren).
    * Sumarle `100 px` a la posición `y` del texto, el generador corre al texto aproximadamente 100px para arriba, esto es para contrarrestarlo.
8. Ya tenés un template

### Emails

Se encuentra configurado para poder mandar emails automatizados agregando un parámetro (`--send`) a la ejecución del programa.

Para esto es necesario tener una variable de entorno (`EMAIL_PASS`) seteada con la cuenta de Google de `computersociety@itba.edu.ar`.

Para generar la variable en Linux/OSx se hace:
```bash
export EMAIL_PASS="contraseña"
```

## Funcionamiento

Para ejecutar al generador se puede correr:
```bash
python generator.py [--send]
```

La opción `--send` habilita el envío automatizado de mails a los participantes listados en el archivo XLSX. Si no se quiere enviar mails no es necesario usarla.

**Recomendación**: Correr el programa primero sin la opción `--send` para ver que todos los certificados se generen bien. Luego volver a correr pero con la opción para que los envíe.

Author: **Gonzalo Hirsch** --> ghirsch@itba.edu.ar