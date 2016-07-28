# Complemento de censura de JavaScript para Word

Obtenga información sobre cómo crear un complemento que busca, resalta y censura texto.    

## Tabla de contenido
* [Historial de cambios](#change-history)
* [Nota importante](#important-note)
* [Requisitos previos](#prerequisites)
* [Configurar el proyecto](#configure-the-project)
* [Ejecutar el proyecto](#run-the-project)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## Historial de cambios

25 de julio de 2016:
* Versión de ejemplo inicial.

##Nota importante

En este ejemplo se censura texto al reemplazar cada letra con un asterisco.  El diseño no se conserva tal cual.  No se realiza ninguna medida ni cálculo para determinar el tipo de letra, el tamaño ni el formato de cada letra.

## Requisitos previos

* Word 2016 para Windows, compilación 16.0.6912.1000 o posterior.
* [Nodo y npm](https://nodejs.org/en/)
* [GIT Bash](https://git-scm.com/downloads): debe usar una compilación posterior, ya que las compilaciones anteriores pueden mostrar un error al generar los certificados.

## Configurar el proyecto

Ejecute los siguientes comandos desde el shell de Bash en la raíz de este proyecto:

1. Clone este repositorio en el equipo local.
2. ```npm install``` para instalar todas las dependencias en package.json.
3. ```bash gen-cert.sh``` para crear los certificados necesarios para ejecutar este ejemplo. 
* Después, en el repositorio en el equipo local, haga doble clic en ca.crt y seleccione **Instalar certificado**. 
* Seleccione **Máquina local** y **Siguiente** para continuar. 
* Seleccione **Colocar todos los certificados en el siguiente almacén** y **Examinar**.  
* Seleccione **Entidades de certificación raíz de confianza** y **Aceptar**. 
* Seleccione **Siguiente** y **Finalizar**. Ahora, se ha agregado el certificado de la autoridad de certificación al almacén de certificados.
4. ```npm start``` para iniciar el servicio de servidor web del nodo.

Ahora debe indicarle a Microsoft Word dónde encontrar el complemento.

1. Cree un recurso compartido de red o [comparta una carpeta en la red](https://technet.microsoft.com/es-es/library/cc770880.aspx) y coloque el archivo de manifiesto [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) en él.
3. Inicie Word y abra un documento.
4. Seleccione la pestaña **Archivo** y haga clic en **Opciones**.
5. Haga clic en **Centro de confianza** y seleccione el botón **Configuración del Centro de confianza**.
6. Seleccione **Catálogos de complementos de confianza**.
7. En el campo **Dirección URL del catálogo**, escriba la ruta de red al recurso compartido de carpeta que contiene word-add-in-js-redact-manifest.xml y después elija **Agregar catálogo**.
8. Active la casilla **Mostrar en menú** y haga clic en **Aceptar**.
9. Aparecerá un mensaje para informarle de que la configuración se aplicará la próxima vez que inicie Microsoft Office. Cierre y vuelva a iniciar Word.

## Ejecutar el proyecto

1. Abra un documento de Word.
2. En la pestaña **Insertar** de Word 2016, elija **Mis complementos**.
3. Seleccione la pestaña **Carpeta compartida**.
4. Elija **Word Redact add-in** (Complemento de censura de Word) y seleccione **Aceptar**.
5. Si su versión de Word admite los comandos de complemento, la interfaz de usuario le informará de que se ha cargado el complemento.

### Interfaz de usuario de la cinta de opciones

En la cinta de opciones:
* Seleccione la ficha **Revisión** y elija **Mostrar el panel de tareas de censura** para iniciar el panel de tareas.

 > Nota: El complemento se cargará en un panel de tareas si los comandos del complemento no son compatibles con su versión de Word.

### Interfaz de usuario del panel de tareas

En el panel de tareas, puede:
* Busque y resalte texto en un documento escribiendo la palabra en el cuadro de texto y seleccionando el botón **Buscar y resaltar**.
  
> NOTA:  La búsqueda distingue mayúsculas de minúsculas.  Para deshacer una acción, pulse Ctrl + Z en el teclado.

* Censure texto en un documento escribiendo palabras en el cuadro de texto y seleccionando el botón **Censurar**.
  
> NOTA:  La censura distingue mayúsculas de minúsculas.   

* Censure el texto seleccionado en un documento seleccionando el botón **Censurar el texto seleccionado**.
  
> NOTA:  La censura distingue mayúsculas de minúsculas.       
  
### Comandos de complementos

En este ejemplo se muestra cómo crear comandos de complemento en la cinta de opciones. Las declaraciones del comando del complemento se encuentran en word-add-in-js-redact-manifest.xml. 

## Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre el ejemplo de censura de Word. Puede enviarnos sus comentarios a través de la sección *Problemas* de este repositorio.

Las preguntas generales sobre desarrollo en Microsoft Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Asegúrese de que sus preguntas se etiquetan con [office-js] y [API].

## Recursos adicionales

* [Documentación de complemento de Office](https://msdn.microsoft.com/es-es/library/office/jj220060.aspx)
* [Centro para desarrolladores de Office](http://dev.office.com/)
* [Proyectos de inicio de las API de Office 365 y ejemplos de código](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Copyright
Copyright (c) 2016 Microsoft Corporation. Todos los derechos reservados.


