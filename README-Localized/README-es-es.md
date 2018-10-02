# <a name="word--javascript-redact-add-in"></a>Complemento de censura de JavaScript para Word

Obtenga información sobre cómo crear un complemento que busca, resalta y censura texto.    

## <a name="table-of-contents"></a>Tabla de contenido
* [Historial de cambios](#change-history)
* [Nota importante](#important-note)
* [Requisitos previos](#prerequisites)
* [Configurar el proyecto](#configure-the-project)
* [Ejecutar el proyecto](#run-the-project)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## <a name="change-history"></a>Historial de cambios

25 de julio de 2016:
* Versión de ejemplo inicial.

## <a name="important-note"></a>Nota importante

En este ejemplo se censura texto al reemplazar cada letra con un asterisco.  El diseño no se conserva tal cual.  No se realiza ninguna medida ni cálculo para determinar el tipo de letra, el tamaño ni el formato de cada letra.

## <a name="prerequisites"></a>Requisitos previos

* Word 2016 para Windows, compilación 16.0.6912.1000 o posterior.
* [Nodo y npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads): debe usar una compilación posterior, ya que las compilaciones anteriores pueden mostrar un error al generar los certificados.

## <a name="configure-the-project"></a>Configurar el proyecto

Ejecute los siguientes comandos desde el shell de Bash en la raíz de este proyecto:

1. Clone este repositorio en el equipo local.
2. ```npm install``` para instalar todas las dependencias en package.json.
3. ```bash gen-cert.sh``` para crear los certificados necesarios para ejecutar este ejemplo. 
* Después, en el repositorio del equipo local, haga doble clic en ca.crt y seleccione **Instalar certificado**. 
* Seleccione **Máquina local** y **Siguiente** para continuar. 
* Seleccione **Colocar todos los certificados en el siguiente almacén** y **Examinar**.  
* Seleccione **Entidades de certificación raíz de confianza** y **Aceptar**. 
* Seleccione **Siguiente** y **Finalizar**. Ahora, se ha agregado el certificado de la autoridad de certificación al almacén de certificados.
4. ```npm start``` para iniciar el servicio de servidor web del nodo.

Ahora debe indicarle a Microsoft Word dónde encontrar el complemento.

1. Cree un recurso compartido de red o [comparta una carpeta en la red](https://technet.microsoft.com/en-us/library/cc770880.aspx) y coloque allí el archivo de manifiesto [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml).
3. Inicie Word y abra un documento.
4. Seleccione la pestaña **Archivo** y haga clic en **Opciones**.
5. Haga clic en **Centro de confianza** y seleccione el botón **Configuración del Centro de confianza**.
6. Seleccione **Catálogos de complementos de confianza**.
7. En el campo **Dirección URL del catálogo**, escriba la ruta de red al recurso compartido de carpeta que contiene word-add-in-js-redact-manifest.xml y, después, elija **Agregar catálogo**.
8. Seleccione la casilla **Mostrar en menú** y, luego, elija **Aceptar**.
9. Aparecerá un mensaje para informarle de que la configuración se aplicará la próxima vez que inicie Microsoft Office. Cierre y vuelva a iniciar Word.

## <a name="run-the-project"></a>Ejecutar el proyecto

1. Abra un documento de Word.
2. En la pestaña **Insertar** de Word 2016, elija **Mis complementos**.
3. Seleccione la pestaña **CARPETA COMPARTIDA**.
4. Elija **el complemento Redact de Word** y seleccione **Aceptar**.
5. Si su versión de Word admite los comandos de complemento, la interfaz de usuario le informará de que se ha cargado el complemento.

### <a name="ribbon-ui"></a>Interfaz de usuario de la cinta de opciones

En la cinta de opciones:
* Seleccione la pestaña **Revisar** y elija **Mostrar el panel de tareas de censura** para iniciar el panel de tareas.

 > Nota: El complemento se cargará en un panel de tareas si los comandos del complemento no son compatibles con su versión de Word.

### <a name="task-pane-ui"></a>Interfaz de usuario del panel de tareas

En el panel de tareas, puede:
* Busque y resalte texto en un documento escribiendo la palabra en el cuadro de texto y seleccionando el botón **Buscar y resaltar**.
  
> NOTA:  La búsqueda distingue mayúsculas de minúsculas.  Para deshacer una acción, pulse Ctrl + Z en el teclado.

* Censure texto en un documento escribiendo palabras en el cuadro de texto y seleccionando el botón **Censurar**.
  
> NOTA:  La censura distingue mayúsculas de minúsculas.   

* Censure el texto seleccionado en un documento seleccionando el botón **Censurar el texto seleccionado**.
  
> NOTA:  La censura distingue mayúsculas de minúsculas.       
  
### <a name="add-in-commands"></a>Comandos de complemento

En este ejemplo se muestra cómo crear comandos de complemento en la cinta de opciones. Las declaraciones del comando del complemento se encuentran en word-add-in-js-redact-manifest.xml. 

## <a name="questions-and-comments"></a>Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre el ejemplo de censura de Word. Puede enviarnos sus comentarios a través de la sección *Problemas* de este repositorio.

Las preguntas generales sobre desarrollo en Microsoft Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Asegúrese de que sus preguntas se etiquetan con [office-js] y [API].

## <a name="additional-resources"></a>Recursos adicionales

* 

  [Documentación de complementos de Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centro de desarrollo de Office](http://dev.office.com/)
* [Proyectos de inicio y ejemplos de código de las API de Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Todos los derechos reservados.



Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, consulte las [preguntas más frecuentes sobre el Código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
