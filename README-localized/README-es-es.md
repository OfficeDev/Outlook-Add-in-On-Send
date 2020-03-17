---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/20/2017 11:55:13 PM
---
# Código de ejemplo de complemento de envío de Outlook

Descubra cómo buscar palabras restringidas en el cuerpo de un mensaje de correo electrónico de Outlook, agregar un destinatario a la línea CC y comprobar si hay un asunto en el correo electrónico al enviar.

>**Nota:** 

* Actualmente, la característica de envío solo es compatible en Outlook en la Web en Office 365. 
* Para obtener más información sobre la característica de envío, consulte [Característica de envío para complementos de Outlook](https://dev.office.com/docs/add-ins/outlook/outlook-on-send-addins).  
* Para obtener un tutorial de código, consulte [Ejemplos de código](https://docs.microsoft.com/es-es/outlook/add-ins/outlook-on-send-addins#code-examples).

## Tabla de contenido
* [Historial de cambios](#change-history)
* [Requisitos previos](#prerequisites)
* [Configurar e instalar el ejemplo](#configure)
* [Cargar los manifiestos](#manifests)
* [Ejecutar el complemento](#test-the-add-in)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## Historial de cambios

Abril de 2017

* Versión inicial

## Requisitos previos

* Un servidor web de confianza para hospedar los archivos de ejemplo. El servidor debe poder aceptar solicitudes protegidas por SSL (HTTPS) y tener un certificado SSL válido.
* Una cuenta de correo electrónico de Office 365.
* Habilitar la característica de envío: la función de envío está deshabilitada de forma predeterminada. Los complementos para Outlook en la Web que usen la característica de envío se ejecutarán para los usuarios a los que se les asigne una directiva de buzón de Outlook en la Web que tenga la marca **OnSendAddinsEnabled** establecida en **true**. Los administradores pueden habilitarla ejecutando los cmdlets de PowerShell de Exchange Online. Para obtener más información sobre qué cmdlets se deben ejecutar, consulte [Implementar complementos de Outlook que usan la característica de envío](https://docs.microsoft.com/es-es/outlook/add-ins/outlook-on-send-addins#installing-outlook-add-ins-that-use-on-send).

## Configurar e instalar el ejemplo

1. Descargue o bifurque el repositorio.
2. Abra app.js. En la función `addCCOnSend`, cambie `Contoso@contoso.onmicrosoft.com` por su propia dirección de correo electrónico.
2. Cargue los archivos de complemento en un directorio de su servidor web. Los archivos que debe cargar son index.html y app.js.
3. Abra los archivos de manifiesto `Contoso Message Body Checker.xml` y `Contoso Subject and CC Checker.xml` en un editor de texto. Reemplace todas las instancias de `https://localhost:3000` por la URL HTTPS del directorio donde se encuentran los archivos que cargó en el paso anterior. Guarde los cambios.

   >  Para obtener más información acerca de:
   * cómo ejecutar complementos de Outllook, consulte [Ejecutar un complemento de Outlook en una cuenta de Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
   * los manifiestos, consulte [Manifiestos de complementos de Outlook](https://dev.office.com/docs/add-ins/outlook/manifests/manifests)

## Cargar los manifiestos

1. Inicie sesión en [Outlook Web App](https://outlook.office365.com).
2. Haga clic en **Configuración** (el engranaje en la esquina superior derecha de la página) para abrir la página **Configuración** (como se muestra en la siguiente captura de pantalla).

  ![Página de configuración](./readme-images/block-on-send-settings.png)

3. En la sección **Configuración de la aplicación** de la página **Configuración**, elija **Correo**.
4. En la página **Opciones**, seleccione **General** y, a continuación, **Administrar complementos** (como se muestra en la siguiente captura de pantalla).

 ![Página de administración de complementos](./readme-images/block-on-send-manage-addins.png)

5. En la página **Administrar complementos**, haga clic en el signo "+" y seleccione **Agregar desde archivo**. Navegue al archivo de manifiesto `Contoso Message Body Checker.xml` que se incluye en el proyecto. Haga clic en **Siguiente** y luego en **Instalar**. Finalmente, haga clic en **Aceptar**.
6. Repita el paso 5 para instalar el archivo de manifiesto `Contoso Subject and CC Checker.xml`.
7. Vuelva a la vista de correo de Outlook Web App.


## Ejecutar el complemento

### Comprobador de asunto y CC

1. Redacte un nuevo mensaje de correo electrónico de Outlook Web App. 
2. Deje la línea de asunto en blanco.
3. Agregue un destinatario en la línea **Para**. 
4. Haga clic en **Enviar**. 

* Se ha agregado una copia carbón a la línea CC. En este ejemplo, se agregó `Contoso@contoso.onmicrosoft.com`.
* El envío del correo electrónico está bloqueado y se muestra un mensaje de error en la barra de información para notificar al remitente que debe agregar un asunto (como se muestra en la siguiente captura de pantalla).  

 ![Barra de información del comprobador de asunto y CC](./readme-images/block-on-send-subject-cc-inforbar.png) 

6. Agregue una línea de asunto.
7. El texto `[Comprobado]:` se agrega al principio de la línea de asunto y el correo electrónico se envía.

### Comprobador del cuerpo del mensaje

1. Redacte un nuevo mensaje de correo electrónico de Outlook Web App. 
2. En el cuerpo del mensaje, escriba `blockedword`, `blockedword1` o `blockedword2`. (Esta es la matriz de palabras restringidas en el archivo app.js de la función `checkBodyOnlyOnSendCallBack`).
3. Agregue un destinatario en la línea **Para**. 
5. Haga clic en **Enviar**.  

* Se bloquea el envío del correo electrónico por las palabras que se encuentran en el cuerpo del mensaje.  
* Se muestra un mensaje de error en la barra de información para notificar al remitente que se encontraron palabras bloqueadas (como se muestra en la siguiente captura de pantalla).  

 ![Barra de información del comprobador del cuerpo del mensaje](./readme-images/block-on-send-body.png)

5. Para enviar el correo electrónico, quite las palabras bloqueadas.

## Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre este ejemplo. Usted puede enviarnos sus comentarios a través de la sección *Problemas* de este repositorio.

Las preguntas generales sobre el desarrollo de Microsoft Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si su pregunta trata sobre las API de JavaScript para Office, asegúrese de que su pregunta se etiqueta con [office-js] y [API].

## Recursos adicionales

* [Documentación de complemento de Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centro para desarrolladores de Office](http://dev.office.com/)
* Más ejemplos de complementos de Office en [OfficeDev en GitHub](https://github.com/officedev)

## Derechos de autor
Copyright (c) 2016 Microsoft Corporation. Todos los derechos reservados.



Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
