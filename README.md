# Word  JavaScript Redact Add-in

Learn how you can create an add-in that searches, highlights and redacts text.    

## Table of Contents
* [Change History](#change-history)
* [Important note](#important-note)
* [Prerequisites](#prerequisites)
* [Configure the project](#configure-the-project)
* [Run the project](#run-the-project)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Change History

July 25, 2016:
* Initial sample version.

##Important note

This sample redacts text by replacing each letter with an asterisk.  The layout is not preserve as such.  No measurement and calculation is done to determine font, size, or formatting of each letter.

## Prerequisites

* Word 2016 for Windows, build 16.0.6912.1000 or later.
* [Node and npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) - You should use a later build as earlier builds can show an error when generating the certificates.

## Configure the project

Run the following commands from your Bash shell at the root of this project:

1. Clone this repo to your local machine.
2. ```npm install``` to install all of the dependencies in package.json.
3. ```bash gen-cert.sh``` to create the certificates needed to run this sample. 
* Then in the repo on your local machine, double-click ca.crt, and select **Install Certificate**. 
* Select **Local Machine** and select **Next** to continue. 
* Select **Place all certificates in the following store** and then select **Browse**.  
* Select **Trusted Root Certification Authorities** and then select **OK**. 
* Select **Next** and then **Finish** and now your certificate authority cert has been added to your certificate store.
4. ```npm start``` to start the node web server service.

Now you need to let Microsoft Word know where to find the add-in.

1. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx) and place the [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) manifest file in it.
3. Launch Word and open a document.
4. Choose the **File** tab, and then choose **Options**.
5. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
6. Choose **Trusted Add-ins Catalogs**.
7. In the **Catalog Url** field, enter the network path to the folder share that contains word-add-in-js-redact-manifest.xml, and then choose **Add Catalog**.
8. Select the **Show in Menu** check box, and then choose **OK**.
9. A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close and restart Word.

## Run the project

1. Open a Word document.
2. On the **Insert** tab in Word 2016, choose **My Add-ins**.
3. Select the **SHARED FOLDER** tab.
4. Choose **Word Redact add-in**, and then select **OK**.
5. If add-in commands are supported by your version of Word, the UI will inform you that the add-in was loaded.

### Ribbon UI

On the Ribbon:
* Select **Review** tab and choose **Show Redaction Task Pane** to launch the task pane.

 > Note: The add-in will load in a task pane if add-in commands are not supported by your version of Word.

### Task pane UI

On the task pane, you can:
* Search and highlight text in a document by entering word in the text box and selecting the **Search and Highlight** button.
  
> NOTE:  Search is case-sensitive.  To undo an action, press Ctrl+Z on your keyboard.

* Redact text in a document by entering word in the text box and selecting **Redact** button.
  
> NOTE:  Redact is case-insensitive.   

* Redact selected text in a document by choosing **Redact selected text** button.
  
> NOTE:  Redact is case-insensitive.       
  
### Add-in commands

This sample shows how to create add-in commands in the ribbon. The add-in command declarations are located in word-add-in-js-redact-manifest.xml. 

## Questions and comments

We'd love to get your feedback about the Word Redact sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Microsoft Office 365 development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Make sure that your questions are tagged with [office-js] and [API].

## Additional resources

* [Office add-in documentation](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs starter projects and code samples](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## Copyright
Copyright (c) 2016 Microsoft Corporation. All rights reserved.

