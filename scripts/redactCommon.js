/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

/**
 * * Common shared functions.
     */

var RedactAddin;
(function (RedactAddin) {

    /**
    * Redact text and phrases function.
    */
    function redactWordsCommon(redactInput) {
        Word.run(function (context) {
            // Get the OOXML of the original document.
            var docOOXML = context.document.body.getOoxml();

            return context.sync()
                .then(function () {
                    // Create a new document.
                    var newDoc = context.application.createDocument();

                    var redactInputLength = redactInput.length;

                    // Construct regex to /userinput/gi  
                    // Do a global case insensitive search.           
                    var regularExpression = new RegExp(redactInput, "gi");

                    // Before adding any content to the document, 
                    // replace all intances of the input text (case insensitive), with asterisks.                
                    var newString = docOOXML.value.replace(regularExpression, Array(redactInputLength + 1).join('*'));

                    // Insert the resulting OOXML in the document.
                    newDoc.body.insertOoxml(newString, "replace");

                    // Search for the asterisks.
                    var newSearchResults = newDoc.body.search(Array(redactInputLength + 1).join('*'));
                    context.load(newSearchResults);

                    return { doc: newDoc, searchResults: newSearchResults };
                })
                .then(context.sync)
                .then(function (data) {
                    for (var i = 0; i < data.searchResults.items.length; i++) {
                        // Black out each search result.
                        data.searchResults.items[i].font.highlightColor = "#000000";
                    }
                    // Open the document.
                    data.doc.open();
                })
                .then(context.sync)
        }).catch(RedactAddin.handleErrors);
    }
    RedactAddin.redactWordsCommon = redactWordsCommon;

    /**
    * Error handling function.
    */
    function handleErrors(error) {
        var url = createUrlForDialog('dialogalert.html',
            { message: "Error", description: error.toString() });

        Office.context.ui.displayDialogAsync(url,
            { height: 17, width: 15});

        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    }
    RedactAddin.handleErrors = handleErrors;

     /**
     * Open the dialog to provide notification of found words.
     */
    function createUrlForDialog(pageUrl, data) {
        var urlComponents = [];
        if (data) {
            for (var d in data) {
                urlComponents.push(encodeURIComponent(d) + "=" + encodeURIComponent(data[d]));
            }
        }
        return window.location.protocol + '//' + window.location.host + window.location.pathname + pageUrl + (data ? "?" : "") + urlComponents.join("&");
    }
    RedactAddin.createUrlForDialog = createUrlForDialog;

})(RedactAddin || (RedactAddin = {}));

