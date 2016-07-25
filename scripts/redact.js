/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/


var RedactAddin;
(function (RedactAddin) {

    /**
    * Search and redact text.
    */
    function redactUserInput() {
        // Get user text input from textfield in the taskpane.
        var redactInput = document.getElementById('inputRedact').value;
        RedactAddin.redactWordsCommon(redactInput);
    }
    RedactAddin.redactUserInput = redactUserInput;

    /**
    * Redact current selected text.
    */
    function redactSelectedTexts() {     
        Word.run(function (context) {
            // Get selected text.
            var selection = context.document.getSelection();
            selection.load("text");
            return context.sync().then(function() {
                RedactAddin.redactWordsCommon(selection.text);
            });
        }).catch(RedactAddin.handleErrors);
    }
    RedactAddin.redactSelectedTexts = redactSelectedTexts;

})(RedactAddin || (RedactAddin = {}));

