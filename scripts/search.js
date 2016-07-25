/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
* Search and highlight texts.
*/

var RedactAddin;
(function (RedactAddin) {

    function searchWords() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            var searchInput = document.getElementById('inputSearch').value;

            // Search the document.
            var searchResults = context.document.body.search(searchInput, {matchWildCards: true});

            // Load the search results and get the font property values.
            context.load(searchResults, "");

            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    var count = searchResults.items.length;
                    // // Queue a set of commands to change the font for each found item.
                    for (var i = 0; i < count; i++) {
                        searchResults.items[i].font.highlightColor = '#FFFF00'; //Yellow            
                    }
                    return count;
                })
                .then(context.sync)
                .then(reportWordsFound);

        }).catch(RedactAddin.handleErrors);
    }
    RedactAddin.searchWords = searchWords;

    /**
     * Open the dialog to provide notification of found words.
     */
    function reportWordsFound(count) {
        var url = RedactAddin.createUrlForDialog('dialogCount.html', {count: count});
        Office.context.ui.displayDialogAsync(url,
            { height: 11, width: 12, requireHTTPS: true });
    }

})(RedactAddin || (RedactAddin = {}));





