/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
(function() {
    "use strict";

    // The initialize function is run each time a page of the add-in is loaded into the task pane.
    Office.initialize = function(reason) {
        $(document).ready(function() {

            // Use this to check whether the new API is supported in the Word client.
            // The createDocument method call is in the 1.3 requirement set. 
            // The 1.3 requirement set check is not implemented in preview. 
            // The 1.2 API requirement set check is the minimum requirement check in Word. 
            // Update this to target the correct set after 1.3 is generally available.
            if (Office.context.requirements.isSetSupported("WordApi", 1.2)) {
                // Initialize stylized fabric UI for text fields.
                $(".ms-TextField").TextField();

                // If you need to initialize something you can do so here.
                // This is in search.js.
                $('#btnSearch').click(RedactAddin.searchWords);
                // This is in redact.js.
                $('#btnRedact').click(RedactAddin.redactUserInput);  
                 // This is in redact.js.
                $('#btnRedactSelection').click(RedactAddin.redactSelectedTexts);
            }
            else {                
                // Just letting you know that this code will not work with your version of Word.
                console.log('This add-in requires the latest version of Word.');
            }
        });
    };    
})();
