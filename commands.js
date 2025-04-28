/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(function() {
  // Register the function
  Office.actions.associate("autoOpenTaskpane", autoOpenTaskpane);
});

function autoOpenTaskpane(event) {
  // Auto-open the taskpane when the message is being composed
  // Use a minimal size dialog that's less intrusive
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/taskpane.html",
    { height: 10, width: 20, displayInIframe: true, promptBeforeOpen: false },
    function(result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // If the dialog fails to open, handle the error
        console.error("Failed to open taskpane: " + result.error.message);
      } else {
        // Store the dialog for later reference
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    }
  );
  
  // Return a success message
  event.completed();
}

function processMessage(message) {
  // Process messages from the dialog if needed
  console.log("Message received from dialog: " + message.message);
} 