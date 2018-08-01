// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.

Office.initialize = function () {
}

// Helper function to add a status message to
// the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

// Adds text into the body of the item, then reports the results
// to the info bar.
function addTextToBody(text, icon, event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, 
    function (asyncResult){
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        statusUpdate(icon, "\"" + text + "\" inserted successfully.");
      }
      else {
        Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
          type: "errorMessage",
          message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
        });
      }
      event.completed();
    });
}

function addDefaultMsgToBody(event) {
  addTextToBody("Inserted by the Add-in Command Demo add-in.", "blue-icon-16", event);
}
// Gets the subject of the item and displays it in the info bar.

function getAttendees(event) {
  var RequiredMeetingAttendees = Office.context.mailbox.item.requiredAttendees.getAsync(callback);
  var OptionalMeetingAttendees = Office.context.mailbox.item.optionalAttendees.getAsync(callback);
	
	Office.context.mailbox.item.notificationMessages.addAsync("RequiredMeetingAttendees", {
    type: "informationalMessage",
    icon: "red-icon-16",
    message: "Required Meeting Attendees: " + RequiredMeetingAttendees,
    persistent: false
	
  });
  
  event.completed();
}

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {

};

// Add any ui-less function here
