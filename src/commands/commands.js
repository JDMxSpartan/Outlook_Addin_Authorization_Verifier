/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = function (reason) {};

/**
 * Handles the OnMessageRecipientsChanged event.
 * @param {*} event The Office event object
 */
function MessageSendVerificationHandler(event) {
  console.log("MessageSendVerificationHandler method"); //debugging
  console.log("event: " + JSON.stringify(event)); //debugging
  
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get To recipients. " + JSON.stringify(asyncResult.error));
      return;
    }
    
    const toRecipients = asyncResult.value;
    console.log("checking the classification of recipient: "+ toRecipients);
    checkRecipientClassification(toRecipients)
      .then(allowEvent => {
        if (!allowEvent) {
          // Prevent sending the email
          event.completed({ allowEvent: false });
          Office.context.mailbox.item.notificationMessages.addAsync(
            "unauthorizedSending",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
              message: "You are not authorized to send this email to meaganbmueller@gmail.com."
            }
          );
        }
      })
      .catch(error => {
        console.error("Error occurred while checking recipient classification: " + error);
      });
  });
}

/**
 * Checks the classification level of the recipients.
 * @param {array} recipients The array of recipients
 * @returns {Promise<boolean>} A promise that resolves with a boolean indicating whether the event should proceed
 */
function checkRecipientClassification(recipients) {
  console.log("checkRecipientClassification method"); //debugging

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    
    recipients.forEach(function (recipient) {
      const emailAddress = recipient.emailAddress;
      console.log(emailAddress);

      // Check if recipient is unauthorized
      if (isUnauthorized(emailAddress)) {
        console.log("isUnauthorized returned: " + isUnauthorized(emailAddress));
        allowEvent = false;
      }

    });

    console.log("event should proceed since isUnauthorized returned false");

    // Allow event to proceed if no unauthorized recipient found
    resolve(allowEvent);
  });
}

/**
 * Determines if the recipient is unauthorized.
 * @param {string} emailAddress The recipient's email address
 * @returns {boolean} True if unauthorized, false otherwise
 */
function isUnauthorized(emailAddress) {
  // Check if the recipient's email address matches the unauthorized email address
  return emailAddress === 'meaganbmueller@gmail.com';
}

/**
 * Retrieves the clearance level based on the recipient's email address.
 * @param {string} emailAddress The recipient's email address
 * @returns {string|null} The clearance level required or null if no clearance is needed
 */
function getClearanceLevel(emailAddress) {
  // Perform your logic to determine the clearance level based on the recipient's email address
  // For demonstration, let's assume 'meaganbmueller@gmail.com' requires a 'Classified' clearance
  if (emailAddress === 'meaganbmueller@gmail.com') {
    return 'Classified';
  }
  // If the recipient doesn't require any special clearance, return null
  return null;
}

 function _setSessionData(key, value) {
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    function(asyncResult) {
      // Handle success or error.
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
      if (value) {
        _tagExternal(value);
      } else {
        _checkForExternal();
      }
    } else {
      console.error(`Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(asyncResult.error)}`);
      return;
    }
  });
}


// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("MessageSendVerificationHandler", MessageSendVerificationHandler);