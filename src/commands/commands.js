/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var mailboxItem;

Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox.item;
};

// Called when dialog signs in the user. this is an example, can be deleted
function userSignedIn() {
  Office.context.ui.messageParent(true.toString());
}

/**
 * Handles the OnMessageSend event.
 /**
 * Initializes all of the code.\
 * @param {*} event The Office event object
 */
function MessageSendVerificationHandler(event) {
  //promise is to encapsulate all the asynch functions
  Promise.all([
    getToRecipientsAsync(),
    getSenderAsync(),
    getBodyAsync(),
    fetchAndParseCSV(),
  ]).then(([toRecipients, sender, body]) => {
    console.log("To recipients:");
    toRecipients.forEach((recipient) => console.log(recipient.emailAddress));
    console.log("Sender:" + sender.displayName + " " + sender.emailAddress);
    console.log("Body:" + body);
    //const bannerMarkings = parseBannerMarkings(body);
    const banner = getBannerFromBody(body);

    // Check if the banner is null
    if (banner == null) {
      console.log("banner is null, so should not send email");

      const options = {
        height: 30,
        width: 20,
        promptBeforeOpen: false,
        displayInIframe: true,
    };
      //the commented out displays a new window...
      Office.context.ui.displayDialogAsync('https://meg217.github.io/Outlook_Addin_Authorization_Verifier/src/commands/dialog.html', options);
      //difference between errorMessage and informationalMessage?
      mailboxItem.notificationMessages.addAsync("NoSend", {
        type: "errorMessage",
        message: "Please enter a banner marking for this email.",
      });

      //maybe shouln't de-allow event? instead make a dialog box show up?
      // console.log("event should be denied");
      // event.completed({ allowEvent: false });
      
      var errorElement = document.querySelector('div.zezGF');
      var errorElement2 = document.querySelector('div.ms-Dialog-main');

      console.log(errorElement2);
      console.log(errorElement);
  

      return;
    }

    //const messageBodyTest = "TOP SECRET//COMINT-GAMMA/TALENT KEYHOLE//ORIGINATOR CONTROLLED";
    const bannerMarkings = parseBannerMarkings(banner);
    console.log(bannerMarkings);

    checkRecipientClassification(toRecipients)
      .then((allowEvent) => {
        if (!allowEvent) {
          // Prevent sending the email
          console.log("Prevent sending email");
          event.completed({ allowEvent: false });
          Office.context.mailbox.item.notificationMessages.addAsync(
            "unauthorizedSending",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType
                .ErrorMessage,
              message: "You are not authorized to send this email",
            },
            (result) => {
              console.log(result);
            }
          );
        } else {
          // Allow sending the email
          event.completed({ allowEvent: true });
        }
      })
      .catch((error) => {
        console.error(
          "Error occurred while checking recipient classification: " + error
        );
      });
  });
}

function fetchCSVData(url) {
  return fetch(url).then((csvData) => parseCSV(csvData));
}

/**
 * sets session data
 * key and value parameters
 */
function _setSessionData(key, value) {
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
        if (value) {
          _tagExternal(value);
        } else {
          _checkForExternal();
        }
      } else {
        console.error(
          `Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(
            asyncResult.error
          )}`
        );
        return;
      }
    }
  );
}

//this is the old code from the example
/**
 * Handles the 'to' authentication.
 * @param {*} event The Office event object
 */
function FAKEtoHandler(event) {
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(
        "Failed to get To recipients. " + JSON.stringify(asyncResult.error)
      );
      return;
    }

    const toRecipients = asyncResult.value;
    console.log("checking the classification of recipient: " + toRecipients);
    checkRecipientClassification(toRecipients)
      .then((allowEvent) => {
        if (!allowEvent) {
          // Prevent sending the email
          event.completed({ allowEvent: false });
          Office.context.mailbox.item.notificationMessages.addAsync(
            "unauthorizedSending",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType
                .ErrorMessage,
              message:
                "You are not authorized to send this email to meaganbmueller@gmail.com.",
            }
          );
        } else {
          // Allow sending the email
          event.completed({ allowEvent: true });
        }
      })
      .catch((error) => {
        console.error(
          "Error occurred while checking recipient classification: " + error
        );
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
  return emailAddress === "meaganbmueller@gmail.com";
}

/**
 * Retrieves the clearance level based on the recipient's email address.
 * @param {string} emailAddress The recipient's email address
 * @returns {string|null} The clearance level required or null if no clearance is needed
 */
function getClearanceLevel(emailAddress) {
  // Perform your logic to determine the clearance level based on the recipient's email address
  // For demonstration, let's assume 'meaganbmueller@gmail.com' requires a 'Classified' clearance
  if (emailAddress === "meaganbmueller@gmail.com") {
    return "Classified";
  }
  // If the recipient doesn't require any special clearance, return null
  return null;
}

//  function _setSessionData(key, value) {
//   Office.context.mailbox.item.sessionData.setAsync(
//     key,
//     value.toString(),
//     function(asyncResult) {
//       // Handle success or error.
//       if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//       console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
//       if (value) {
//         _tagExternal(value);
//       } else {
//         _checkForExternal();
//       }
//     } else {
//       console.error(`Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(asyncResult.error)}`);
//       return;
//     }
//   });
// }

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate(
  "MessageSendVerificationHandler",
  MessageSendVerificationHandler
);
