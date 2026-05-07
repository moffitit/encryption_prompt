/*
 * launchevent.js — Smart Alerts handler for OnMessageSend
 *
 * SendMode="PromptUser" means:
 *   - If we call event.completed({ allowEvent: false }), Outlook shows a dialog:
 *       "Does this email need to be encrypted?"
 *       [Don't Send]  [Send Anyway]
 *   - "Don't Send" returns the user to the compose window where they can
 *     click the Encrypt Email ribbon button to apply OME, then send again.
 *   - "Send Anyway" fires the event again — we detect this via a custom
 *     property and allow it through.
 *
 * This file must be a plain .js file served at the URL in LaunchEvent.Url.
 * It cannot be inside an HTML page.
 */

/* global Office */

// Key we stamp on the message after the user chooses "Send Anyway"
var SEND_ANYWAY_KEY = "EncryptPromptShown";

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;

  // Check if the user already saw the prompt and chose "Send Anyway"
  item.sessionData.getAsync(SEND_ANYWAY_KEY, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value === "true") {
      // User already answered — allow send
      item.sessionData.clearAsync(function() {
        event.completed({ allowEvent: true });
      });
      return;
    }

    // First time hitting Send — stamp the key and block with the prompt
    item.sessionData.setAsync(SEND_ANYWAY_KEY, "true", function() {
      event.completed({
        allowEvent: false,
        cancelLabel: "Don't Send",          // button to return to compose
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
        errorMessage: "Does this email need to be encrypted? Click \u2018Don\u2019t Send\u2019 to go back and click \u2018Encrypt Email\u2019 in the ribbon to apply Microsoft Purview OME encryption, or \u2018Send Anyway\u2019 to send without encryption."
      });
    });
  });
}

// Register the handler
Office.onReady(function() {});
