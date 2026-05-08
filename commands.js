/* ------------------------------------------------------------------ *
 *  Encryption Reminder – Outlook Add-in                               *
 *  commands.js                                                        *
 * ------------------------------------------------------------------ */

Office.onReady(function () {
  Office.actions.associate("onItemSend", onItemSend);
});

/**
 * Called when the user clicks Send.
 * Passes allowEvent: false + an errorMessage, which triggers Outlook's
 * built-in "Send Anyway / Don't Send" prompt (no custom dialog needed).
 */
function onItemSend(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Does this email need to be encrypted?"
  });
}
