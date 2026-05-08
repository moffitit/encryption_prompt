Office.onReady(function () {
  Office.actions.associate("onItemSend", onItemSend);
});

/**
 * Smart Alert handler — triggered by OnMessageSend.
 * Returning allowEvent: false with an errorMessage causes Outlook to show
 * its built-in prompt: "Does this email need to be encrypted?
 * [Send Anyway]  [Don't Send]"
 */
function onItemSend(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Does this email need to be encrypted?"
  });
}
