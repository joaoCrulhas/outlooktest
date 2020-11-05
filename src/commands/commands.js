
const newBody = '<br>' +
  '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
  '<br><br>' +
  'Phone Dial-in: +1(123)456-7890' +
  '<br><br>' +
  'Meeting ID: 123 456 789' +
  '<br><br>' +
  'Want to test your video connection?' +
  '<br><br>' +
  '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
  '<br><br>';

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
var mailboxItem;

Office.onReady(() => {
  mailboxItem = Office.context.mailbox.item;
  console.log(mailboxItem);
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {    // Get HTML body from the client.
  mailboxItem.body.getAsync("html",
    { asyncContext: event },
    function (getBodyResult) {
      if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
        updateBody(getBodyResult.asyncContext, getBodyResult.value);
      } else {
        console.error("Failed to get HTML body.");
        getBodyResult.asyncContext.completed({ allowEvent: false });
      }
    }
  );
}

function updateBody(event, existingBody) {
  // Append new body to the existing body.
  mailboxItem.body.setAsync(existingBody + newBody,
    { asyncContext: event, coercionType: "html" },
    function (setBodyResult) {
      if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
        setBodyResult.asyncContext.completed({ allowEvent: true });
      } else {
        console.error("Failed to set HTML body.");
        setBodyResult.asyncContext.completed({ allowEvent: false });
      }
    }
  );
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
