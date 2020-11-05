
const newBody = `\n<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">\n\n<html>\n  <head>\n    <title>Fuze Meeting</title>\n    <meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">\n    <style media=\"all\" type=\"text/css\">\n      body {\n          font-family: arial, sans-serif;\n          font-size: 1.0em;\n          background-color: white;\n      }\n      \n      #content {\n          margin: 0 auto;\n          width: 700px;\n      }\n      #email_body {\n          margin: 1em;\n      }\n  </style>\n  \n  <style type=\"text/css\">\n    * {\n      font-family: Arial, sans-serif;\n      font-size: 14px;\n    }\n    ul {\n      margin-top: 0;\n    }\n  </style>\n\n  </head>\n  <body>\n    \n<script type=\"application/ld+json\">\n{\n  \"@context\": \"http://schema.org\",\n  \"@type\": \"EmailMessage\",\n  \"potentialAction\": {\n    \"@type\": \"ViewAction\",\n    \"target\": \"http://fuze.me/7428363\",\n    \"name\": \"Join meeting\"\n  },\n  \"description\": \"Join Online Meeting: New Meeting\"\n}\n</script>\n\n    <div id=\"content\">\n      <div id=\"email_body\">\n        \n  \n\n  <p style=\"font-family: Arial, sans-serif; font-size: 14px; margin-top: 14px; margin-bottom: 14px;\">\n    \t **** Please do not edit the information below **** \t\n  </p>\n\n  <p style=\"font-family: Arial, sans-serif; font-size: 14px; margin-top: 14px; margin-bottom: 14px;\">\n    <b style=\"font-family: Arial, sans-serif; font-size: 14px;\">Join from your computer or mobile:</b> <a href=\"https://intg.fuze.me/7428363\">https://intg.fuze.me/7428363</a>\n  </p>\n\n  \n    \n    \n    \n    \n    \n    \n    \n    \n\n    \n    <p style=\"font-family: Arial, sans-serif; font-size: 14px; margin-top: 14px; margin-bottom: 14px;\">\n      <b style=\"font-family: Arial, sans-serif; font-size: 14px;\">Or join by phone:</b>\n    <ul style=\"font-family: Arial, sans-serif; font-size: 14px;\">\n      <li> Bulgaria:</li>\n    </ul>\n\n    \n    <span style=\"font-family: Arial, sans-serif; font-size: 14px;\">Additional numbers:</span><br><br>\n    <ul style=\"font-family: Arial, sans-serif; font-size: 14px;\">\n      \n        <li> Bulgaria: 0000014 (toll free) </li> \n\n      \n        <li> Australia:   8002222222 (toll free) </li>  <li> Canada:  <b>&#43;1 647-560-9999</b>   </li>  <li> Cyprus:  <b>&#43;357 80 096942</b>   </li>  \n   \n      <li>International numbers available <a href=\"https://intg.fuze.me/7428363/dialin\">here</a></li>\n\n    </ul>\n  \n\n  ______________\n  <p><br /></p> \n  <p  style=\"font-family: Arial, sans-serif; font-size: 14px;\"> <b style=\"font-family: Arial, sans-serif; font-size: 14px;\">First Fuze meeting?</b>  <br>\n    <a href=\"https://www.fuze.com/download?utm_source=Meeting-Invite\" style=\"font-family: Arial, sans-serif; font-size: 14px;\">Download Fuze</a> ahead of time for the best experience.</p>\n  \n    <span style=\"font-family: Arial, sans-serif; font-size: 14px;\">Joining from a video-conferencing room or system?</span><br>\n    <span style=\"font-family: Arial, sans-serif; font-size: 14px;\">-Dial: <a href=\"7428363@tp.fuze.me\">7428363@tp.fuze.me </a></span>\n   \n  \n\n      </div>\n    </div>\n  </body>\n</html>\n`;

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
