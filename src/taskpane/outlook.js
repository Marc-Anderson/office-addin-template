/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
let userToken;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {

    // Link to full token sample: https://github.com/OfficeDev/office-js-snippets/blob/main/samples/outlook/85-tokens-and-service-calls/user-identity-token.yaml
    Office.context.mailbox.getUserIdentityTokenAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(`Token retrieval failed with message: ${result.error.message}`);
        } else {
            console.log(result.value);
            userToken = result.value
        }
    });

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run(event) {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;

  // sometimes values arent immediately available so you need to use getAsync() on them
  // in this case, the subect is available when reading, but not when composing
  // because we're using the same fn for both we need to account for that
  // generally you would just use callbacks but not today
  let subject = item.subject;
  if(typeof item.subject !== "string"){
    subject = await getEmailSubjectAsync(item);
  }
  function getEmailSubjectAsync(email) {
    return new Promise((resolve, reject) => {
        email.subject.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(new Error("Failed to get email subject."));
            }
        });
    });
  }



  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + subject;
  document.getElementById("user-token").innerHTML = "<b>Token:</b> <br/>" + userToken;

}
