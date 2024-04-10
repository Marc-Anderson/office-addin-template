/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
async function action(event) {

    const currentEmail = Office.context.mailbox.item;
    // const subject = currentEmail.subject;

    // define a function for fetching the email body asynchronously and await before proceeding
    const emailBody = await getEmailBodyAsync(currentEmail);
    function getEmailBodyAsync(email) {
        return new Promise((resolve, reject) => {
            email.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error("Failed to get email body."));
                }
            });
        });
    }
    
    // process the email body
    const emailAboutTacos = emailBody.includes("tacos") ? "is" : "is not";

    // create a notification message.
    // https://learn.microsoft.com/en-us/javascript/api/outlook/office.notificationmessages
    const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `This email ${emailAboutTacos} about tacos`,
        icon: "Icon.80x80",
        persistent: true,
    };

    // Show the notification message.
    currentEmail.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

    // Be sure to indicate when the add-in command function is complete.
    event.completed();
    
}



async function insertTextSample(event) {

    // target the currently selected mail item
    const item = Office.context.mailbox.item

    // define some string to insert into the message body
    let textToInsertIntoMessage = "Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui dolorem ipsum quia dolor sit amet, consectetur, adipisci velit"

    // Identify the body type of the mail item.
    // this is required since trying to insert text using the wrong coerciontype will result in an error
    // calling body.getTypeAsync as stated in the docs intermittently skips calling the callback
    // docs: learn.microsoft.com/en-us/javascript/api/outlook/office.body
    let emailBodyTypeResponse = await new Promise((resolve, reject) => {
        item.body.getTypeAsync((asyncResult) => {
            if (asyncResult) {
                resolve(asyncResult);
            } else {
                reject(new Error("No result from getTypeAsync"));
            }
        });
    // if there was a problem processing email body type, try to use html since thats the most common
    }).catch(e=>{ return {value: 'html', status: 'failed', message: 'no result received from body.getTypeAsync'}})

    // insert the text into the selected email with the above mentioned coercion type
    item.body.setSelectedDataAsync(textToInsertIntoMessage, {coercionType: emailBodyTypeResponse.value}, function(result){ 
        // console.log("do nothing, but you could");
    })
    
    // Be sure to indicate when the add-in command function is complete.
    event.completed();
}


// Register the function with Office.
Office.actions.associate("action", action);
Office.actions.associate("insertTextSample", insertTextSample);
