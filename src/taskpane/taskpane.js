/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("myButton").onclick = run;
  }
});

export async function run() {
  var item = Office.context.mailbox.item;
  // Get the body of the message 
  item.body.getAsync("text", function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("The body of the message is: " + result.value);
      let email_content = result.value;

      // Construct the Prompt

      let prompt = `Generate a polite and professional HTML formatted email response to the following unsolicited email: ${email_content}
      I am not currently interested in whatever this user and email presents. 
      The email should express appreciation for their interest, explain that we are currently focused on internal projects, 
      and indicate that we will keep their information for future consideration. 
      Please find the sender's name, company name and use it for a sincere email.
      Address the user by his name.
      The email should close with a polite sign-off.
      The response should be in HTML format with appropriate paragraph breaks for readability.      
      `;
      sendToOpenAI(prompt, item);
    } else {
      console.log("Failed to get the body of the message. Error: " + result.error.message);
    }
  });

  
}

export async function sendToOpenAI(prompt, item) {
  console.log('sendToOpenAI')
  return fetch('https://api.openai.com/v1/completions', {
      method: 'POST',
      headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ...'
      },
      body: JSON.stringify({
          model: "text-davinci-003",
          prompt: prompt,
          temperature: 0,
          max_tokens: 1024,
          top_p: 1,
          frequency_penalty: 0,
          presence_penalty: 0,
      })
  }).then(response => {
      console.log('response taken')
      response.json().then(res => {
        console.log(res.choices[0].text)
        // Display a reply form
        item.displayReplyForm({
          'htmlBody': res.choices[0].text
        });
      })
  });
};