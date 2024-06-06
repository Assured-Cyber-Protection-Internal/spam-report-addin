/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
 */

// Ensures the Office.js library is loaded.
Office.onReady(() => {
    /**
     * IMPORTANT: To ensure your add-in is supported in the classic Outlook client on Windows,
     * remember to map the event handler name specified in the manifest to its JavaScript counterpart.
     */
    //if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
      Office.actions.associate("onSpamReport", onSpamReport); //}
  }
);
  



  // Handles the SpamReporting event to process a reported message.
  function onSpamReport(event) {
    // body.log("onSpamReport")
    document.getElementById("ModalFocusTrapZone972").innerHTML += "<h1>HELLO</h1>";
    showNotification("onSpamReport");
    // alert("onSpamReport");
    // Get the Base64-encoded EML format of a reported message.
    Office.context.mailbox.item.getAsFileAsync(
      { asyncContext: event },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(
            `Error encountered during message processing: ${asyncResult.error.message}`
          );
          return;
        }
  
        // Get the user's responses to the options and text box in the preprocessing dialog.
        // const spamReportingEvent = asyncResult.asyncContext;
        // const reportedOptions = spamReportingEvent.options;
        // const additionalInfo = spamReportingEvent.freeText;
  
        
        // Run additional processing operations here.
        
        // Send a POST request to the web service at http://127.0.0.1:5000/api/interceptions/report with the following information:
        // {
        //   from: <email the email came from>
        //   to: <Email that received the email>
        //   body: <the entirety of the email html>
        //   subject: <the subject of the email>
        //   receivedDate: <Date the email was received>
        // }
        var xhr = new XMLHttpRequest();
        var url = 'https://api-beta.republic.recyber.com/api/interceptions/report';

        xhr.open('POST', url, true);
        xhr.setRequestHeader('Content-Type', 'application/json');

        xhr.onreadystatechange = function () {
          if (xhr.readyState === 4 && xhr.status === 200) {
            var res = JSON.parse(xhr.responseText);
            console.log(res);

            const event = asyncResult.asyncContext;
            event.completed({
              // onErrorDeleteItem: true,
              // moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
              showPostProcessingDialog: {
                title: "Phishing Reporting",
                description: "Thank you for reporting this message.",
              },
            });
          }
        };

        var data = JSON.stringify({
          from: Office.context.mailbox.item.from.emailAddress,
          to: Office.context.mailbox.item.to[0].emailAddress,
          // body: "Axios " + Office.context.mailbox.item.body.getAsync(),
          body: "Axios",
          subject: Office.context.mailbox.item.subject,
          receivedDate: Office.context.mailbox.item.dateTimeCreated
        });

        xhr.send(data);

        // axios.post('https://api-beta.republic.recyber.com/api/interceptions/report', {
        //   from: Office.context.mailbox.item.from.emailAddress,
        //   to: Office.context.mailbox.item.to[0].emailAddress,
        //   // body: "Axios " + Office.context.mailbox.item.body.getAsync(),
        //   body: "Axios",
        //   subject: Office.context.mailbox.item.subject,
        //   receivedDate: Office.context.mailbox.item.dateTimeCreated
        // }, {
        //   headers: {
        //     'Content-Type': 'application/json',
        //   }
        // }).then(res => {
        //   console.log(res.data)
        // }).catch(err => {
        //   console.error(err)
        // })
// //https://vczw8bvg-5000.uks1.devtunnels.ms/
//         console.log("Fetching");
//         fetch('https://vczw8bvg-5000.uks1.devtunnels.ms/api/interceptions/report', {
//           method: 'POST',
//           headers: {
//             'Content-Type': 'application/json',
//           },
//           body: JSON.stringify({
//             from: Office.context.mailbox.item.from.emailAddress,
//             to: Office.context.mailbox.item.to[0].emailAddress,
//             body: "Fetch " + Office.context.mailbox.item.body.getAsync(),
//             subject: Office.context.mailbox.item.subject,
//             receivedDate: Office.context.mailbox.item.dateTimeCreated
//           })
//         })
//         .then(response => response.json())
//         .then(data => {
//           console.log('Success:', data);
//         })
//         .catch((error) => {
//           console.error('Error:', error);
//         });




        /**
         * Signals that the spam-reporting event has completed processing.
         * It then moves the reported message to the Junk Email folder of the mailbox,
         * then shows a post-processing dialog to the user.
         * If an error occurs while the message is being processed,
         * the `onErrorDeleteItem` property determines whether the message will be deleted.
         */
      }
    );
  }