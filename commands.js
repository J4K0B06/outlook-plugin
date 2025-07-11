Office.onReady(() => {
  // Office is ready
});

/*
function reportPhishing(event) {
  console.log("reportPhishing triggered");
  const item = Office.context.mailbox.item;
  if (!item) {
    console.log("No item found");
  } else {
    console.log("Item subject:", item.subject);
  }

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["info@safebyte.be"],
    subject: "Suspicious email: " + item.subject,
    htmlBody: `
      <p>This message was reported as phishing.</p>
      <p><strong>From:</strong> ${item.from && item.from.emailAddress}</p>
      <p><strong>Subject:</strong> ${item.subject}</p>
      <p>Body: (please forward the original message manually if needed)</p>
    `
  });

  event.completed();
}
*/

function reportPhishing(event) {
  console.log("reportPhishing triggered");
  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["info@safebyte.be"],
    subject: "Test phishing report",
    htmlBody: "<p>Test phishing report body</p>"
  });
  event.completed();
}

Office.actions.associate("reportPhishing", reportPhishing);
