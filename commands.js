Office.onReady(() => {
  // Office is ready
});

function reportPhishing(event) {
  const item = Office.context.mailbox.item;

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: ["phishing@safebyte.be"],
    subject: "Suspicious email: " + item.subject,
    htmlBody: `
      <p>This message was reported as phishing.</p>
      <p><strong>From:</strong> ${item.from?.emailAddress || 'Unknown'}</p>
      <p><strong>Subject:</strong> ${item.subject}</p>
      <p>Please forward the original email as an attachment for full headers.</p>
    `
  });

  event.completed();
}

Office.actions.associate("reportPhishing", reportPhishing);
