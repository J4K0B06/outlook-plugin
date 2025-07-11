Office.onReady(() => {
  // Office is ready
});

function reportPhishing(event) {
  console.log("reportPhishing triggered");
  const item = Office.context.mailbox.item;
  if (!item) {
    console.error("❌ No item selected or Reading Pane is disabled.");
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["info@safebyte.be"],
      subject: "Suspicious email",
      htmlBody: `<p>⚠️ No item was selected when report was triggered.</p>`
    });
  } else {
    console.log("Item subject:", item.subject);
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["info@safebyte.be"],
      subject: "Suspicious email: " + item.subject,
      htmlBody: `
        <p>This message was reported as phishing.</p>
        <p><strong>From:</strong> ${item.from?.emailAddress}</p>
        <p><strong>Subject:</strong> ${item.subject}</p>
        <p>Body: (please forward the original message MANUALLY)</p>
      `
    });
  }

  event.completed();
}

Office.actions.associate("reportPhishing", reportPhishing);
