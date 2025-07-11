Office.onReady(() => {
  document.getElementById("sendReport").addEventListener("click", async () => {
    const item = Office.context.mailbox.item;
    const subject = item && item.subject ? item.subject : "Suspicious email";

    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["info@safebyte.be"],
      subject: "Suspicious email: " + subject,
      htmlBody: `
        <p>This message was reported as phishing.</p>
        <p><strong>From:</strong> ${item?.from?.emailAddress}</p>
        <p><strong>Subject:</strong> ${subject}</p>
        <p>Body: (please forward the original message manually if needed)</p>
      `
    });
  });
});
