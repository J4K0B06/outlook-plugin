Office.onReady(() => {
  document.getElementById("sendReport").addEventListener("click", async () => {
    const item = Office.context.mailbox.item;

    item.displayForwardForm({
      toRecipients: ["info@safebyte.be"],
      htmlBody: `
        <p>This message was reported as phishing.</p>
        <p><strong>Please investigate.</strong></p>
        <br/>
      `
    });
  });
});
