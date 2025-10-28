async function onNewMessageCompose(event) {
  try {
    // URL unde ții semnătura HTML (poate fi Azure Blob / SharePoint)
    const signatureUrl = "https://raw.githubusercontent.com/cosminvizitiuab/cosminvizitiuab.github.io/refs/heads/main/semnatura%20Cosmin%20Vizitiu%20cu%20base64.html";

    const resp = await fetch(signatureUrl, { cache: "no-cache" });
    if (!resp.ok) throw new Error("Nu am putut prelua semnatura: " + resp.status);
    const signatureHtml = await resp.text();

    // Inserează semnătura în corpul mesajului
    Office.context.mailbox.item.body.setAsync(
      signatureHtml,
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("Eroare setAsync:", asyncResult.error);
        }
        event.completed();
      }
    );
  } catch (err) {
    console.error("Eroare onNewMessageCompose:", err);
    event.completed();
  }
}

// Export pentru runtime
if (typeof module !== "undefined") {
  module.exports = { onNewMessageCompose };
}
