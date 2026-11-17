const PREFIX = "ICT";        // קידומת הקבצים לסימון
const LOGO_URL = "logo.png"; // הלוגו (PNG, עדיף שקוף), באותה תיקייה
const LOGO_WIDTH_RATIO = 0.6; // רוחב הלוגו ביחס לרוחב העמוד
const LOGO_OPACITY = 0.25;    // שקיפות (0-1)

Office.onReady();

const call = (fn) => new Promise((ok, fail) =>
  fn(r => r.status === Office.AsyncResultStatus.Succeeded ? ok(r.value) : fail(r.error)));

function notify(item, message, type) {
  item.notificationMessages.replaceAsync("ictwm", {
    type, message: message.slice(0, 150), icon: "Icon.16x16", persistent: false
  });
}

async function addLogo(b64pdf, logoBytes) {
  const { PDFDocument } = PDFLib;
  const pdf = await PDFDocument.load(b64pdf);
  const logo = await pdf.embedPng(logoBytes);
  for (const page of pdf.getPages()) {
    const { width, height } = page.getSize();
    const w = width * LOGO_WIDTH_RATIO;
    const h = w * (logo.height / logo.width);
    page.drawImage(logo, {
      x: (width - w) / 2, y: (height - h) / 2,
      width: w, height: h, opacity: LOGO_OPACITY
    });
  }
  return await pdf.saveAsBase64();
}

async function watermarkIct(event) {
  const item = Office.context.mailbox.item;
  const T = Office.MailboxEnums.ItemNotificationMessageType;
  try {
    const atts = await call(cb => item.getAttachmentsAsync(cb));
    const targets = atts.filter(a =>
      a.attachmentType === Office.MailboxEnums.AttachmentType.File &&
      a.name.toUpperCase().startsWith(PREFIX) &&
      a.name.toLowerCase().endsWith(".pdf"));

    if (!targets.length) {
      notify(item, `לא נמצאו נספחי PDF שמתחילים ב-${PREFIX}`, T.InformationalMessage);
      return event.completed();
    }

    const logoBytes = await (await fetch(LOGO_URL)).arrayBuffer();
    let count = 0;
    for (const att of targets) {
      const content = await call(cb => item.getAttachmentContentAsync(att.id, cb));
      const marked = await addLogo(content.content, new Uint8Array(logoBytes));
      await call(cb => item.removeAttachmentAsync(att.id, cb));
      await call(cb => item.addFileAttachmentFromBase64Async(marked, att.name, cb));
      count++;
    }
    notify(item, `סומנו ${count} נספחים. אל תלחץ שוב כדי לא לסמן פעמיים.`, T.InformationalMessage);
  } catch (e) {
    notify(item, "הסימון נכשל, בדוק את הנספחים לפני שליחה", T.ErrorMessage);
  }
  event.completed();
}

Office.actions.associate("watermarkIct", watermarkIct);
