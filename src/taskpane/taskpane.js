Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("logStyleContentButton").onclick = getListInfoFromSelection;
  }
});

async function getListInfoFromSelection() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const selectionRange = selection.getRange();
      const paragraphs = selectionRange.paragraphs;
      paragraphs.load("items");
      await context.sync();

      let clipboardData = [];
      let parentNumbering = [];
      let paragraphCounter = 1;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,style,isListItem");
        await context.sync();

        let text = paragraph.text.trim();
        const isListItem = paragraph.isListItem;

        // Remove non-printable characters
        text = text.replace(/[^\x20-\x7E]/g, "");

        // Remove a leading dot if present
        if (text.startsWith(".")) {
          text = text.substring(1).trim();
        }

        if (text.length <= 1) {
          continue;
        }

        if (isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

          if (level <= parentNumbering.length) {
            parentNumbering = parentNumbering.slice(0, level);
          }
          parentNumbering[level] = listString;

          const fullNumbering = parentNumbering.slice(0, level + 1).join(".");

          clipboardData.push(`"${fullNumbering}": "${text}"`);
        } else {
          const parentKey = parentNumbering.length > 0 ? parentNumbering.join(".") : `paragraph_${paragraphCounter}`;
          clipboardData.push(`"${parentKey}.text": "${text}"`);
          paragraphCounter++;
        }
      }

      const clipboardString = `{\n${clipboardData.join(",\n")}\n}`;
      copyToClipboard(clipboardString);

      console.log("All data copied to clipboard:");
      console.log(clipboardString);
    });
  } catch (error) {
    console.error("An error occurred:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

function copyToClipboard(text) {
  const textArea = document.createElement("textarea");
  textArea.value = text;

  textArea.style.position = "fixed";
  textArea.style.left = "-999999px";
  textArea.style.top = "-999999px";
  document.body.appendChild(textArea);

  textArea.focus();
  textArea.select();

  try {
    const successful = document.execCommand("copy");
    const msg = successful ? "successful" : "unsuccessful";
    console.log("Copying text was " + msg);

    if (successful) {
      const copyMessage = document.getElementById("copyMessage");
      copyMessage.style.display = "block";

      setTimeout(() => {
        copyMessage.style.display = "none";
      }, 15000);
    }
  } catch (err) {
    console.error("Unable to copy to clipboard", err);
  }

  document.body.removeChild(textArea);
}
