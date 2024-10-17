let clipboardData = [];
let parentNumbering = [];
let lastParentKey = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("logStyleContentButton").onclick = getListInfoFromSelection;
    document.getElementById("clearContentButton").onclick = clearCopiedContent;
  }
});

async function getListInfoFromSelection() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      let paragraphCounter = 1;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,style,isListItem");
        await context.sync();

        let text = paragraph.text.trim();
        const isListItem = paragraph.isListItem;

        text = text.replace(/[^\x20-\x7E]/g, "");

        if (text.length <= 1) {
          continue;
        }

        if (isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

          // Adjust the parentNumbering array based on the current level
          if (level <= parentNumbering.length) {
            parentNumbering = parentNumbering.slice(0, level);
          }

          parentNumbering[level] = listString;

          let fullNumbering = "";
          for (let j = 0; j <= level; j++) {
            if (parentNumbering[j]) {
              fullNumbering += `${parentNumbering[j]}.`;
            }
          }

          fullNumbering = fullNumbering.replace(/\.$/, "");
          clipboardData.push({ key: fullNumbering, value: text });

          lastParentKey = fullNumbering;
        } else {
          const parentKey = lastParentKey || `paragraph_${paragraphCounter}`;
          clipboardData.push({ key: parentKey + ".text", value: text });
          paragraphCounter++;
        }
      }

      updateCopiedContentDisplay();

      const clipboardString = formatClipboardData();
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

function formatClipboardData() {
  return `{\n${clipboardData.map((pair) => `"${pair.key}": "${pair.value}"`).join(",\n")}\n}`;
}

function updateCopiedContentDisplay() {
  const copiedContentElement = document.getElementById("copiedContent");

  copiedContentElement.innerHTML = "";

  clipboardData.forEach((pair) => {
    const keySpan = `<span class="key">${pair.key}</span>`;
    const valueSpan = `<span class="value">${pair.value}</span>`;

    const formattedPair = `<div class="pair">${keySpan}: ${valueSpan}</div>`;
    copiedContentElement.innerHTML += formattedPair;
  });

  copiedContentElement.scrollTop = copiedContentElement.scrollHeight;
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

function clearCopiedContent() {
  clipboardData = [];
  parentNumbering = [];
  lastParentKey = "";
  document.getElementById("copiedContent").innerText = "";
}
