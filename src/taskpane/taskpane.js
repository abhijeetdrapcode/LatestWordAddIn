let clipboardData = [];
let allParagraphsData = [];
let isDataLoaded = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const logStyleContentButton = document.getElementById("logStyleContentButton");
    logStyleContentButton.disabled = true;
    logStyleContentButton.onclick = getListInfoFromSelection;
    document.getElementById("clearContentButton").onclick = clearCopiedContent;
    loadAllParagraphsData();
  }
});

function normalizeText(text) {
  return text
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[^\x20-\x7E]/g, "");
}

async function loadAllParagraphsData() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      allParagraphsData = [];
      let parentNumbering = [];
      let lastNumbering = "";

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,isListItem");
        await context.sync();

        let text = normalizeText(paragraph.text);

        if (text.length <= 1) {
          continue;
        }

        if (paragraph.isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

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
          lastNumbering = fullNumbering;

          allParagraphsData.push({
            key: fullNumbering,
            value: text,
            originalText: paragraph.text.trim(),
            isListItem: true,
          });
        } else {
          // Non-numbered paragraphs get `(text)` appended to the last valid numbering
          const key = lastNumbering ? `${lastNumbering} (text)` : `text_${i + 1}`;
          allParagraphsData.push({
            key: key,
            value: text,
            originalText: paragraph.text.trim(),
            isListItem: false,
          });
        }
      }

      allParagraphsData.forEach((item) => {
        if (item.key.endsWith(".text")) {
          allParagraphsData.push({
            key: item.key.replace(".text", ""),
            value: item.value,
            originalText: item.originalText,
            isListItem: false,
          });
        }
      });

      allParagraphsData = allParagraphsData.filter((item) => !item.key.endsWith(".text"));

      allParagraphsData.sort((a, b) => {
        const aMatch = a.key.match(/(\d+)(?=\.)/);
        const bMatch = b.key.match(/(\d+)(?=\.)/);

        if (aMatch && bMatch) {
          return parseInt(aMatch[0]) - parseInt(bMatch[0]);
        }
        return a.key.localeCompare(b.key);
      });

      console.log("All paragraphs data loaded:", allParagraphsData);
      document.getElementById("logStyleContentButton").disabled = false;
      isDataLoaded = true;
    });
  } catch (error) {
    console.error("An error occurred while loading all paragraphs data:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
    document.getElementById("logStyleContentButton").disabled = false;
  }
}

async function getListInfoFromSelection() {
  if (!isDataLoaded) {
    console.log("Data is still loading. Please wait.");
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      let newSelections = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const selectedParagraph = paragraphs.items[i];
        selectedParagraph.load("text");
        await context.sync();

        const selectedText = selectedParagraph.text.trim();
        const normalizedSelectedText = normalizeText(selectedText);

        const matchingParagraphData = allParagraphsData.find(
          (para) => para.value === normalizedSelectedText || para.originalText === selectedText
        );

        if (matchingParagraphData) {
          const isDuplicate = clipboardData.some(
            (item) => item.key === matchingParagraphData.key && item.value === matchingParagraphData.value
          );

          if (!isDuplicate) {
            newSelections.push({
              key: matchingParagraphData.key,
              value: matchingParagraphData.value,
            });
          }
        } else {
          console.log("No match found for:", selectedText);
        }
      }

      if (newSelections.length > 0) {
        clipboardData = [...clipboardData, ...newSelections];

        clipboardData.sort((a, b) => {
          const aMatch = a.key.match(/\d+/g);
          const bMatch = b.key.match(/\d+/g);

          if (aMatch && bMatch) {
            return parseInt(aMatch[0]) - parseInt(bMatch[0]);
          }
          return a.key.localeCompare(b.key);
        });

        updateCopiedContentDisplay();
        const clipboardString = formatClipboardData();
        await copyToClipboard(clipboardString);

        console.log("Updated clipboard data:", clipboardString);
      } else {
        console.log("No new paragraphs to add.");
      }
    });
  } catch (error) {
    console.error("An error occurred while copying data:", error);
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

  copiedContentElement.style.display = clipboardData.length > 0 ? "block" : "none";
  copiedContentElement.scrollTop = copiedContentElement.scrollHeight;
}

async function copyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    showCopyMessage(true);
  } catch (err) {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.position = "fixed";
    textArea.style.left = "-999999px";
    textArea.style.top = "-999999px";
    document.body.appendChild(textArea);

    try {
      textArea.focus();
      textArea.select();
      const successful = document.execCommand("copy");
      showCopyMessage(successful);
    } catch (err) {
      console.error("Unable to copy to clipboard", err);
      showCopyMessage(false);
    } finally {
      document.body.removeChild(textArea);
    }
  }
}

function showCopyMessage(successful) {
  const copyMessage = document.getElementById("copyMessage");
  copyMessage.style.display = "block";
  copyMessage.textContent = successful ? "Content added and copied to clipboard!" : "Failed to copy content";
  copyMessage.style.color = successful ? "green" : "red";

  setTimeout(() => {
    copyMessage.style.display = "none";
  }, 3000);
}

function clearCopiedContent() {
  clipboardData = [];
  const copiedContentElement = document.getElementById("copiedContent");
  copiedContentElement.innerHTML = "";
  copiedContentElement.style.display = "none";
}
