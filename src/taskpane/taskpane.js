/* eslint-disable office-addins/load-object-before-read */
/* eslint-disable office-addins/call-sync-before-read */
/* eslint-disable @typescript-eslint/no-unused-vars */

let categoryData = {
  closing: [],
  postClosing: [],
  representation: [],
};

window.categoryData = categoryData;

let allParagraphsData = [];
let isDataLoaded = false;
let documentContentHash = "";

// Initialize Word operations when Office is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    loadSavedCategoryData();
    initWordOperations();
  }
});

function initWordOperations() {
  const logStyleContentButton = document.getElementById("logStyleContentButton");
  const categorySelect = document.getElementById("categorySelect");
  const reloadButton = document.getElementById("reloadButton");
  const dismissButton = document.getElementById("dismissNotification");

  logStyleContentButton.disabled = true;
  logStyleContentButton.onclick = getListInfoFromSelection;
  document.getElementById("clearContentButton").onclick = clearCurrentContent;
  reloadButton.onclick = handleReloadContent;

  if (dismissButton) {
    dismissButton.onclick = dismissChangeNotification;
  }

  categorySelect.onchange = handleCategoryChange;
  handleCategoryChange();

  setInitialContentHash();
  setInterval(checkForDocumentChanges, 2000);
  loadAllParagraphsData();
}

async function setInitialContentHash() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      documentContentHash = await calculateHash(body.text);
    });
  } catch (error) {
    console.error("Error setting initial content hash:", error);
  }
}

function dismissChangeNotification() {
  const changeNotification = document.getElementById("changeNotification");
  if (changeNotification) {
    changeNotification.style.display = "none";
  }
}

async function calculateHash(text) {
  const encoder = new TextEncoder();
  const data = encoder.encode(text);
  const hashBuffer = await crypto.subtle.digest("SHA-256", data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function checkForDocumentChanges() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      const currentHash = await calculateHash(body.text);

      if (currentHash !== documentContentHash) {
        documentContentHash = currentHash;
        const changeNotification = document.getElementById("changeNotification");
        if (changeNotification) {
          changeNotification.style.display = "block";
        }
      }
    });
  } catch (error) {
    console.error("Error checking for document changes:", error);
  }
}

async function handleReloadContent() {
  const changeNotification = document.getElementById("changeNotification");
  if (changeNotification) {
    changeNotification.style.display = "none";
  }
  await setInitialContentHash();
  await loadAllParagraphsData();
}

async function handleCategoryChange() {
  const categorySelect = document.getElementById("categorySelect");
  const selectedCategory = categorySelect.value;

  document.querySelectorAll(".category-content").forEach((section) => {
    section.classList.remove("active");
  });

  const contentId = `${selectedCategory}Content`;
  document.getElementById(contentId).classList.add("active");

  document.getElementById("logStyleContentButton").disabled = !isDataLoaded || !selectedCategory;

  if (selectedCategory && categoryData[selectedCategory]) {
    // const clipboardString = formatCategoryData(selectedCategory);
    // await silentCopyToClipboard(clipboardString);
  }
}

function normalizeText(text) {
  if (!text) return "";
  return text
    .trim()
    .replace(/^\.\s*/, "")
    .replace(/\s+/g, " ")
    .replace(/[^\x20-\x7E]/g, "");
  // .toLowerCase();
}

async function loadAllParagraphsData() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;

      paragraphs.load("items, items/text, items/isListItem");
      await context.sync();

      let parentNumbering = [];
      let lastNumbering = "";
      const today = new Date().toISOString().split("T")[0];

      document.getElementById("logStyleContentButton").disabled = true;
      isDataLoaded = false;

      const listItems = paragraphs.items.filter((p) => p.isListItem);
      listItems.forEach((item) => item.listItem.load("level, listString"));
      await context.sync();

      allParagraphsData = []; // Reset the data before reloading

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        const text = normalizeText(paragraph.text);

        if (text.length <= 1) continue;

        let existingItem = allParagraphsData.find((item) => item.index === i);
        let revisedKey = "";

        if (paragraph.isListItem) {
          const listItem = paragraph.listItem;
          const level = listItem.level;
          const listString = listItem.listString || "";

          if (level <= parentNumbering.length) {
            parentNumbering = parentNumbering.slice(0, level);
          }
          parentNumbering[level] = listString;

          const fullNumbering = parentNumbering
            .slice(0, level + 1)
            .filter(Boolean)
            .join(".");
          lastNumbering = fullNumbering;

          if (existingItem && existingItem.value !== text) {
            revisedKey = `${fullNumbering} [Revised ${today}]`;
          } else {
            revisedKey = fullNumbering;
          }

          allParagraphsData.push({
            key: revisedKey,
            value: text,
            originalText: paragraph.text.trim().replace(/^\.\s*/, ""),
            isListItem: true,
            index: i,
            level: level,
            listString: listString,
            parentNumbers: [...parentNumbering],
          });
        } else {
          const baseKey = lastNumbering ? `${lastNumbering} (text)` : `text_${i + 1}`;

          if (existingItem && existingItem.value !== text) {
            revisedKey = `${baseKey} [Revised ${today}]`;
          } else {
            revisedKey = baseKey;
          }

          allParagraphsData.push({
            key: revisedKey,
            value: text,
            originalText: paragraph.text.trim().replace(/^\.\s*/, ""),
            isListItem: false,
            index: i,
            level: -1,
          });
        }
      }

      allParagraphsData = allParagraphsData.filter((item) => !item.key.endsWith(".text"));

      const categorySelect = document.getElementById("categorySelect");
      document.getElementById("logStyleContentButton").disabled = !categorySelect.value;
      isDataLoaded = true;
      console.log("All paragraphs data loaded:", allParagraphsData.length);
    });
  } catch (error) {
    console.error("An error occurred while loading all paragraphs data:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
    document.getElementById("logStyleContentButton").disabled = true;
    isDataLoaded = false;
  }
}

async function getListInfoFromSelection() {
  if (!isDataLoaded) {
    console.log("Data is still loading. Please wait.");
    showCopyMessage(false, "Data is still loading. Please wait.");
    return;
  }

  const selectedCategory = document.getElementById("categorySelect").value;
  console.log("Selected Category:", selectedCategory);

  if (!selectedCategory) {
    console.log("No category selected");
    showCopyMessage(false, "No category selected");
    return;
  }

  try {
    await Word.run(async (context) => {
      // 1. Get current data from localStorage
      const savedData = JSON.parse(localStorage.getItem("categoryData") || "{}");
      const currentCategoryItems = savedData[selectedCategory] || [];

      // 2. Process Word selection
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      console.log("Selected paragraphs count:", paragraphs.items.length);
      if (paragraphs.items.length === 0) {
        console.log("No paragraphs found in selection");
        showCopyMessage(false, "No paragraphs found in selection");
        return;
      }

      // Load paragraph details
      const paragraphPromises = paragraphs.items.map((paragraph) => {
        paragraph.load("text,isListItem");
        if (paragraph.isListItem) {
          paragraph.listItem.load("level,listString");
        }
        return paragraph;
      });
      await context.sync();

      let newSelections = [];

      // 3. Match selected paragraphs with document data
      for (const paragraph of paragraphs.items) {
        const selectedText = paragraph.text.trim().replace(/^\.\s*/, "");
        const normalizedSelectedText = normalizeText(selectedText);

        const matchingParagraphs = allParagraphsData.filter((para) => {
          const normalizedParaText = normalizeText(para.originalText || para.value);
          return (
            normalizedParaText === normalizedSelectedText ||
            para.originalText === selectedText ||
            para.value === normalizedSelectedText
          );
        });

        if (matchingParagraphs.length > 0) {
          let bestMatch = matchingParagraphs[0];

          // Handle list items more precisely
          if (matchingParagraphs.length > 1 && paragraph.isListItem) {
            const selectedLevel = paragraph.listItem.level;
            const selectedListString = paragraph.listItem.listString;

            const exactMatch = matchingParagraphs.find(
              (para) => para.isListItem && para.level === selectedLevel && para.listString === selectedListString
            );

            if (exactMatch) bestMatch = exactMatch;
          }

          // Check against localStorage data for duplicates
          const isDuplicate = currentCategoryItems.some(
            (item) => (item.key === bestMatch.key && item.value === bestMatch.value) || item.content === bestMatch.value // For closing/post-closing items
          );

          if (!isDuplicate) {
            if (selectedCategory === "closing" || selectedCategory === "postClosing") {
              if (bestMatch.key) {
                const keyParts = bestMatch.key.split(/(?<=^[^\d]+)(?=\d)/);
                const mainHeadingKey = keyParts[0].trim().replace(/\.$/, "");
                const sectionHeading = bestMatch.key.trim();
                const content = bestMatch.value.trim();

                const matchedParagraph = allParagraphsData.find((para) => para.key.trim() === mainHeadingKey);
                const fullMainHeading = matchedParagraph
                  ? mainHeadingKey + " " + matchedParagraph.value
                  : mainHeadingKey;

                newSelections.push({
                  mainHeading: fullMainHeading,
                  sectionHeading: sectionHeading,
                  content: content,
                });
              }
            } else {
              newSelections.push({
                key: bestMatch.key,
                value: bestMatch.value,
              });
            }
          }
        }
      }

      // 4. Update localStorage if new items found
      if (newSelections.length > 0) {
        const updatedData = {
          ...savedData,
          [selectedCategory]: [...currentCategoryItems, ...newSelections].sort((a, b) => {
            const aNumbers = a.key ? a.key.split(".").map(Number) : [];
            const bNumbers = b.key ? b.key.split(".").map(Number) : [];

            for (let i = 0; i < Math.max(aNumbers.length, bNumbers.length); i++) {
              if (isNaN(aNumbers[i])) return 1;
              if (isNaN(bNumbers[i])) return -1;
              if (aNumbers[i] !== bNumbers[i]) return aNumbers[i] - bNumbers[i];
            }
            return 0;
          }),
        };

        localStorage.setItem("categoryData", JSON.stringify(updatedData));
        updateCategoryDisplay(selectedCategory);
        handleCategoryChange();
        showCopyMessage(true, "Content added successfully!");
      } else {
        console.log("No new selections to add");
        showCopyMessage(false, "No new matching content found");
      }
    });
  } catch (error) {
    console.error("Selection processing failed:", error);
    showCopyMessage(false, "Error: " + error.message);
  }
}

function formatCategoryData(category) {
  if (!categoryData[category] || categoryData[category].length === 0) {
    console.error("No data available for category:", category);
    return "{}";
  }

  try {
    if (category === "closing" || category === "postClosing") {
      return formatClosingChecklistData(category);
    }

    const pairs = categoryData[category]
      .filter((item) => item.key && item.value) // Filter out invalid entries
      .map((pair) => `"${pair.key}": "${pair.value.replace(/"/g, '\\"')}"`)
      .join(",\n");

    return pairs ? `{\n${pairs}\n}` : "{}";
  } catch (error) {
    console.error("Formatting error:", error);
    return "{}";
  }
}
function formatClosingChecklistData(data, selectedCategory) {
  if (!data || !data[selectedCategory]) {
    console.error("Invalid or empty data for category:", selectedCategory);
    return "{}";
  }

  const selections = data[selectedCategory];
  if (!Array.isArray(selections) || selections.length === 0) {
    console.error("Invalid or empty selections data for category:", selectedCategory);
    return "{}";
  }

  const formattedData = {};

  selections.forEach((selection) => {
    if (!selection.mainHeading || !selection.sectionHeading || !selection.content) {
      console.error("Missing data in selection:", selection);
      return;
    }

    const mainHeading = selection.mainHeading.trim();
    const sectionHeading = selection.sectionHeading.trim();
    const content = selection.content.trim();

    if (!formattedData[mainHeading]) {
      formattedData[mainHeading] = {
        title: mainHeading,
        sections: [],
      };
    }

    formattedData[mainHeading].sections.push({
      sectionHeading: sectionHeading,
      content: content,
    });
  });

  return JSON.stringify(formattedData, null, 2);
}
function updateCategoryDisplay(category) {
  // Get data DIRECTLY from localStorage (only change needed)
  const displayData = JSON.parse(localStorage.getItem("categoryData") || {})[category] || [];

  // Rest of the original function remains EXACTLY the same
  const contentElement = document.querySelector(`#${category}Content .content-area`);
  if (!contentElement) {
    console.error("Content element not found for category:", category);
    return;
  }

  contentElement.innerHTML = "";

  if (displayData.length > 0) {
    displayData.forEach((pair) => {
      if (category === "closing" || category === "postClosing") {
        const entries = [
          { key: "Article", value: pair.mainHeading },
          { key: "Section", value: pair.sectionHeading },
          { key: "Clause", value: pair.content },
        ];

        entries.forEach((entry) => {
          const keySpan = `<span class="key">${entry.key}</span>`;
          const valueSpan = `<span class="value">${entry.value}</span>`;
          const formattedPair = `<div class="pair">${keySpan}: ${valueSpan}</div>`;
          contentElement.innerHTML += formattedPair;
        });

        contentElement.innerHTML += "<br><br>";
      } else {
        const keySpan = `<span class="key">${pair.key}</span>`;
        const valueSpan = `<span class="value">${pair.value}</span>`;
        const formattedPair = `<div class="pair">${keySpan}: ${valueSpan}</div>`;
        contentElement.innerHTML += formattedPair;
      }
    });
  } else {
    contentElement.innerHTML = "<p>No content available for this category</p>";
  }
}

function showCopyMessage(successful, message) {
  const copyMessage = document.getElementById("copyMessage");
  if (!copyMessage) {
    console.error("Copy message element not found");
    return;
  }

  copyMessage.style.display = "block";
  copyMessage.textContent =
    message || (successful ? "Content added and copied to clipboard!" : "Failed to copy content");
  copyMessage.style.color = successful ? "green" : "red";

  setTimeout(() => {
    copyMessage.style.display = "none";
  }, 3000);
}

async function clearCurrentContent() {
  const selectedCategory = document.getElementById("categorySelect").value;
  if (!selectedCategory) {
    console.log("No category selected");
    return;
  }

  categoryData[selectedCategory] = [];
  saveCategoryData();
  const contentElement = document.querySelector(`#${selectedCategory}Content .content-area`);
  if (contentElement) {
    contentElement.innerHTML = "<p>No content available for this category</p>";
  }

  const clipboardString = "{}";
  // await silentCopyToClipboard(clipboardString);

  console.log(`Cleared content for category: ${selectedCategory}`);
  showCopyMessage(true, "Category content cleared");
}

function saveCategoryData() {
  const dataToSave = {
    closing: Array.isArray(categoryData.closing) ? categoryData.closing : [],
    postClosing: Array.isArray(categoryData.postClosing) ? categoryData.postClosing : [],
    representation: Array.isArray(categoryData.representation) ? categoryData.representation : [],
  };

  try {
    localStorage.setItem("categoryData", JSON.stringify(dataToSave));
    console.log("Saved structured categoryData to localStorage");
  } catch (error) {
    console.error("Failed to save categoryData:", error);
  }
}

// Helper to load saved data on startup
function loadSavedCategoryData() {
  const defaultStructure = {
    closing: [],
    postClosing: [],
    representation: [],
  };

  try {
    const saved = JSON.parse(localStorage.getItem("categoryData") || "{}");

    // Fix missing keys and wrong structure
    categoryData = {
      closing: Array.isArray(saved.closing) ? saved.closing : [],
      postClosing: Array.isArray(saved.postClosing) ? saved.postClosing : [],
      representation: Array.isArray(saved.representation) ? saved.representation : [],
    };

    // Save back the cleaned-up version
    localStorage.setItem("categoryData", JSON.stringify(categoryData));

    console.log("Loaded and corrected categoryData:", categoryData);
  } catch (error) {
    console.error("Failed to load/parse categoryData:", error);
    categoryData = { ...defaultStructure };
    localStorage.setItem("categoryData", JSON.stringify(categoryData));
  }
}
