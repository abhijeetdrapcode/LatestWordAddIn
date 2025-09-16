console.log("dealDriver.js loaded");

// let selectedCategory = "";
// let selectedCategory = localStorage.getItem("selectedCategory") || "repsAndWarranty";
let selectedCategory = "representation";
console.log("This is the selectedCategory: ", selectedCategory);
window.selectedCategory = selectedCategory;
let isLoggedIn = false;
let loginResponseData = null;
let selectedEnvironment;
let dealSelect, sendDealButton, fetchDataButton;

// Initialize Deal Driver integration when Office is ready
document.addEventListener("DOMContentLoaded", function () {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      initDealDriverIntegration();
    }
  });
});

function initDealDriverIntegration() {
  // Only get elements that are always visible
  const loginButton = document.getElementById("loginButton");
  const loginModal = document.getElementById("loginModal");
  const loginForm = document.getElementById("loginForm");
  const closeModal = document.querySelector(".close-modal");
  const loginError = document.getElementById("loginError");
  const dealOptions = document.getElementById("dealOptions");

  if (!loginButton || !loginModal || !loginForm || !closeModal || !loginError || !dealOptions) {
    console.error("One or more required elements not found in DOM");
    return;
  }

  // Login button click handler
  loginButton.addEventListener("click", () => {
    if (!isLoggedIn) {
      loginModal.style.display = "block";
    } else {
      // Handle logout
      handleLogout();
    }
  });

  const categorySelect = document.getElementById("categorySelect");
  if (categorySelect) {
    // Set initial value to repsAndWarranty
    categorySelect.value = selectedCategory;
    localStorage.setItem("selectedCategory", selectedCategory); // Store default

    categorySelect.addEventListener("change", (e) => {
      selectedCategory = e.target.value;
      window.selectedCategory = selectedCategory;
      localStorage.setItem("selectedCategory", selectedCategory);
      console.log(`Category changed to: ${selectedCategory}`);
    });
  }

  // Close modal handlers
  closeModal.addEventListener("click", () => {
    loginModal.style.display = "none";
    loginError.style.display = "none";
  });

  window.addEventListener("click", (event) => {
    if (event.target === loginModal) {
      loginModal.style.display = "none";
      loginError.style.display = "none";
    }
  });

  // Login form submission handler
  loginForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const userName = document.getElementById("userName").value;
    const password = document.getElementById("password").value;

    const loginSuccess = await handleLogin(userName, password);
    if (loginSuccess) {
      isLoggedIn = true;
      loginButton.textContent = "Logout";
      loginModal.style.display = "none";
      loginError.style.display = "none";
      loginForm.reset();

      // Show the deal options dropdown and button
      dealOptions.style.display = "block";
      const categoryData = {
        closing: [],
        postClosing: [],
        representation: [],
      };

      localStorage.setItem("categoryData", JSON.stringify(categoryData));

      // Initialize post-login elements and event listeners
      initPostLoginElements();
    } else {
      loginError.style.display = "block";
    }
  });
}

function initPostLoginElements() {
  // Now that user is logged in, get references to elements that were hidden
  dealSelect = document.getElementById("dealSelect");
  fetchDataButton = document.getElementById("fetchDataButton");
  sendDealButton = document.getElementById("sendDealButton");

  // Define the send deal handler function
  sendDealButtonHandler = async () => {
    // Use the existing isApiCallInProgress flag from your handleSendDeal function
    if (isApiCallInProgress) return;

    try {
      await handleSendDeal();
    } catch (error) {
      console.error("Error in sendDealButtonHandler:", error);
    }
  };

  // Define the fetch data handler function
  fetchDataButtonHandler = async () => {
    // Create local isFetching flag for this specific handler
    if (fetchDataButtonHandler.isFetching) return;
    fetchDataButtonHandler.isFetching = true;

    if (fetchDataButton) {
      fetchDataButton.disabled = true;
    }

    try {
      let categoryFromSelection = localStorage.getItem("selectedCategory");
      if (categoryFromSelection === "representation") {
        await fetchRepresentationData();
      } else if (categoryFromSelection === "closing") {
        await fetchClosingData();
      }
    } catch (error) {
      console.error("Error fetching data:", error);
    } finally {
      fetchDataButtonHandler.isFetching = false;
      if (fetchDataButton) {
        fetchDataButton.disabled = false;
      }
    }
  };

  // Remove any existing event listeners before adding new ones
  if (sendDealButton) {
    if (sendDealButtonHandler) {
      sendDealButton.removeEventListener("click", sendDealButtonHandler);
    }
    sendDealButton.addEventListener("click", sendDealButtonHandler);
  }

  if (fetchDataButton) {
    if (fetchDataButtonHandler) {
      fetchDataButton.removeEventListener("click", fetchDataButtonHandler);
    }
    fetchDataButton.addEventListener("click", fetchDataButtonHandler);
  }

  // Show main content
  const mainContent = document.getElementById("mainContent");
  if (mainContent) {
    mainContent.classList.remove("hidden");
  }
}
// Add these variables at the top of your file (global scope)
let sendDealButtonHandler = null;
let fetchDataButtonHandler = null;

// Updated handleLogout function
function handleLogout() {
  // Remove event listeners first, before clearing element references
  if (fetchDataButton && fetchDataButtonHandler) {
    fetchDataButton.removeEventListener("click", fetchDataButtonHandler);
  }
  if (sendDealButton && sendDealButtonHandler) {
    sendDealButton.removeEventListener("click", sendDealButtonHandler);
  }

  // Reset handler references
  fetchDataButtonHandler = null;
  sendDealButtonHandler = null;

  // Reset global state
  isLoggedIn = false;
  loginResponseData = null;
  selectedEnvironment = null;

  // Clear localStorage
  localStorage.clear();

  // Reset category data
  categoryData = {
    closing: [],
    postClosing: [],
    representation: [],
  };

  // Hide and reset UI elements
  document.getElementById("mainContent").classList.add("hidden");
  document.querySelectorAll(".content-area").forEach((el) => {
    el.innerHTML = "<p>No content available</p>";
  });

  // Remove any active classes
  document.querySelectorAll(".category-content").forEach((el) => {
    el.classList.remove("active");
  });

  // Update UI elements
  const loginButton = document.getElementById("loginButton");
  const dealOptions = document.getElementById("dealOptions");
  const mainContent = document.getElementById("mainContent");

  if (loginButton) loginButton.textContent = "Login To Deal Driver";
  if (dealOptions) dealOptions.style.display = "none";
  if (mainContent) mainContent.classList.add("hidden");

  // Clear references to post-login elements
  dealSelect = null;
  sendDealButton = null;
  fetchDataButton = null;
}
async function handleLogin(userName, password) {
  try {
    const environmentSelect = document.getElementById("environmentSelect");
    selectedEnvironment = environmentSelect.value;
    localStorage.setItem("selectedEnvironment", selectedEnvironment);
    let categorySelected = "representation";
    localStorage.setItem("selectedCategory", "representation");

    const apiUrl =
      selectedEnvironment === "production"
        ? "https://deal-driver-20245869.api.drapcode.io/api/v1/developer/login"
        : selectedEnvironment === "preview"
          ? "https://deal-driver-20245869.api.preview.drapcode.io/api/v1/developer/login"
          : "https://deal-driver-20245869.api.sandbox.drapcode.io/api/v1/developer/login";

    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        userName,
        password,
      }),
    });

    if (response.ok) {
      const data = await response.json();
      localStorage.setItem("loginResponseData", JSON.stringify(data));
      loginResponseData = data;
      // localStorage.setItem("authToken", data.token);

      const dealNames = data.userDetails?.tenantId || [];
      populateDealDropdown(dealNames);

      if (dealNames.length > 0) {
        document.getElementById("dealOptions").style.display = "block";
      } else {
        document.getElementById("dealOptions").style.display = "none";
      }

      return true;
    } else {
      console.error("Login failed with status:", response.status);
      return false;
    }
  } catch (error) {
    console.error("Login error:", error);
    return false;
  }
}

function populateDealDropdown(dealNames) {
  const dealSelectElement = document.getElementById("dealSelect");
  if (!dealSelectElement) {
    console.error("Deal select element not found");
    return;
  }

  dealSelectElement.innerHTML = "";

  dealNames.forEach((deal) => {
    const option = document.createElement("option");
    option.value = deal._id;
    option.textContent = deal.name;
    dealSelectElement.appendChild(option);
  });
}

// Add this at the top of your file (global scope)
let isApiCallInProgress = false;

async function handleSendDeal() {
  // 1. Check if API call is already in progress
  if (isApiCallInProgress) {
    console.log("API call already in progress - ignoring duplicate request");
    return;
  }

  if (!dealSelect || !sendDealButton) {
    console.error("Required elements not available");
    return;
  }

  const messageElement = ensureMessageElementExists();
  const showMessage = createMessageHandler(messageElement);

  // 2. Set the flag immediately when starting
  isApiCallInProgress = true;

  try {
    // Disable the send button during processing
    sendDealButton.disabled = true;
    sendDealButton.style.opacity = "0.5";
    sendDealButton.style.cursor = "not-allowed";

    const selectedDealName = dealSelect.options[dealSelect.selectedIndex].text;
    selectedCategory = document.getElementById("categorySelect").value;
    const loginResponseDataString = localStorage.getItem("loginResponseData");

    if (!loginResponseDataString) {
      showMessage("Login data not found", true);
      return;
    }

    const loginResponseData = JSON.parse(loginResponseDataString);
    const dealsArray = loginResponseData.userDetails.tenantId || [];
    const matchedDeal = dealsArray.find((deal) => deal.name === selectedDealName);

    if (!matchedDeal) {
      showMessage("Could not find matching deal", true);
      return;
    }

    const matchedDealSettingUUID = matchedDeal.deal[0].deal_user_setting;
    const baseUrl = getBaseUrlForEnvironment(selectedEnvironment);

    if (!baseUrl) {
      showMessage("Invalid environment selected", true);
      return;
    }

    // 3. Add debug logs to track API calls
    console.log("Starting API call process at:", new Date().toISOString());
    console.log("Selected category:", selectedCategory);

    const dealSettingData = await fetch(
      `${baseUrl}/api/v1/developer/collection/user_setting/item/${matchedDealSettingUUID}`
    );
    const finalData = await dealSettingData.json();

    const permissionArray = finalData.permissions || [];
    const requiredPermissions = [
      "CREATE_POST_CLOSING_CHECKLIST",
      "CREATE_REPRESENTATION_WARRANTY",
      "CREATE_REVISED_CLOSING_CHECKLIST",
    ];
    const areaPermissions = permissionArray.filter((permission) => requiredPermissions.includes(permission));

    const dealUuid = matchedDeal.deal[0].uuid;
    localStorage.setItem("selectedDealId", dealUuid);
    const tenantId = loginResponseData.tenant.uuid;

    // 4. Track which API is being called
    console.log("Preparing to call API for:", selectedCategory);

    if (selectedCategory === "closing") {
      await sendClosingData(
        dealUuid,
        tenantId,
        selectedEnvironment,
        areaPermissions,
        selectedCategory,
        selectedDealName,
        showMessage
      );
    } else if (selectedCategory === "postClosing") {
      await sendPostClosingData(
        dealUuid,
        tenantId,
        selectedEnvironment,
        areaPermissions,
        selectedCategory,
        selectedDealName,
        showMessage
      );
    } else {
      await sendRepresentationData(
        dealUuid,
        tenantId,
        selectedEnvironment,
        areaPermissions,
        selectedCategory,
        selectedDealName,
        showMessage
      );
    }

    console.log("API call completed successfully at:", new Date().toISOString());
  } catch (error) {
    console.error("API call failed:", error);
    showMessage("Error sending deal", true);
  } finally {
    // 5. Always reset the flag and button state
    isApiCallInProgress = false;

    if (sendDealButton) {
      sendDealButton.disabled = false;
      sendDealButton.style.opacity = "1";
      sendDealButton.style.cursor = "pointer";
    }
  }
}

function ensureMessageElementExists() {
  let messageElement = document.getElementById("dealSendMessage");
  if (!messageElement && sendDealButton) {
    messageElement = document.createElement("div");
    messageElement.id = "dealSendMessage";
    messageElement.style.position = "absolute";
    messageElement.style.top = "-50px";
    messageElement.style.left = "0";
    messageElement.style.width = "100%";
    messageElement.style.padding = "10px";
    messageElement.style.textAlign = "center";
    messageElement.style.transition = "top 0.3s ease";
    sendDealButton.parentNode.insertBefore(messageElement, sendDealButton);
  }
  return messageElement;
}

function createMessageHandler(messageElement) {
  return (message, isError = false) => {
    if (!messageElement) return;

    messageElement.textContent = message;
    messageElement.style.backgroundColor = isError ? "#ffdddd" : "#ddffdd";
    messageElement.style.color = isError ? "red" : "green";
    messageElement.style.top = "0";

    setTimeout(() => {
      messageElement.style.top = "-50px";
    }, 9000);
  };
}

function getBaseUrlForEnvironment(env) {
  switch (env) {
    case "sandbox":
      return "https://deal-driver-20245869.api.sandbox.drapcode.io";
    case "preview":
      return "https://deal-driver-20245869.api.preview.drapcode.io";
    case "production":
      return "https://deal-driver-20245869.api.drapcode.io";
    default:
      return null;
  }
}

async function sendClosingData(dealUuid, tenantId, environment, permissions, category, dealName, showMessage) {
  try {
    // Properly parse the localStorage data
    const data = localStorage.getItem("categoryData");
    if (!data) {
      throw new Error("No category data found in localStorage");
    }

    const parsedCategoryData = JSON.parse(data);
    if (!parsedCategoryData[category]) {
      throw new Error(`No data available for category: ${category}`);
    }

    // Format the data correctly
    const formattedData = formatClosingChecklistData(parsedCategoryData, category);
    if (!formattedData || formattedData === "{}") {
      throw new Error("Formatted data is empty");
    }

    // Convert permissions array to comma-separated string if it's an array
    let permissionsHeader = permissions;
    if (Array.isArray(permissions)) {
      permissionsHeader = permissions.join(",");
    }

    const response = await fetch("https://dealdriverapi.drapcode.co/addClosingData", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        dealId: dealUuid,
        tenantId: tenantId,
        environment: environment,
        permissions: permissionsHeader, // Send as comma-separated string
      },
      body: formattedData,
    });

    if (response.ok) {
      const responseData = await response.json();
      showMessage(`${category} data sent successfully to ${dealName}`);
      console.log("Server response:", responseData);
    } else {
      const errorData = await response.text();
      showMessage("Error while sending the data", true);
      console.error(`Failed to send deal. Status: ${response.status}`);
      console.error("Error details:", errorData);
    }
  } catch (error) {
    console.error("Error in sendClosingData:", error);
    showMessage(`Error: ${error.message}`, true);
  }
}

async function sendPostClosingData(dealUuid, tenantId, environment, permissions, category, dealName, showMessage) {
  const response = await fetch("https://dealdriverapi.drapcode.co/addPostClosingData", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      dealId: dealUuid,
      tenantId: tenantId,
      environment: environment,
      permissions: permissions,
    },
    body: formatClosingChecklistData(category),
  });

  if (response.ok) {
    const responseData = await response.json();
    showMessage(`${category} data sent successfully to ${dealName}`);
    console.log("Server response:", responseData);
  } else {
    const errorData = await response.text();
    showMessage("Error while sending the data", true);
    console.error(`Failed to send deal. Status: ${response.status}`);
    console.error("Error details:", errorData);
  }
}

async function sendRepresentationData(dealUuid, tenantId, environment, permissions, category, dealName, showMessage) {
  console.log("THis is the category: ", window.categoryData);
  const repsAndWarrantyData = localStorage.getItem("categoryData");
  console.log("THis is repsWarranty data: ", repsAndWarrantyData);

  try {
    // Parse the JSON string from localStorage
    const parsedData = JSON.parse(repsAndWarrantyData);

    // Check if parsedData exists and has the representation property
    if (!parsedData || !parsedData.representation) {
      throw new Error("Invalid data format: representation array not found");
    }

    const formattedcategoryData = parsedData.representation.reduce((acc, item) => {
      if (item && item.key && item.value) {
        acc[item.key] = item.value;
      }
      return acc;
    }, {});

    console.log("This is the data i am sending : ", formattedcategoryData);
    const response = await fetch("https://dealdriverapi.drapcode.co/parseWord", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        dealId: dealUuid,
        tenantId: tenantId,
        environment: environment,
        permissions: permissions,
      },
      body: JSON.stringify(formattedcategoryData),
    });

    if (response.ok) {
      const responseData = await response.json();
      showMessage(`${category} data sent successfully to ${dealName}`);
      console.log("Server response:", responseData);
    } else {
      const errorData = await response.text();
      showMessage(errorData, true);
      console.error(`Failed to send deal. Status: ${response.status}`);
      console.error("Error details:", errorData);
    }
  } catch (error) {
    console.error("Error processing data:", error);
    showMessage(`Error processing data: ${error.message}`, true);
  }
}
function togglePassword() {
  var toggler = document.getElementById("password");
  if (toggler.type === "password") {
    toggler.type = "text";
  } else {
    toggler.type = "password";
  }
}

// Generic data fetcher for different data types
class DataFetcher {
  constructor(dataType) {
    this.dataType = dataType;
    this.config = this.getConfig(dataType);
  }

  getConfig(dataType) {
    const configs = {
      representation: {
        apiEndpoint: "https://dealdriverapi.drapcode.co/getRepsData",
        dataKey: "representation",
        storageKey: "lastRepresentationApiResponse",
        oldDataKey: "fetchedOldRepresentationData",
        itemKey: "article",
        contentKey: "clause",
        previousKey: "previousClause",
        newKey: "newClause",
      },
      closing: {
        apiEndpoint: "https://dealdriverapi.drapcode.co/getClosingData",
        dataKey: "closing",
        storageKey: "lastClosingApiResponse",
        oldDataKey: "fetchedOldClosingData",
        itemKey: "sectionHeading",
        contentKey: "content",
        previousKey: "previousContent",
        newKey: "newContent",
      },
    };
    return configs[dataType];
  }

  // Helper function to display messages
  showMessage(message, isError = false) {
    console.log(isError ? `Error: ${message}` : `Success: ${message}`);

    const contentArea = document.getElementById("databaseContentArea") || document.body;
    const div = document.createElement("div");
    div.textContent = message;
    div.style.background = isError ? "#ffe6e6" : "#e6ffe6";
    div.style.color = isError ? "red" : "green";
    div.style.padding = "10px";
    div.style.marginBottom = "10px";
    div.style.border = "1px solid #ccc";
    div.style.borderRadius = "5px";
    div.style.fontSize = "14px";
    contentArea.prepend(div);

    setTimeout(() => div.remove(), 4000);
  }

  // Validate required elements and data
  validateRequirements() {
    const dealSelect = document.getElementById("dealSelect");
    if (!dealSelect) {
      throw new Error("Deal selection not found.");
    }

    const selectedDealName = dealSelect.options[dealSelect.selectedIndex]?.text;
    if (!selectedDealName) {
      throw new Error("Please select a deal first.");
    }

    const environment = localStorage.getItem("selectedEnvironment");
    if (!environment) {
      throw new Error("Environment not found. Please login again.");
    }

    const loginResponseDataString = localStorage.getItem("loginResponseData");
    if (!loginResponseDataString) {
      throw new Error("Deal data not found in local storage.");
    }

    return { selectedDealName, environment, loginResponseDataString };
  }

  // Get data from localStorage
  getLocalData() {
    const categoryDataString = localStorage.getItem("categoryData");
    let data = [];

    if (categoryDataString) {
      try {
        const categoryData = JSON.parse(categoryDataString);
        data = categoryData[this.config.dataKey] || [];
        console.log(`Found ${this.dataType} data in localStorage:`, data);
      } catch (e) {
        console.error(`Error parsing categoryData for ${this.dataType}:`, e);
      }
    }

    return data;
  }

  // Find deal UUID from login data
  getDealUuid(loginResponseDataString, selectedDealName) {
    const loginResponseData = JSON.parse(loginResponseDataString);
    const dealsArray = loginResponseData.userDetails?.tenantId || [];
    const matchedDeal = dealsArray.find((deal) => deal.name === selectedDealName);

    if (!matchedDeal) {
      throw new Error("Could not find matching deal for the selected name.");
    }

    const dealUuid = matchedDeal.deal?.[0]?.uuid;
    if (!dealUuid) {
      throw new Error("Could not find deal UUID for the selected deal.");
    }

    return dealUuid;
  }

  // Main fetch function
  async fetchData() {
    console.log(`Fetch ${this.dataType} data function triggered`);

    try {
      const { selectedDealName, environment, loginResponseDataString } = this.validateRequirements();
      const localData = this.getLocalData();
      const dealUuid = this.getDealUuid(loginResponseDataString, selectedDealName);

      // Prepare API request payload
      const requestBody = {
        environment: environment,
        dealId: dealUuid,
        [`${this.dataType}Data`]: localData,
      };

      console.log(`Sending API request with payload for ${this.dataType}:`, requestBody);

      // Make API call
      const response = await fetch(this.config.apiEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      });

      if (response.ok) {
        const data = await response.json();
        console.log(`Data after fetch for ${this.dataType}:`, data);
        this.displayData(data);
        this.showMessage(`${this.dataType} data fetched successfully!`);
      } else {
        const errorData = await response.text();
        this.showMessage(`Failed to fetch data. Status: ${response.status}`, true);
      }
    } catch (error) {
      console.error(`Error in fetch${this.dataType}Data:`, error);
      this.showMessage(`An unexpected error occurred: ${error.message}`, true);
    }
  }

  // Display data with change highlighting
  displayData(response) {
    const contentArea = document.getElementById("databaseContentArea");
    if (!contentArea) return;

    if (!response.success || !response.data) {
      contentArea.innerHTML = '<div style="color:#dc3545;padding:4px 0;font-size:13px;">Error loading data</div>';
      return;
    }

    // Save to localStorage
    try {
      localStorage.setItem(this.config.storageKey, JSON.stringify(response));
      if (response.data.oldData) {
        localStorage.setItem(this.config.oldDataKey, JSON.stringify(response.data.oldData));
      }
    } catch (e) {
      console.warn("localStorage not available");
    }

    const { changedItems = [], newItems = [], hasChanges } = response.data;
    let html = "";

    if (!hasChanges || changedItems.length == 0) {
      html = '<div style="color:#28a745;padding:4px 0;font-size:13px;">No changes</div>';
    } else {
      // Changed items section
      if (changedItems.length > 0) {
        html += this.renderChangedItems(changedItems);
      }

      // New items section (for closing data)
      // if (newItems && newItems.length > 0) {
      //   html += this.renderNewItems(newItems);
      // }
    }

    contentArea.innerHTML = html;
  }

  // Render changed items
  renderChangedItems(changedItems) {
    const safeDataType = this.dataType.charAt(0).toUpperCase() + this.dataType.slice(1);

    let html = `
    <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
      <div style="font-weight:600;font-size:13px;color:#333;">Modified (${changedItems.length})</div>
      <button onclick="dataFetchers.${this.dataType}.replaceAllItems()" 
              style="background-color:#dc3545;color:white;border:none;padding:4px 8px;border-radius:4px;cursor:pointer;font-size:12px;font-weight:500;transition:all 0.2s;"
              onmouseover="this.style.backgroundColor='#c82333';this.style.transform='translateY(-1px)'"
              onmouseout="this.style.backgroundColor='#dc3545';this.style.transform='none'"
              onmousedown="this.style.transform='translateY(1px)'"
              onmouseup="this.style.transform='translateY(-1px)'">
        Replace All
      </button>
    </div>
  `;

    changedItems.forEach((item) => {
      const itemIdentifier = item[this.config.itemKey] || "No identifier";
      const safeIdentifier = itemIdentifier.replace(/'/g, "\\'");

      const previousContent = item[this.config.previousKey] || "";
      const newContent = item[this.config.newKey] || "";

      const highlighted = this.highlightDifferences(previousContent, newContent);

      html += `
      <div style="margin-bottom:12px;border-bottom:1px solid #eee;padding-bottom:8px;">
        <div style="display:flex;align-items:center;margin-bottom:2px;">
          <div style="font-weight:600;font-size:13px;color:#1a1a1a;">${itemIdentifier}</div>
          <button onclick="dataFetchers.${this.dataType}.replaceSingleItem('${safeIdentifier}')" 
                  style="margin-left:8px;background-color:#198754;color:#FFFFFF;border:none;padding:3px 6px;border-radius:4px;cursor:pointer;font-size:11px;font-weight:500;transition:all 0.2s;"
                  onmouseover="this.style.backgroundColor='#198754';this.style.transform='translateY(-1px)'"
                  onmouseout="this.style.backgroundColor='#115736ff';this.style.transform='none'"
                  onmousedown="this.style.transform='translateY(1px)'"
                  onmouseup="this.style.transform='translateY(-1px)'">
            Replace
          </button>
        </div>
        <div style="display:flex;gap:12px;font-size:13px;margin-top:0px;">
          <div style="flex:1;border-right:1px solid #ddd;padding-right:8px;">
            <div style="font-size:11px;color:#000000;margin-bottom:2px;font-weight:800;">Previous</div>
            <div style="color:#1a1a1a;line-height:1.4;">${highlighted.oldText}</div>
          </div>
          <div style="flex:1;padding-left:8px;">
            <div style="font-size:11px;color:#000000;margin-bottom:2px;font-weight:800;">Updated</div>
            <div style="color:#1a1a1a;line-height:1.4;">${highlighted.newText}</div>
          </div>
        </div>
      </div>
    `;
    });

    return html;
  }

  // Render new items (mainly for closing data)
  renderNewItems(newItems) {
    let html = `
      <div style="margin-top:8px;font-weight:600;font-size:13px;margin-bottom:4px;">
        New Items (${newItems.length})
      </div>
    `;

    newItems.forEach((item) => {
      const content = item[this.config.contentKey] || "No content";
      const heading = item[this.config.itemKey] || "No heading";

      html += `
        <div style="margin-bottom:10px;border-bottom:1px solid #eee;padding-bottom:6px;">
          <div style="font-weight:600;font-size:13px;margin-bottom:2px;color:#1a1a1a;">${heading}</div>
          <div style="font-size:13px;color:#1a1a1a;margin-top:0px;line-height:1.4;">
            ${content}
          </div>
        </div>
      `;
    });

    return html;
  }

  // Highlight differences between two texts
  // Put this method inside your class (replaces the old highlightDifferences)
  highlightDifferences(oldText = "", newText = "") {
    // escape HTML to avoid breaking the layout or XSS
    const escapeHtml = (s) =>
      s
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");

    // Split into word tokens (splits on whitespace). This keeps things simple.
    const oldTokens = oldText.trim() === "" ? [] : oldText.match(/\S+/g) || [];
    const newTokens = newText.trim() === "" ? [] : newText.match(/\S+/g) || [];

    const m = oldTokens.length;
    const n = newTokens.length;

    // Build LCS table (m+1) x (n+1)
    const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        if (oldTokens[i - 1] === newTokens[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;
        else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }

    // Backtrack to find the matching token positions (in reverse)
    const matches = [];
    let i = m,
      j = n;
    while (i > 0 && j > 0) {
      if (oldTokens[i - 1] === newTokens[j - 1]) {
        matches.push([i - 1, j - 1]); // matched token indices
        i--;
        j--;
      } else if (dp[i - 1][j] >= dp[i][j - 1]) {
        i--;
      } else {
        j--;
      }
    }
    matches.reverse(); // now in ascending order

    // Build highlighted outputs by walking through matches
    let oldOut = [];
    let newOut = [];
    let prevA = 0;
    let prevB = 0;

    const pushDeleted = (start, end) => {
      for (let k = start; k < end; k++) {
        oldOut.push(`<span style="text-decoration:line-through;color:#dc3545;">${escapeHtml(oldTokens[k])}</span>`);
      }
    };
    const pushInserted = (start, end) => {
      for (let k = start; k < end; k++) {
        newOut.push(`<span style="text-decoration:underline;color:#28a745;">${escapeHtml(newTokens[k])}</span>`);
      }
    };
    const pushUnchanged = (aIdx, bIdx) => {
      // tokens are equal
      const token = escapeHtml(oldTokens[aIdx]);
      oldOut.push(token);
      newOut.push(token);
    };

    for (const [aIdx, bIdx] of matches) {
      // any deletes in old between prevA..aIdx-1
      if (aIdx > prevA) pushDeleted(prevA, aIdx);
      // any inserts in new between prevB..bIdx-1
      if (bIdx > prevB) pushInserted(prevB, bIdx);
      // matched token
      pushUnchanged(aIdx, bIdx);
      prevA = aIdx + 1;
      prevB = bIdx + 1;
    }

    // tail leftovers
    if (prevA < m) pushDeleted(prevA, m);
    if (prevB < n) pushInserted(prevB, n);

    // Join with a single space between tokens (keeps text readable)
    return {
      oldText: oldOut.join(" "),
      newText: newOut.join(" "),
    };
  }

  // Replace all changed items with previous versions
  replaceAllItems() {
    try {
      const localData = JSON.parse(localStorage.getItem("categoryData") || "{}");
      const dataArray = localData[this.config.dataKey] || [];

      const lastResponse = JSON.parse(localStorage.getItem(this.config.storageKey) || "{}");
      const changedItems = lastResponse.data?.changedItems || [];

      if (!changedItems.length) {
        this.showMessage("No changed items found.", true);
        return;
      }

      let replaced = 0;

      changedItems.forEach((changedItem) => {
        const identifier = changedItem[this.config.itemKey];
        const previousContent = changedItem[this.config.previousKey];

        const index = dataArray.findIndex(
          (item) => item.key === identifier || item[this.config.itemKey] === identifier
        );

        if (index !== -1 && previousContent) {
          // Handle different data structures
          if (dataArray[index].value !== undefined) {
            dataArray[index].value = previousContent;
          } else if (dataArray[index][this.config.contentKey] !== undefined) {
            dataArray[index][this.config.contentKey] = previousContent;
          }
          replaced++;
        }
      });

      localData[this.config.dataKey] = dataArray;
      localStorage.setItem("categoryData", JSON.stringify(localData));

      // Sync updated localData to global categoryData
      if (typeof categoryData !== "undefined") {
        Object.assign(categoryData, localData);
      }

      if (typeof updateCategoryDisplay === "function") {
        updateCategoryDisplay(this.config.dataKey);
      }

      this.showMessage(`Replaced ${replaced} item(s) with previous values.`);
      this.fetchData();
    } catch (error) {
      console.error(`Error in replaceAll${this.dataType}Items:`, error);
      this.showMessage("An error occurred while replacing data.", true);
    }
  }

  // Replace single item with previous version
  replaceSingleItem(identifier) {
    try {
      const localData = JSON.parse(localStorage.getItem("categoryData") || "{}");
      const dataArray = localData[this.config.dataKey] || [];

      const lastResponse = JSON.parse(localStorage.getItem(this.config.storageKey) || "{}");
      const changedItems = lastResponse.data?.changedItems || [];

      const changedItem = changedItems.find((item) => item[this.config.itemKey] === identifier);
      if (!changedItem?.[this.config.previousKey]) {
        this.showMessage(`Previous ${this.config.contentKey} not found for this item.`, true);
        return;
      }

      const index = dataArray.findIndex(
        (item) => item[this.config.itemKey] === identifier || item.key === identifier || item.actionItem === identifier
      );

      if (index === -1) {
        console.error(`Could not find ${identifier} in:`, dataArray);
        this.showMessage(`Item "${identifier}" not found in current data.`, true);
        return;
      }

      // Handle different data structures
      if (dataArray[index].value !== undefined) {
        dataArray[index].value = changedItem[this.config.previousKey];
      } else if (dataArray[index][this.config.contentKey] !== undefined) {
        dataArray[index][this.config.contentKey] = changedItem[this.config.previousKey];
      }

      localData[this.config.dataKey] = dataArray;
      localStorage.setItem("categoryData", JSON.stringify(localData));

      // Update global state if available
      if (typeof categoryData !== "undefined") {
        Object.assign(categoryData, localData);
      }

      if (typeof updateCategoryDisplay === "function") {
        updateCategoryDisplay(this.config.dataKey);
      }

      this.showMessage(`Replaced "${identifier}" with previous version.`);
      this.fetchData();
    } catch (error) {
      console.error("Replacement error:", error);
      this.showMessage("Failed to replace item: " + error.message, true);
    }
  }
}

// Create global instances for different data types
const dataFetchers = {
  representation: new DataFetcher("representation"),
  closing: new DataFetcher("closing"),
};

// Export the main functions for backward compatibility
async function fetchRepresentationData() {
  return dataFetchers.representation.fetchData();
}

async function fetchClosingData() {
  return dataFetchers.closing.fetchData();
}

function displayRepresentationData(response) {
  return dataFetchers.representation.displayData(response);
}

function displayClosingData(response) {
  return dataFetchers.closing.displayData(response);
}

function replaceWithOldData() {
  return dataFetchers.representation.replaceAllItems();
}

function replaceWithOldClosingData() {
  return dataFetchers.closing.replaceAllItems();
}

function replaceSingleItem(identifier) {
  return dataFetchers.representation.replaceSingleItem(identifier);
}

function replaceSingleClosingItem(identifier) {
  return dataFetchers.closing.replaceSingleItem(identifier);
}

// Standalone showMessage function for backward compatibility
function showMessage(msg, isError = false) {
  const contentArea = document.getElementById("databaseContentArea") || document.body;

  const div = document.createElement("div");
  div.textContent = msg;
  div.style.background = isError ? "#ffe6e6" : "#e6ffe6";
  div.style.color = isError ? "red" : "green";
  div.style.padding = "10px";
  div.style.marginBottom = "10px";
  div.style.border = "1px solid #ccc";
  div.style.borderRadius = "5px";
  div.style.fontSize = "14px";
  contentArea.prepend(div);

  setTimeout(() => div.remove(), 4000);
}
