console.log("dealDriver.js loaded");

// let selectedCategory = "";
let selectedCategory = localStorage.getItem("selectedCategory") || "repsAndWarranty";

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
  sendDealButton = document.getElementById("sendDealButton");
  fetchDataButton = document.getElementById("fetchDataButton");

  // Add event listeners for post-login functionality
  if (sendDealButton) {
    sendDealButton.addEventListener("click", async () => {
      await handleSendDeal();
    });
  }

  if (fetchDataButton) {
    fetchDataButton.addEventListener("click", async () => {
      await fetchRepresentationData();
    });
  }

  // Show main content
  const mainContent = document.getElementById("mainContent");
  if (mainContent) {
    mainContent.classList.remove("hidden");
  }
}

function handleLogout() {
  isLoggedIn = false;
  loginResponseData = null;
  selectedEnvironment = null;

  // Clear localStorage
  localStorage.removeItem("authToken");
  localStorage.removeItem("loginResponseData");
  localStorage.removeItem("selectedEnvironment");
  localStorage.removeItem("selectedDealId");
  localStorage.removeItem("categoryData");

  // Update UI
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
      localStorage.setItem("authToken", data.token);

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

async function handleSendDeal() {
  if (!dealSelect || !sendDealButton) {
    console.error("Required elements not available");
    return;
  }

  const messageElement = ensureMessageElementExists();
  const showMessage = createMessageHandler(messageElement);

  // Disable the send button during processing
  sendDealButton.disabled = true;
  sendDealButton.style.opacity = "0.5";
  sendDealButton.style.cursor = "not-allowed";

  try {
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
  } catch (error) {
    showMessage("Error sending deal", true);
    console.error("Error sending deal:", error);
  } finally {
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

    const response = await fetch("http://localhost:3002/addClosingData", {
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
    const response = await fetch("http://localhost:3002/parseWord", {
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

async function fetchRepresentationData() {
  console.log("Fetch representation data function triggered");

  // Helper function to display messages
  const showMessage = (message, isError = false) => {
    console.log(isError ? `Error: ${message}` : `Success: ${message}`);
  };

  try {
    // Check for required UI elements and selections
    const dealSelect = document.getElementById("dealSelect");
    if (!dealSelect) {
      showMessage("Deal selection not found.", true);
      return;
    }

    const selectedDealName = dealSelect.options[dealSelect.selectedIndex]?.text;
    if (!selectedDealName) {
      showMessage("Please select a deal first.", true);
      return;
    }

    // Get required data from localStorage
    const environment = localStorage.getItem("selectedEnvironment");
    if (!environment) {
      showMessage("Environment not found. Please login again.", true);
      return;
    }

    const loginResponseDataString = localStorage.getItem("loginResponseData");
    if (!loginResponseDataString) {
      showMessage("Deal data not found in local storage.", true);
      return;
    }

    // Get representation data from localStorage
    const categoryDataString = localStorage.getItem("categoryData");
    let representationData = [];

    if (categoryDataString) {
      try {
        const categoryData = JSON.parse(categoryDataString);
        representationData = categoryData.representation || [];
        console.log("Found representation data in localStorage:", representationData);
      } catch (e) {
        console.error("Error parsing categoryData:", e);
      }
    }

    // Find deal UUID from login data
    const loginResponseData = JSON.parse(loginResponseDataString);
    const dealsArray = loginResponseData.userDetails?.tenantId || [];
    const matchedDeal = dealsArray.find((deal) => deal.name === selectedDealName);

    if (!matchedDeal) {
      showMessage("Could not find matching deal for the selected name.", true);
      return;
    }

    const dealUuid = matchedDeal.deal?.[0]?.uuid;
    if (!dealUuid) {
      showMessage("Could not find deal UUID for the selected deal.", true);
      return;
    }

    // Prepare API request payload
    const requestBody = {
      environment: environment,
      dealId: dealUuid,
      representationData: representationData, // Include the representation data
    };

    console.log("Sending API request with payload:", requestBody);

    // Make API call
    const response = await fetch("http://localhost:3002/getRepsData", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });
    console.log("This is the api response for fetching the collection Data: ", response);
    if (response.ok) {
      const data = await response.json();
      console.log("This is the data after fetch: ", data);
      displayRepresentationData(data);
      showMessage("Data fetched successfully!");
    } else {
      const errorData = await response.text();
      showMessage(`Failed to fetch data. Status: ${response.status}`, true);
    }
  } catch (error) {
    console.error("Error in fetchRepresentationData:", error);
    showMessage(`An unexpected error occurred: ${error.message}`, true);
  }
}

function displayRepresentationData(response) {
  const contentArea = document.getElementById("databaseContentArea");
  if (!contentArea) return;

  if (!response.success || !response.data) {
    contentArea.innerHTML = `
      <div style="color: red; text-align: center; padding: 20px;">
        Error loading data
      </div>
    `;
    return;
  }

  // Save to localStorage
  localStorage.setItem("lastRepresentationApiResponse", JSON.stringify(response));
  if (response.data.oldData) {
    localStorage.setItem("fetchedOldRepresentationData", JSON.stringify(response.data.oldData));
  }

  const { changedItems, newItems, hasChanges } = response.data;
  let html = "";

  if (!hasChanges) {
    html = `
      <div style="text-align: center; padding: 20px; color: #666;">
        No changes detected - all data matches
      </div>
    `;
  } else {
    html += `
      <div style="text-align: center; margin-bottom: 20px;">
        <button 
          onclick="replaceWithOldData()" 
          style="background-color: #dc3545; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px; font-weight: bold;"
        >
          Replace All with Old Data
        </button>
      </div>
    `;

    if (changedItems.length > 0) {
      html += `
        <div style="margin-bottom: 30px;">
          <h3 style="color: #007bff; margin-bottom: 15px; border-bottom: 2px solid #007bff; padding-bottom: 5px;">
            Modified Items (${changedItems.length})
          </h3>
      `;

      changedItems.forEach((item) => {
        html += `
          <div style="background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; padding: 15px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h4 style="color: #495057; margin: 0 0 10px 0; font-size: 16px; font-weight: 600;">
              ${item.article}
            </h4>
            <div style="margin-bottom: 8px;">
              <strong style="color: #dc3545;">Previous:</strong> 
              <span style="background-color: #f8d7da; padding: 2px 6px; border-radius: 3px; display: inline-block; margin-top: 5px;">
                ${item.previousClause}
              </span>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center;">
              <div>
                <strong style="color: #28a745;">Updated:</strong> 
                <span style="background-color: #d4edda; padding: 2px 6px; border-radius: 3px; display: inline-block; margin-top: 5px;">
                  ${item.newClause}
                </span>
              </div>
              <button 
                onclick="replaceSingleItem('${item.article.replace(/'/g, "\\'")}')" 
                style="margin-left: 10px; background-color: #ffc107; border: none; color: black; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px;"
              >
                Replace This
              </button>
            </div>
          </div>
        `;
      });

      html += `</div>`;
    }

    if (newItems.length > 0) {
      html += `
        <div style="margin-bottom: 30px;">
          <h3 style="color: #28a745; margin-bottom: 15px; border-bottom: 2px solid #28a745; padding-bottom: 5px;">
            New Items (${newItems.length})
          </h3>
      `;

      newItems.forEach((item) => {
        html += `
          <div style="background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; padding: 15px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
            <h4 style="color: #495057; margin: 0 0 10px 0; font-size: 16px; font-weight: 600;">
              ${item.article}
            </h4>
            <div style="background-color: #d4edda; padding: 8px; border-radius: 5px; border-left: 4px solid #28a745;">
              ${item.clause}
            </div>
          </div>
        `;
      });

      html += `</div>`;
    }
  }

  contentArea.innerHTML = html;
}

function replaceWithOldData() {
  try {
    const categoryData = JSON.parse(localStorage.getItem("categoryData") || "{}");
    const representation = categoryData.representation || [];

    const lastResponse = JSON.parse(localStorage.getItem("lastRepresentationApiResponse") || "{}");
    const changedItems = lastResponse.data?.changedItems || [];

    if (!changedItems.length) {
      showMessage("No changed items found.", true);
      return;
    }

    let replaced = 0;

    changedItems.forEach((changedItem) => {
      const article = changedItem.article;
      const previousClause = changedItem.previousClause;

      const index = representation.findIndex((item) => item.key === article);
      if (index !== -1 && previousClause) {
        representation[index].value = previousClause;
        replaced++;
      }
    });

    categoryData.representation = representation;
    localStorage.setItem("categoryData", JSON.stringify(categoryData));

    showMessage(`Replaced ${replaced} item(s) with previous clauses.`);
    typeof fetchRepresentationData === "function" ? fetchRepresentationData() : location.reload();
  } catch (error) {
    console.error("Error in replaceWithOldData:", error);
    showMessage("An error occurred while replacing data.", true);
  }
}
function replaceSingleItem(article) {
  try {
    const categoryData = JSON.parse(localStorage.getItem("categoryData") || "{}");
    const representation = categoryData.representation || [];

    const lastResponse = JSON.parse(localStorage.getItem("lastRepresentationApiResponse") || "{}");
    const changedItems = lastResponse.data?.changedItems || [];

    const changedItem = changedItems.find((item) => item.article === article);
    if (!changedItem || !changedItem.previousClause) {
      showMessage("Previous clause not found for this article.", true);
      return;
    }

    const index = representation.findIndex((item) => item.key === article);
    if (index === -1) {
      showMessage("Article not found in current representation.", true);
      return;
    }

    representation[index].value = changedItem.previousClause;
    categoryData.representation = representation;
    localStorage.setItem("categoryData", JSON.stringify(categoryData));

    showMessage(`Replaced "${article}" with previous clause.`);
    typeof fetchRepresentationData === "function" ? fetchRepresentationData() : location.reload();
  } catch (error) {
    console.error("Error in replaceSingleItem:", error);
    showMessage("An error occurred while replacing the item.", true);
  }
}

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
