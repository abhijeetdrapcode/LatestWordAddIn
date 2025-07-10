console.log("dealDriver.js loaded");

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
    const selectedCategory = document.getElementById("categorySelect").value;
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
  const response = await fetch("https://dealdriverapi.drapcode.co/addClosingData", {
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
  const formattedcategoryData = window.categoryData[category].reduce((acc, item) => {
    acc[item.key] = item.value;
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

  // Helper function to display messages (for console and potential UI feedback)
  const showMessage = (message, isError = false) => {
    console.log(isError ? `Error: ${message}` : `Success: ${message}`);
    // You can implement a proper message display in the UI here if needed
  };

  try {
    // Ensure the deal selection element exists
    const dealSelect = document.getElementById("dealSelect"); // Assuming dealSelect is globally available or passed
    if (!dealSelect) {
      console.error("Deal select element not available");
      showMessage("Initialization error: Deal selection not found.", true);
      return;
    }

    // Get the selected deal name from the dropdown
    const selectedDealName = dealSelect.options[dealSelect.selectedIndex]?.text;
    console.log("Selected deal name:", selectedDealName);

    if (!selectedDealName) {
      const errorMsg = "Please select a deal first.";
      console.error(errorMsg);
      showMessage(errorMsg, true);
      return;
    }

    // Retrieve environment from local storage
    const environment = localStorage.getItem("selectedEnvironment");
    console.log("Environment from localStorage:", environment);

    if (!environment) {
      const errorMsg = "Environment not found. Please login again.";
      console.error(errorMsg);
      showMessage(errorMsg, true);
      return;
    }

    // Retrieve login response data from local storage to get deal ID
    const loginResponseDataString = localStorage.getItem("loginResponseData");
    console.log("Login response data from localStorage:", loginResponseDataString);

    if (!loginResponseDataString) {
      const errorMsg = "Deal data not found in local storage.";
      console.error(errorMsg);
      showMessage(errorMsg, true);
      return;
    }

    const loginResponseData = JSON.parse(loginResponseDataString);
    const dealsArray = loginResponseData.userDetails?.tenantId || [];
    console.log("Deals array from login response:", dealsArray);

    // Find the matched deal UUID based on the selected deal name
    const matchedDeal = dealsArray.find((deal) => deal.name === selectedDealName);
    console.log("Matched deal:", matchedDeal);

    if (!matchedDeal) {
      const errorMsg = "Could not find matching deal for the selected name.";
      console.error(errorMsg);
      showMessage(errorMsg, true);
      return;
    }

    const dealUuid = matchedDeal.deal?.[0]?.uuid;
    console.log("Deal UUID:", dealUuid);

    if (!dealUuid) {
      const errorMsg = "Could not find deal UUID for the selected deal.";
      console.error(errorMsg);
      showMessage(errorMsg, true);
      return;
    }

    // Attempt to fetch data from the API
    console.log("Attempting to fetch data from API...");
    const response = await fetch("http://localhost:3002/getRepsData", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        environment: environment,
        dealId: dealUuid,
      }),
    });

    console.log("API response status:", response.status);

    if (response.ok) {
      const data = await response.json();
      console.log("API response data:", data);

      // Call the display function with the fetched data
      displayRepresentationData(data);

      showMessage("Data fetched successfully!");
      console.log("Data fetch completed successfully.");
    } else {
      const errorData = await response.text();
      console.error(`Failed to fetch data. Status: ${response.status}`);
      console.error("Error details:", errorData);
      showMessage(`Failed to fetch data. Server responded with status: ${response.status}`, true);
    }
  } catch (error) {
    console.error("Error in fetchRepresentationData:", error);
    showMessage(`An unexpected error occurred: ${error.message}`, true);
  }
}

function displayRepresentationData(data) {
  let displayContent = "";

  if (data.success && data.data && Array.isArray(data.data)) {
    // Filter out items that do not have both article and clause
    const validItems = data.data.filter((item) => item.article && item.clause);

    if (validItems.length > 0) {
      // Format the data to show article (bold part before colon, normal after)
      // and clause (normal text) using HTML
      const formattedItems = validItems.map((item, index) => {
        let formattedArticleHtml = "";
        const colonIndex = item.article.indexOf(":");

        if (colonIndex !== -1) {
          // Split the article into a bold part (before colon) and a normal part (after colon)
          const boldPart = item.article.substring(0, colonIndex);
          const normalPart = item.article.substring(colonIndex); // Includes the colon and the rest
          formattedArticleHtml = `<span style="font-weight: bold;">${boldPart}</span>${normalPart}`;
        } else {
          // If no colon, make the entire article bold as a fallback
          formattedArticleHtml = `<span style="font-weight: bold;">${item.article}</span>`;
        }

        return `
            <div style="margin-bottom: 20px; padding: 10px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9;">
              <div style="color: #333; margin-bottom: 5px; font-size: 1.1em;">
                ${index + 1}. ${formattedArticleHtml}
              </div>
              <div style="color: #666; line-height: 1.4; font-size: 0.95em;">
                ${item.clause}
              </div>
            </div>
          `;
      });

      displayContent = formattedItems.join("");
    } else {
      displayContent =
        "<div style='padding: 10px; color: #888;'>No valid articles and clauses found in the data.</div>";
    }
  } else {
    displayContent = "<div style='padding: 10px; color: #888;'>No valid data found in the API response.</div>";
  }

  // Display the formatted data in the designated content area
  const databaseContentArea = document.getElementById("databaseContentArea");
  if (databaseContentArea) {
    databaseContentArea.innerHTML = displayContent; // Use innerHTML to render HTML content
    console.log("Article and clause data displayed in content area.");
  } else {
    console.error("Database content area element with ID 'databaseContentArea' not found.");
    // Note: You might want to handle this error differently since we're not in the fetch function anymore
  }
}
