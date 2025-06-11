/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
// import { apiURL } from "./config.js";
const apiURL = "http://localhost:3001";

var threadId, templates = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    getData();
    run();
  }
});

export async function run() {
  setTimeout(() => {
    if (Office.context.mailbox.item) {
      if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
        if (Office.context.mailbox.item.body.getTypeAsync) {
          // This method exists only in Compose mode
          console.log("Compose mode");
        } else {
          console.log("Read mode");
        }
      }
    }

    const item = Office.context.mailbox.item;
    const itemId = item.itemId;
    console.log(item.conversationId);

  }, 3000); // Optional: adjust timeout if needed

  const customSelect = document.querySelector(".custom-select");
  const selectedText = document.querySelector(".selected-text");
  const searchInput = document.querySelector(".search-input");
  const options = document.querySelectorAll(".option");

  let selectedTone = "";

  // Custom select functionality
  customSelect.addEventListener("click", (e) => {
    if (e.target.closest(".search-input")) return;
    customSelect.classList.toggle("open");
    if (customSelect.classList.contains("open")) {
      searchInput.focus();
    }
  });

  // Close select when clicking outside
  document.addEventListener("click", (e) => {
    if (!customSelect.contains(e.target)) {
      customSelect.classList.remove("open");
    }
  });

  // Search functionality
  searchInput.addEventListener("input", (e) => {
    const searchTerm = e.target.value.toLowerCase();
    options.forEach((option) => {
      const text = option.textContent.toLowerCase();
      option.classList.toggle("hidden", !text.includes(searchTerm));
    });
  });

  // Option selection
  options.forEach((option) => {
    option.addEventListener("click", (e) => {
      e.stopPropagation(); // Stop event from bubbling up
      selectedTone = option.dataset.value;
      selectedText.textContent = option.textContent;
      options.forEach((opt) => opt.classList.remove("selected"));
      option.classList.add("selected");
      customSelect.classList.remove("open");
      searchInput.value = ""; // Clear search input
      // Reset visibility of all options
      options.forEach((opt) => opt.classList.remove("hidden"));
    });
  });

  document.querySelector(".generate-btn").addEventListener("click", async (e) => {
    e.preventDefault();
    if (document.querySelector(".selected-text").textContent == "Select a tone") {
      showNotification("Select tone to continue.");
    } else if (document.querySelector(".input-text").value == "") {
      showNotification("Fill prompt to continue.");
    } else {
      freezeGenerateBtn();
      const myHeaders = new Headers();
      myHeaders.append("Content-Type", "application/json");
      myHeaders.append("Authorization", `Bearer ${localStorage.getItem("access_token")}`);

      const raw = JSON.stringify({
        model: "gpt-4o",
        temperature: 0.7,
        messages: [
          {
            role: "system",
            content: `You are an AI assistant that writes highly professional and structured emails with a ${document.querySelector(".selected-text").textContent} tone.`,
          },
          {
            role: "user",
            content: document.querySelector(".input-text").value,
          },
          {
            role: "user",
            content: "Exclude Subject from email only email body is required",
          },
        ],
      });

      const requestOptions = {
        method: "POST",
        headers: myHeaders,
        body: raw,
        redirect: "follow",
      };

      fetch(`${apiURL}/openai`, requestOptions)
        .then(async (response) => {
          if (!response.ok) {
            // Server responded with an error status
            const errorText = await response.text();
            throw new Error(`HTTP error! Status: ${response.status}, Message: ${errorText}`);
          }
          try {
            const result = await response.json();
            // Handle the result here
            console.log('Success:', result);
            insertSimpleText(result.content);
          } catch (jsonError) {
            throw new Error('Failed to parse JSON: ' + jsonError.message);
          }
        })
        .catch((error) => {
          console.log(error);
          if (error.message.includes("Invalid token")) {
            showNotification("Session expired.");
            setTimeout(() => {
              localStorage.removeItem("access_token");
              localStorage.removeItem("userLoggedIn");
              window.location.href = "login.html";
            }, 1500);
          } else {
            showNotification("Something went wrong.");
          }
        }).finally(() => {
          releaseGenerateBtn();
        });
    }
  });
}

function freezeGenerateBtn(){
  const generateBtn = document.querySelector(".generate-btn");
  generateBtn.disabled = true;
  generateBtn.innerHTML =
    'Generating... <span class="emoji">⚡</span>';
}

function releaseGenerateBtn(){
  const generateBtn = document.querySelector(".generate-btn");
  generateBtn.disabled = false;
  generateBtn.innerHTML =
    'Generate Response <span class="emoji">✨</span>';
}

function showContentLoader() {
  const wrapper = document.querySelector('.hubspot-overlay');
  const loader = document.createElement('div');
  loader.className = 'content-loader';
  loader.innerHTML = `
      <div class="loader-spinner"></div>
      <div class="loader-text">Loading content...</div>
  `;
  wrapper.appendChild(loader);
}

function hideContentLoader() {
  const loader = document.querySelector('.content-loader');
  if (loader) {
      loader.remove();
  }
}

async function getData() {
  const item = Office.context.mailbox.item;
  threadId = item.conversationId;

  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");
  myHeaders.append("Authorization", `Bearer ${localStorage.getItem("access_token")}`);

  const raw = JSON.stringify({
    "thread_id": decodeURIComponent(item.conversationId)
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow"
  };

  fetch(`${apiURL}/get-data`, requestOptions)
    .then(response => {
      if (!response.ok) {
        return response.text().then(text => {
          throw new Error(`HTTP error ${response.status}: ${text}`);
        });
      }
      return response.json();
    })
    .then(async (result) => {
      // Send a response back to content.js
      console.log(result);
      templates = result.templates;
      menuTemplate(result);
    })
    .catch((error) => {
      console.log(error);
      if (error.message.includes("Invalid token")) {
        showNotification("Session expired.");
        setTimeout(() => {
          localStorage.removeItem("access_token");
          localStorage.removeItem("userLoggedIn");
          window.location.href = "login.html";
        }, 1500);
      } else {
        showNotification("Something went wrong.");
      }
    });
}

function menuTemplate(data) {
  let template = document.createElement("DIV");
  template.style.position = "sticky";
  template.innerHTML = `<div><nav class="top-nav">
        <ul class="nav-list">
            <li class="nav-item">
                ${data.emails != null && data.emails.length > 0 ? `
                <button class="nav-button menu-sources">
                    <svg class="nav-icon" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V9z"/>
                        <polyline points="9 22 9 12 15 12 15 22"/>
                    </svg>
                    <span>Sources</span>
                </button>` :
      `<button class="nav-button menu-compose">
                    <svg class="nav-icon" viewBox="0 0 24 24" fill="currentColor">
                        <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V9z"/>
                        <polyline points="9 22 9 12 15 12 15 22"/>
                    </svg>
                    <span>Compose</span>
                </button>`
    }
            </li>
            <li class="nav-item">
                <button class="nav-button menu-iterate">
                    <svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                        <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
                        <line x1="16" y1="2" x2="16" y2="6"/>
                        <line x1="8" y1="2" x2="8" y2="6"/>
                        <line x1="3" y1="10" x2="21" y2="10"/>
                    </svg>
                    <span>Iterate</span>
                </button>
            </li>
            <li class="nav-item">
                <button class="nav-button menu-templates">
                    <svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                        <polyline points="14 2 14 8 20 8"></polyline>
                        <line x1="16" y1="13" x2="8" y2="13"></line>
                        <line x1="16" y1="17" x2="8" y2="17"></line>
                        <polyline points="10 9 9 9 8 9"></polyline>
                    </svg>
                    <span>Templates</span>
                </button>
                <div class="documentation-popup">
                    <div class="search-container">
                        <input type="text" class="doc-search" placeholder="Search template...">
                    </div>
                    <div class="doc-list">
                        ${data && data.templates.map(template => `
                            <div class="doc-item" data-template-id="${template.id}">${template.name}</div>
                        `).join('')}
                    </div>
                    <div class="doc-footer">
                        <a href="#" class="doc-link">Alles weergeven</a>
                        <a href="#" class="doc-link">Nieuwe aanmaken</a>
                    </div>
                </div>
            </li>
            <li class="nav-item">
                ${data.emails != null && data.emails.length > 0 ? `
                <div class="nav-button" id="menuCustomA">
                    <svg class="nav-icon" viewBox="0 0 24 24" fill="currentColor">
                        <circle cx="12" cy="12" r="10"/>
                        <circle cx="12" cy="12" r="6"/>
                        <circle cx="12" cy="12" r="2"/>
                    </svg>
                    <span>${data.emails[0].hoa}</span>
                </div>` : ``}
            </li>
            <li class="nav-item">
                ${data.emails != null && data.emails.length > 0 ? `
                <div class="nav-button" id="menuCustomB">
                    <svg class="nav-icon" viewBox="0 0 24 24" fill="currentColor">
                        <circle cx="12" cy="12" r="10"/>
                        <circle cx="12" cy="12" r="6"/>
                        <circle cx="12" cy="12" r="2"/>
                    </svg>
                    <span>${data.emails[0].sensitivity}</span>
                </div>` : ``}
            </li>
        </ul>

        <div class="right-section">
          <label class="nav-button">
              <input type="checkbox" class="checkbox-input menu=logbook">
              <span class="checkbox-text">Logboek</span>
          </label>
          <label class="nav-button">
              <input type="checkbox" class="checkbox-input menu-follow">
              <span class="checkbox-text">Volgen</span>
          </label>
        </div>
    </nav></div>`;
  document.querySelector(".nav-section").appendChild(template);
  hideContentLoader();

  // Search functionality
  document.querySelectorAll('.doc-search').forEach(searchInput => {
    searchInput.addEventListener('input', function (e) {
      const searchText = e.target.value.toLowerCase();
      const items = document.querySelectorAll('.doc-item'); // All items to filter

      items.forEach(item => {
        const text = item.textContent.toLowerCase();
        item.style.display = text.includes(searchText) ? 'block' : 'none';
      });
    });
  });

  // Close popup when clicking outside
  document.addEventListener('click', function (e) {
    try {
      const popups = document.querySelectorAll('.documentation-popup');
      const buttons = document.querySelectorAll('.menu-templates');

      popups.forEach(popup => {
        buttons.forEach(btn => {
          // Close the popup if clicked outside of the popup or the button
          if (!popup.contains(e.target) && !btn.contains(e.target)) {
            popup.classList.remove('show');
          }
        });
      });
    } catch (e) {
      console.log(e);
    }
  });

  // Toggle popup on hover for all matching elements
  document.querySelectorAll('.menu-templates').forEach((docBtn, index) => {
    const popup = document.querySelectorAll('.documentation-popup')[index]; // Assuming the order matches
    docBtn.addEventListener('mouseenter', () => {
      popup.classList.add('show');
    });

    popup.addEventListener('mouseleave', () => {
      popup.classList.remove('show');
    });
  });

  document.querySelector(".menu-sources")?.addEventListener("click", function (e) {
    if (threadId) {
      window.open(`https://agent.triple.blue/chat?threadId=${threadId}`, "_blank");
    } else {
      resetStateAndLaunch();
    }
  });

  document.querySelector(".menu-compose")?.addEventListener("click", function (e) {
    document.querySelector('button[data-ext="genAI"]').click();
  });

  document.querySelector(".menu-iterate")?.addEventListener("click", function (e) {
    if (threadId) {
      window.open(`https://agent.triple.blue/chat?threadId=${threadId}`, "_blank");
    } else {
      resetStateAndLaunch();
    }
  });

  // Select all document items
  const docItems = document.querySelectorAll('.doc-item');

  // Add click event listener to each doc-item
  docItems.forEach(item => {
    item.addEventListener('click', function () {
      const templateId = parseInt(item.getAttribute('data-template-id'), 10);
      getTemplateById(templateId, async function (template) {
        if (template) {
          await insertSimpleText(template.body.replace(/(\r\n|\n|\r)/g, '<br>'));
        } else {
          console.log('Template not found');
        }
      });
    });
  });
}

function getTemplateById(templateId, callback) {
  const templatesMail = templates || [];
  const template = templatesMail.find(t => t.id === templateId);
  callback(template); // Return the template to the callback function
}

async function insertSimpleText(response) {
  const htmlFormattedText = response.replace(/\n/g, '<br>');

  Office.context.mailbox.item.body.setSelectedDataAsync(
    htmlFormattedText,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Text inserted successfully.");
      } else {
        console.error("Error inserting text:", result.error.message);
      }
    }
  );
}

function showNotification(message, type = "error") {
  // Remove existing notification if any
  const existingNotification = document.querySelector(".notification");
  if (existingNotification) {
    existingNotification.remove();
  }

  // Create notification elements
  const notification = document.createElement("div");
  notification.className = `notification ${type}`;

  const icon = document.createElement("span");
  icon.className = "notification-icon";
  icon.textContent = type === "error" ? "" : "";

  const text = document.createElement("span");
  text.className = "notification-message";
  text.textContent = message;

  const closeBtn = document.createElement("button");
  closeBtn.className = "notification-close";
  closeBtn.textContent = "×";

  notification.appendChild(icon);
  notification.appendChild(text);
  notification.appendChild(closeBtn);

  // Add to document
  document.body.appendChild(notification);

  // Show notification with animation
  requestAnimationFrame(() => {
    notification.classList.add("show");
  });

  // Auto-hide after 5 seconds
  const hideTimeout = setTimeout(() => {
    hideNotification(notification);
  }, 5000);

  // Close button handler
  closeBtn.addEventListener("click", () => {
    clearTimeout(hideTimeout);
    hideNotification(notification);
  });
}

function hideNotification(notification) {
  notification.classList.remove("show");
  setTimeout(() => {
    notification.remove();
  }, 300); // Match transition duration
}

function resetState() {
  localStorage.removeItem("access_token");
  localStorage.removeItem("userLoggedIn");
}

function resetStateAndLaunch() {
  localStorage.removeItem("access_token");
  localStorage.removeItem("userLoggedIn");
  window.location.href = "login.html";
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function insertHtmlInBody() {
  var item = Office.context.mailbox.item;

  if (item.itemType === Office.MailboxEnums.ItemType.Message) {
    var htmlToInsert = "<p><strong>Hello, this is the <em>HTML</em> content inserted!</strong></p>";

    item.body.setSelectedDataAsync(htmlToInsert, {
      coercionType: Office.CoercionType.Html
    }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("HTML inserted successfully!");
      } else {
        console.error("Error inserting HTML: ", asyncResult.error.message);
      }
    });
  } else {
    console.error("The item is not a message.");
  }
}
