/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import { apiURL } from "./config.js";

showContentLoader();
// localStorage.removeItem("userLoggedIn");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    
    setTimeout(()=>{
        const isLoggedIn = localStorage.getItem("userLoggedIn");
        if (isLoggedIn == "true") {
            window.location.href = "home.html";
        }else{
            hideContentLoader();
        }
        run();
    },500);
  }
});

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

async function login(e, email, password) {
  const button = e.target;
  const btnText = button.querySelector('.btn-text');
  const spinner = button.querySelector('.spinner');
  
  try {
    // Show loading state
    btnText.style.opacity = '0';
    spinner.style.display = 'block';
  } catch (error) { 
    console.log(error);
  }
  
  const myHeaders = new Headers();
  myHeaders.append("Content-Type", "application/json");

  const raw = JSON.stringify({
    "email": email,
    "password": password,
  });

  const requestOptions = {
    method: "POST",
    headers: myHeaders,
    body: raw,
    redirect: "follow",
  };

  var jwt;
  await fetch(`${apiURL}/login`, requestOptions)
    .then(response => {
      if (!response.ok) {
        return response.text().then(text => {
          throw new Error(`HTTP error ${response.status}: ${text}`);
        });
      }
      return response.json();
    })
    .then(async (result) => {
      try {
        button.disabled = false;
        btnText.style.opacity = '1';
        spinner.style.display = 'none';
      } catch (error) { 
        console.log(error);
      }
      console.log(result);
      // JWT Token can be accessed here:
      jwt = result.access_token;
    })
    .catch((error) => {
      console.log(error);
      showNotification("Something went wrong.");
      try {
        button.disabled = false;
        btnText.style.opacity = '1';
        spinner.style.display = 'none';
      } catch (error) { 
        console.log(error);
      }
    });
  // Store in localStorage or chrome.storage for further use
  return jwt;
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

export async function run() {

  document.querySelector('.login-btn').addEventListener('click', () => showScreen('login-screen'));
  document.querySelector('.signup-btn').addEventListener('click', () => window.open("https://agent.triple.blue/", '_blank'));

  function showScreen(screenId) {
    // Hide all screens
    document.querySelector("#loginModal").querySelectorAll('.extension-card').forEach(screen => {
        screen.style.display = 'none';
    });
    // Show requested screen
    document.getElementById(screenId).style.display = 'flex';
  }

  document.querySelector("#tpSubmitButton").addEventListener("click", async (e) => {
    e.preventDefault();

    const email = document.getElementById("login-email").value;
    const password = document.getElementById("login-password").value;

    if (email && password) {
      const access_token = await login(e, email, password);
      console.log(access_token);
      if (access_token) {
        // Store the access token in localStorage or chrome.storage
        localStorage.setItem("access_token", access_token);
        localStorage.setItem("userLoggedIn", "true");
        // Redirect to home.html
        window.location.href = "home.html";
      } else {
        showNotification("Login failed. Please check your credentials.");
      }
    } else {
      showNotification("Please enter both email and password.");
      try {
        button.disabled = false;
        btnText.style.opacity = '1';
        spinner.style.display = 'none';
      } catch (error) { 
        console.log(error);
      }
    }
  });
}