(function () {
  const config = window.APP_CONFIG;
  const requiredScopes = "openid email profile https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.readonly";

  const state = {
    accessToken: "",
    tokenClient: null,
    currentSheetId: localStorage.getItem("google_sheet_id") || "",
    currentSheetUrl: localStorage.getItem("google_sheet_url") || "",
    recognition: null
  };

  const elements = {
    loginButton: document.getElementById("login-button"),
    logoutButton: document.getElementById("logout-button"),
    micButton: document.getElementById("mic-button"),
    saveTextButton: document.getElementById("save-text-button"),
    taskInput: document.getElementById("task-input"),
    langSelect: document.getElementById("language-select"),
    messageBanner: document.getElementById("message-banner"),
    sheetLink: document.getElementById("sheet-link"),
    authStatus: document.getElementById("auth-status")
  };

  function setMessage(msg, isError) {
    elements.messageBanner.textContent = msg;
    elements.messageBanner.style.background = isError ? "#ffdada" : "#e0ffd0";
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error?.message || "Error");
    return data;
  }

  // --- SMART LOGIC: Pehle Drive mein dhoondo ---
  async function findExistingSheet() {
    setMessage("Searching for your existing sheet...");
    const query = encodeURIComponent("name = 'Vocal Sheet Pro' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false");
    const data = await fetchJson(`https://www.googleapis.com/drive/v3/files?q=${query}&fields=files(id,name)`, {
      headers: { Authorization: "Bearer " + state.accessToken }
    });

    if (data.files && data.files.length > 0) {
      state.currentSheetId = data.files[0].id;
      state.currentSheetUrl = `https://docs.google.com/spreadsheets/d/${state.currentSheetId}/edit`;
      localStorage.setItem("google_sheet_id", state.currentSheetId);
      localStorage.setItem("google_sheet_url", state.currentSheetUrl);
      return true;
    }
    return false;
  }

  async function createSheet() {
    // Pehle search karo, agar mil gayi toh nayi mat banao
    const found = await findExistingSheet();
    if (found) {
      updateSheetUi();
      setMessage("Found your existing sheet!");
      return;
    }

    setMessage("No sheet found. Creating new one...");
    const sheet = await fetchJson("https://sheets.googleapis.com/v4/spreadsheets", {
      method: "POST",
      headers: { Authorization: "Bearer " + state.accessToken, "Content-Type": "application/json" },
      body: JSON.stringify({ properties: { title: "Vocal Sheet Pro" } })
    });

    state.currentSheetId = sheet.spreadsheetId;
    state.currentSheetUrl = sheet.spreadsheetUrl;
    localStorage.setItem("google_sheet_id", state.currentSheetId);
    localStorage.setItem("google_sheet_url", state.currentSheetUrl);

    await fetchJson(`https://sheets.googleapis.com/v4/spreadsheets/${state.currentSheetId}/values/Sheet1!A1:C1?valueInputOption=USER_ENTERED`, {
      method: "PUT",
      headers: { Authorization: "Bearer " + state.accessToken, "Content-Type": "application/json" },
      body: JSON.stringify({ values: [["CATEGORY", "TASK", "TIMESTAMP"]] })
    });

    updateSheetUi();
    setMessage("New Smart Sheet Ready!");
  }

  function updateSheetUi() {
    if (state.currentSheetUrl) {
      elements.sheetLink.href = state.currentSheetUrl;
      elements.sheetLink.style.display = "inline-block";
    }
  }

  async function appendTask() {
    const rawText = elements.taskInput.value.trim();
    if (!rawText) return setMessage("Speak or type something!", true);
    setMessage("Saving...");

    try {
      if (!state.currentSheetId) {
          const found = await findExistingSheet();
          if (!found) await createSheet();
      }

      let category = "General";
      let task = rawText;
      if (rawText.includes(":")) {
        const parts = rawText.split(":");
        category = parts[0].trim().toUpperCase();
        task = parts[1].trim();
      }

      await fetchJson(`https://sheets.googleapis.com/v4/spreadsheets/${state.currentSheetId}/values/Sheet1!A:C:append?valueInputOption=USER_ENTERED`, {
        method: "POST",
        headers: { Authorization: "Bearer " + state.accessToken, "Content-Type": "application/json" },
        body: JSON.stringify({ values: [[category, task, new Date().toLocaleString()]] })
      });

      setMessage("Saved in " + category);
      elements.taskInput.value = "";
    } catch (e) {
      if (e.message.includes("404")) { state.currentSheetId = ""; await appendTask(); }
      else setMessage("Error: " + e.message, true);
    }
  }

  function initAuth() {
    const script = document.createElement("script");
    script.src = "https://accounts.google.com/gsi/client";
    script.onload = () => {
      state.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: config.googleClientId,
        scope: requiredScopes,
        callback: (resp) => {
          state.accessToken = resp.access_token;
          elements.authStatus.textContent = "Signed In";
          elements.loginButton.style.display = "none";
          elements.logoutButton.style.display = "inline";
          // Login hote hi pehle dhoondo
          findExistingSheet().then(() => updateSheetUi());
          setMessage("Logged In! Checking for your sheet...");
        }
      });
    };
    document.head.appendChild(script);
  }

  function initSpeech() {
    const Speech = window.SpeechRecognition || window.webkitSpeechRecognition;
    state.recognition = new Speech();
    state.recognition.onstart = () => setMessage("Listening...");
    state.recognition.onresult = (e) => { elements.taskInput.value = e.results[0][0].transcript; };
    state.recognition.onend = () => setMessage("Stopped.");
  }

  elements.loginButton.onclick = () => state.tokenClient.requestAccessToken({ prompt: "" });
  
  // LOGOUT FIX: localStorage.clear() hata diya
  elements.logoutButton.onclick = () => { 
      state.accessToken = ""; 
      elements.authStatus.textContent = "Signed Out";
      elements.loginButton.style.display = "inline";
      elements.logoutButton.style.display = "none";
      setMessage("Signed Out. Session Saved.");
  };

  elements.micButton.onclick = () => { 
    state.recognition.lang = elements.langSelect.value;
    state.recognition.start(); 
  };
  elements.saveTextButton.onclick = appendTask;

  initAuth();
  initSpeech();
})();
