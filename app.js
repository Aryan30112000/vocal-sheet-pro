(function () {
  const config = window.APP_CONFIG;
  const requiredScopes = "openid email profile https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets";

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

  async function createSheet() {
    setMessage("Creating Smart Sheet...");
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

    elements.sheetLink.href = state.currentSheetUrl;
    elements.sheetLink.style.display = "inline-block";
    setMessage("Smart Sheet Ready!");
  }

  async function appendTask() {
    const rawText = elements.taskInput.value.trim();
    if (!rawText) return setMessage("Speak or type something first!", true);
    setMessage("Saving...");

    try {
      if (!state.currentSheetId) await createSheet();

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
          if(state.currentSheetUrl) { 
              elements.sheetLink.href = state.currentSheetUrl; 
              elements.sheetLink.style.display = "inline-block"; 
          }
          setMessage("Logged In! Speak or Type.");
        }
      });
    };
    document.head.appendChild(script);
  }

  function initSpeech() {
    const Speech = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!Speech) return setMessage("Speech not supported", true);
    state.recognition = new Speech();
    state.recognition.continuous = false;
    state.recognition.interimResults = true;

    state.recognition.onstart = () => { setMessage("Listening..."); };
    state.recognition.onresult = (e) => { elements.taskInput.value = e.results[0][0].transcript; };
    state.recognition.onend = () => { setMessage("Stopped. Edit then Save."); };
  }

  elements.loginButton.onclick = () => state.tokenClient.requestAccessToken({ prompt: "" });
  elements.logoutButton.onclick = () => { localStorage.clear(); location.reload(); };
  elements.micButton.onclick = () => { 
    state.recognition.lang = elements.langSelect.value;
    state.recognition.start(); 
  };
  elements.saveTextButton.onclick = appendTask;

  initAuth();
  initSpeech();
})();