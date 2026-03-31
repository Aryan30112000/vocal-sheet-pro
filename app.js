(function () {
  const config = window.APP_CONFIG;
  const requiredScopes = "openid email profile https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.readonly";

  const state = {
    accessToken: "",
    tokenClient: null,
    currentSheetId: localStorage.getItem("google_sheet_id") || "",
    currentSheetUrl: localStorage.getItem("google_sheet_url") || "",
    recognition: null,
    audioContext: null,
    analyser: null
  };

  const elements = {
    loginButton: document.getElementById("login-button"),
    logoutButton: document.getElementById("logout-button"),
    micButton: document.getElementById("mic-button"),
    saveTextButton: document.getElementById("save-text-button"),
    taskInput: document.getElementById("task-input"),
    langSelect: document.getElementById("language-select"),
    toastContainer: document.getElementById("toast-container"),
    sheetLink: document.getElementById("sheet-link"),
    authStatus: document.getElementById("auth-status"),
    canvas: document.getElementById("visualizer"),
    mainApp: document.getElementById("main-app"),
    authSection: document.getElementById("auth-section"),
    // Dashboard elements
    taskList: document.getElementById("task-list"),
    dashboardSection: document.getElementById("dashboard-section")
  };

  function showToast(msg) {
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.textContent = msg;
    elements.toastContainer.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
  }

  // --- DASHBOARD: Recent Tasks Fetch logic ---
  async function fetchRecentTasks() {
    if (!state.currentSheetId || !state.accessToken) return;
    
    try {
      // Sheet1!A2:C11 se data la rahe hain (Headers chhod kar)
      const data = await fetchJson(`https://sheets.googleapis.com/v4/spreadsheets/${state.currentSheetId}/values/Sheet1!A2:C11`, {
        headers: { Authorization: "Bearer " + state.accessToken }
      });

      elements.taskList.innerHTML = ""; 

      if (data.values && data.values.length > 0) {
        elements.dashboardSection.style.display = "block";
        // Reverse loop taaki naya sabse upar dikhe
        data.values.reverse().slice(0, 5).forEach(row => {
          const card = document.createElement("div");
          card.style.cssText = "background: #1c1c1e; padding: 15px; border-radius: 12px; text-align: left; border-left: 3px solid #0a84ff; margin-bottom: 10px;";
          card.innerHTML = `
            <div style="font-size: 10px; color: #888; text-transform: uppercase; font-weight: bold;">${row[0] || 'GENERAL'}</div>
            <div style="font-size: 15px; margin: 5px 0; color: white;">${row[1] || 'No Task'}</div>
            <div style="font-size: 10px; color: #555;">${row[2] || ''}</div>
          `;
          elements.taskList.appendChild(card);
        });
      }
    } catch (e) {
      console.error("Dashboard error:", e);
    }
  }

  function startVisualizer() {
    if (!state.audioContext) {
      state.audioContext = new (window.AudioContext || window.webkitAudioContext)();
      state.analyser = state.audioContext.createAnalyser();
      navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        const source = state.audioContext.createMediaStreamSource(stream);
        source.connect(state.analyser);
        drawWave();
      }).catch(() => showToast("Mic access denied"));
    }
  }

  function drawWave() {
    const ctx = elements.canvas.getContext("2d");
    const bufferLength = state.analyser.frequencyBinCount;
    const dataArray = new Uint8Array(bufferLength);
    function animate() {
      requestAnimationFrame(animate);
      state.analyser.getByteTimeDomainData(dataArray);
      ctx.clearRect(0, 0, elements.canvas.width, elements.canvas.height);
      ctx.lineWidth = 2; ctx.strokeStyle = "#0a84ff"; ctx.beginPath();
      let sliceWidth = elements.canvas.width / bufferLength;
      let x = 0;
      for (let i = 0; i < bufferLength; i++) {
        let v = dataArray[i] / 128.0;
        let y = v * elements.canvas.height / 2;
        if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        x += sliceWidth;
      }
      ctx.lineTo(elements.canvas.width, elements.canvas.height / 2);
      ctx.stroke();
    }
    animate();
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error?.message || "Error");
    return data;
  }

  async function findExistingSheet() {
    const query = encodeURIComponent("name = 'Vocal Sheet Pro' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false");
    try {
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
    } catch(e) {}
    return false;
  }

  async function createSheet() {
    if (await findExistingSheet()) return updateSheetUi();
    showToast("Creating Stylized Sheet...");
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

    // Formatting Header
    await fetchJson(`https://sheets.googleapis.com/v4/spreadsheets/${state.currentSheetId}:batchUpdate`, {
      method: "POST",
      headers: { Authorization: "Bearer " + state.accessToken, "Content-Type": "application/json" },
      body: JSON.stringify({
        requests: [{
          repeatCell: {
            range: { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 3 },
            cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 12 }, backgroundColor: { red: 0.9, green: 0.9, blue: 0.9 } } },
            fields: "userEnteredFormat(textFormat,backgroundColor)"
          }
        }]
      })
    });
    updateSheetUi();
  }

  function updateSheetUi() {
    if (state.currentSheetUrl) { 
      elements.sheetLink.href = state.currentSheetUrl; 
      elements.sheetLink.style.display = "block"; 
      fetchRecentTasks(); // UI update hote hi dashboard load karein
    }
  }

  async function appendTask() {
    const rawText = elements.taskInput.value.trim();
    if (!rawText) return showToast("Type something!");
    showToast("Saving...");
    try {
      if (!state.currentSheetId && !(await findExistingSheet())) await createSheet();
      let category = "General", task = rawText;
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
      showToast(`Saved to ${category}`);
      elements.taskInput.value = "";
      fetchRecentTasks(); // Task save hote hi dashboard refresh
    } catch (e) {
      if (e.message.toLowerCase().includes("not found")) { state.currentSheetId = ""; await appendTask(); }
      else showToast(e.message);
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
          if (resp.error) {
             showToast("Auth Error: " + resp.error);
             return;
          }
          state.accessToken = resp.access_token;
          // Session ko temporary save karein (optional but helpful)
          sessionStorage.setItem("temp_access_token", resp.access_token);
          
          elements.authSection.style.display = "none";
          elements.mainApp.style.display = "block";
          elements.logoutButton.style.display = "inline-block";
          elements.authStatus.textContent = "Connected";
          findExistingSheet().then(() => updateSheetUi());
          showToast("Welcome Back!");
        }
      });

      // --- AUTO-LOGIN LOGIC ON REFRESH ---
      // Agar pehle se logged in tha (Sheet ID hai), toh token request bhejo bina popup dikhaye
      if (state.currentSheetId) {
          console.log("Re-authenticating session...");
          // 'prompt: none' user ko disturb nahi karega agar session valid hai
          state.tokenClient.requestAccessToken({ prompt: 'none' });
      }
    };
    document.head.appendChild(script);
  }

  function initSpeech() {
    const Speech = window.SpeechRecognition || window.webkitSpeechRecognition;
    state.recognition = new Speech();
    state.recognition.onstart = () => { startVisualizer(); document.getElementById("mic-status").textContent = "Listening..."; };
    state.recognition.onresult = (e) => { elements.taskInput.value = e.results[0][0].transcript; };
    state.recognition.onend = () => { document.getElementById("mic-status").textContent = "Tap to speak"; };
  }

  elements.loginButton.onclick = () => state.tokenClient.requestAccessToken({ prompt: "" });
  elements.logoutButton.onclick = () => { localStorage.clear(); location.reload(); };
  elements.micButton.onclick = () => { state.recognition.lang = elements.langSelect.value; state.recognition.start(); };
  elements.saveTextButton.onclick = appendTask;

  initAuth(); initSpeech();
})();
