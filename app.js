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
    authSection: document.getElementById("auth-section")
  };

  function showToast(msg, type = "success") {
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.style.borderLeft = `4px solid ${type === "success" ? "#30d158" : "#ff453a"}`;
    toast.textContent = msg;
    elements.toastContainer.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
  }

  function startVisualizer() {
    if (!state.audioContext) {
      state.audioContext = new (window.AudioContext || window.webkitAudioContext)();
      state.analyser = state.audioContext.createAnalyser();
      navigator.mediaDevices.getUserMedia({ audio: true }).then(stream => {
        const source = state.audioContext.createMediaStreamSource(stream);
        source.connect(state.analyser);
        drawWave();
      }).catch(() => showToast("Mic access denied", "error"));
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
      ctx.lineWidth = 2;
      ctx.strokeStyle = "#0a84ff";
      ctx.beginPath();
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
    const found = await findExistingSheet();
    if (found) return updateSheetUi();

    showToast("Creating Stylized Pro Sheet...");
    const sheet = await fetchJson("https://sheets.googleapis.com/v4/spreadsheets", {
      method: "POST",
      headers: { Authorization: "Bearer " + state.accessToken, "Content-Type": "application/json" },
      body: JSON.stringify({ properties: { title: "Vocal Sheet Pro" } })
    });

    state.currentSheetId = sheet.spreadsheetId;
    state.currentSheetUrl = sheet.spreadsheetUrl;
    localStorage.setItem("google_sheet_id", state.currentSheetId);
    localStorage.setItem("google_sheet_url", state.currentSheetUrl);

    // 1. Pehle Headers likho
    await fetchJson(`https://sheets.googleapis.com/v4/spreadsheets/${state.currentSheetId}/values/Sheet1!A1:C1?valueInputOption=USER_ENTERED`, {
      method: "PUT",
      headers: { Authorization: "Bearer " + state.accessToken, "Content-Type": "application/json" },
      body: JSON.stringify({ values: [["CATEGORY", "TASK", "TIMESTAMP"]] })
    });

    // 2. Ab Headers ko BOLD aur BADA (Size 12) karo
    await fetchJson(`https://sheets.googleapis.com/v4/spreadsheets/${state.currentSheetId}:batchUpdate`, {
      method: "POST",
      headers: { Authorization: "Bearer " + state.accessToken, "Content-Type": "application/json" },
      body: JSON.stringify({
        requests: [{
          repeatCell: {
            range: { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 3 },
            cell: {
              userEnteredFormat: {
                textFormat: { bold: true, fontSize: 12 },
                backgroundColor: { red: 0.9, green: 0.9, blue: 0.9 } // Light Gray BG
              }
            },
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
    }
  }

  async function appendTask() {
    const rawText = elements.taskInput.value.trim();
    if (!rawText) return showToast("Nothing to save!", "error");
    showToast("Saving task...");

    try {
      if (!state.currentSheetId) {
          if (!(await findExistingSheet())) await createSheet();
      }
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
      showToast(`Saved to ${category}!`);
      elements.taskInput.value = "";
    } catch (e) {
      if (e.message.toLowerCase().includes("not found")) {
        state.currentSheetId = "";
        await appendTask();
      } else showToast(e.message, "error");
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
          elements.authSection.style.display = "none";
          elements.mainApp.style.display = "block";
          elements.logoutButton.style.display = "inline-block";
          findExistingSheet().then(() => updateSheetUi());
          showToast("Connected to Google");
        }
      });
    };
    document.head.appendChild(script);
  }

  function initSpeech() {
    const Speech = window.SpeechRecognition || window.webkitSpeechRecognition;
    state.recognition = new Speech();
    state.recognition.onstart = () => { startVisualizer(); };
    state.recognition.onresult = (e) => { elements.taskInput.value = e.results[0][0].transcript; };
    state.recognition.onend = () => showToast("Mic stopped");
  }

  elements.loginButton.onclick = () => state.tokenClient.requestAccessToken({ prompt: "" });
  elements.logoutButton.onclick = () => { state.accessToken = ""; location.reload(); };
  elements.micButton.onclick = () => { 
    state.recognition.lang = elements.langSelect.value;
    state.recognition.start(); 
  };
  elements.saveTextButton.onclick = appendTask;

  initAuth();
  initSpeech();
})();
