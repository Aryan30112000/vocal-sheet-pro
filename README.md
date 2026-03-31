# Task To Speech

A simple static web app that:

- signs in with Google OAuth
- reads the logged-in user's email and access token
- captures spoken text from the microphone
- creates a new Google Sheet in the user's Google Drive
- adds `Task` and `Timestamp` columns
- appends each spoken or typed task into the sheet

## Files

- `index.html` - app markup
- `styles.css` - simple UI styling
- `app.js` - Google OAuth, speech recognition, and Sheets API logic
- `config.js` - your Google client configuration

## Google setup

1. Open Google Cloud Console and create a project.
2. Enable these APIs:
   - Google Sheets API
   - Google Drive API
3. Create an OAuth 2.0 Client ID for a Web application.
4. Add your local origin to the OAuth client:
   - `http://localhost:8000`
   - or whichever port you use
5. Update `config.js` with:
   - `googleClientId`

## Run locally

Because Google OAuth requires an allowed web origin, run the app from a local server instead of opening the HTML file directly.

If Python is installed:

```bash
python -m http.server 8000
```

Then open:

```text
http://localhost:8000
```

## Notes

- The app uses the browser Web Speech API, so Chrome or Edge is recommended.
- Creating a spreadsheet with the Sheets API automatically places it in the signed-in user's Google Drive.
- The app uses the OAuth access token directly for Google API calls, so a separate API key is not required here.
