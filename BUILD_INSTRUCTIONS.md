# ScriptBuilderProject Build Instructions

## 1. Prerequisites & Environment Setup

### 1.1 Google Account & Permissions
- A Google account with permission to create/edit Google Sheets and Apps Script projects.
- Access to Google Cloud Console for API enablement.

### 1.2 Create & Bind the Script
1. Open (or create) a Google Sheet.
2. Go to **Extensions → Apps Script**.  
3. Rename the default project to **ScriptBuilderProject**.

### 1.3 Enable APIs & Services
1. In the Apps Script editor, click the **Services** icon (left toolbar).  
2. Add **YouTube Data API v3**.
3. In the Cloud Console (linked to your script’s project):
   - Enable **YouTube Data API**.
   - Configure the OAuth consent screen.
   - (If you plan to deploy externally) create OAuth 2.0 client credentials.

### 1.4 Configure `appsscript.json`
Open **View → Show manifest file** and ensure:
```json
{
  "timeZone": "Etc/GMT",
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/youtube.readonly"
  ],
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "YouTube",
        "serviceId": "youtube",
        "version": "v3"
      }
    ]
  }
}
```

---

## 2. Step-by-Step Build Process

> **Note**: Google Apps Script loads `.gs` source files in alphabetical order. To enforce execution order, prefix filenames with numbers or rely on a single entry point (e.g. `onOpen` in **mainSetup.gs**).

### 2.1 Execution Order 0: sharedUtilities.gs
1. Click **+ → Script**, rename to `sharedUtilities.gs`.
2. Paste shared helper functions:
   ```javascript
   // sharedUtilities.gs
   function logDebug(message) {
     Logger.log('[DEBUG] ' + message);
   }
   function setConfig(key, value) {
     PropertiesService.getDocumentProperties().setProperty(key, value);
   }
   function getConfig(key) {
     return PropertiesService.getDocumentProperties().getProperty(key);
   }
   ```

### 2.2 Execution Order 1: mainSetup.gs
1. **+ → Script**, rename to `mainSetup.gs`.
2. Paste UI‐integration and menu setup:
   ```javascript
   // mainSetup.gs
   function onOpen() {
     SpreadsheetApp.getUi()
       .createMenu('ScriptBuilder')
       .addItem('Configure…', 'showConfigDialog')
       .addItem('Open Sidebar', 'showSidebar')
       .addSeparator()
       .addItem('View Logs', 'showLogsViewer')
       .addToUi();
   }
   
   function showConfigDialog() {
     const html = HtmlService
       .createHtmlOutputFromFile('configDialogUi')
       .setWidth(400).setHeight(300);
     SpreadsheetApp.getUi().showModalDialog(html, 'Configure ScriptBuilder');
   }
   
   function showSidebar() {
     const html = HtmlService
       .createHtmlOutputFromFile('sidebarUi')
       .setTitle('ScriptBuilder Sidebar');
     SpreadsheetApp.getUi().showSidebar(html);
   }

   function showLogsViewer() {
     const html = HtmlService
       .createHtmlOutputFromFile('logsViewerUi')
       .setWidth(600).setHeight(400);
     SpreadsheetApp.getUi().showModalDialog(html, 'Execution Logs');
   }
   ```

### 2.3 Execution Order 2: youtubeApiClient.gs
1. **+ → Script**, rename to `youtubeApiClient.gs`.
2. Paste YouTube API wrapper:
   ```javascript
   // youtubeApiClient.gs
   function fetchChannelStats(channelId) {
     logDebug('Fetching stats for ' + channelId);
     const resp = YouTube.Channels.list('statistics', { id: channelId });
     if (!resp.items || !resp.items.length) {
       throw new Error('No data for channel ' + channelId);
     }
     return resp.items[0].statistics;
   }
   ```

### 2.4 Execution Order 3: configDialogUi.html
1. **+ → HTML**, name it `configDialogUi`.
2. Paste:
   ```html
   <!-- configDialogUi.html -->
   <!DOCTYPE html>
   <html>
     <head>
       <base target="_top">
       <?!= HtmlService.createHtmlOutputFromFile('uiComponentsStyles').getContent(); ?>
     </head>
     <body>
       <form id="configForm">
         <label>Channel ID:</label>
         <input type="text" name="channelId"><br><br>
         <button type="button" onclick="saveConfig()">Save</button>
       </form>
       <script>
         function saveConfig() {
           const form = document.getElementById('configForm');
           google.script.run
             .withSuccessHandler(() => google.script.host.close())
             .setConfig('channelId', form.channelId.value);
         }
       </script>
     </body>
   </html>
   ```

### 2.5 Execution Order 4: statsDashboardUi.html & sidebarUi.html
- **statsDashboardUi.html**:
  ```html
  <!-- statsDashboardUi.html -->
  <!DOCTYPE html>
  <html><head><base target="_top"><?!= HtmlService.createHtmlOutputFromFile('uiComponentsStyles').getContent(); ?></head>
  <body>
    <h3>Channel Statistics</h3>
    <div id="stats"></div>
    <script>
      google.script.run
        .withSuccessHandler(data => {
          document.getElementById('stats').innerText = JSON.stringify(data, null, 2);
        })
        .fetchChannelStats(google.script.run.getConfig('channelId'));
    </script>
  </body>
  </html>
  ```
- **sidebarUi.html**:
  ```html
  <!-- sidebarUi.html -->
  <!DOCTYPE html>
  <html><head><base target="_top"><?!= HtmlService.createHtmlOutputFromFile('uiComponentsStyles').getContent(); ?></head>
  <body>
    <button onclick="google.script.run.showStatsDashboard()">View Stats Dashboard</button>
    <div id="spinner"></div>
    <?!= HtmlService.createHtmlOutputFromFile('spinnerUi').getContent(); ?>
    <script>
      function showStatsDashboard() {
        document.getElementById('spinner').style.display = 'block';
        google.script.run
          .withSuccessHandler(() => {
            document.getElementById('spinner').style.display = 'none';
            google.script.run.showModal(statsDashboardUi);
          })
          .fetchChannelStats(google.script.run.getConfig('channelId'));
      }
    </script>
  </body>
  </html>
  ```

### 2.6 Execution Order 5: logsViewerUi.html & spinnerUi.html
- **logsViewerUi.html**:
  ```html
  <!-- logsViewerUi.html -->
  <!DOCTYPE html>
  <html><head><base target="_top"><?!= HtmlService.createHtmlOutputFromFile('uiComponentsStyles').getContent(); ?></head>
  <body>
    <pre id="logContent">Loading…</pre>
    <script>
      google.script.run
        .withSuccessHandler(txt => document.getElementById('logContent').innerText = txt)
        .getExecutionLogs();
    </script>
  </body>
  </html>
  ```
- **spinnerUi.html**:
  ```html
  <!-- spinnerUi.html -->
  <div class="spinner"></div>
  ```

### 2.7 Execution Order 6: uiComponentsStyles.html
1. **+ → HTML**, name it `uiComponentsStyles`.
2. Paste:
   ```html
   <style>
     .spinner {
       border: 4px solid #f3f3f3;
       border-top: 4px solid #3498db;
       border-radius: 50%;
       width: 24px; height: 24px;
       animation: spin 1s linear infinite;
     }
     @keyframes spin { 100% { transform: rotate(360deg); } }
     body { font-family: Arial, sans-serif; }
   </style>
   ```

---

## 3. Deployment Instructions

1. In Apps Script Editor, click **Deploy → New deployment**.
2. Select **“Container-bound”** (no URL) or **“Web app”** if you need an external URL.
3. Set **Execute as**: `Me` and **Who has access**: `Anyone` (or your domain).
4. Click **Deploy** and grant scopes.
5. If Web app, copy the URL; otherwise, use the sheet’s custom menu.

---

## 4. Testing Procedures

1. **Menu Test**: Reload the sheet; confirm **ScriptBuilder** menu appears.
2. **Config Test**:  
   - Click **Configure…**, enter a valid YouTube Channel ID, click **Save**.  
   - Reopen to confirm the value is stored.
3. **API Test**:
   - Run `fetchChannelStats('UC_x5XG1OV2P6uZZ5FSM9Ttw')` in the Apps Script debugger.
   - Inspect `Logger.log`.
4. **UI Test**:
   - Open sidebar, click **View Stats Dashboard**, verify data appears.
   - Open **View Logs**, confirm logs are populated.
5. **Automated Tests** (optional):
   - Integrate QUnitGS2 or clasp‐based tests.
   - Create `tests.gs` and run via `npm test` (if using clasp).

---

## 5. Troubleshooting Common Issues

- **“ReferenceError: utils is not defined”**  
  Ensure you named your utilities file `sharedUtilities.gs` and that you call its functions by exact name (e.g. `logDebug`, not `utils.logDebug`).

- **API Quota Exceeded / 403 Errors**  
  Check your Cloud Console quotas and billing. Make sure OAuth scopes include `youtube.readonly`.

- **UI Doesn’t Update / Cache Issues**  
  Hard‐reload the container (`Ctrl+F5`) or clear browser cache.  

- **Missing Scopes / Access Denied**  
  Edit your manifest (`appsscript.json`), add required scopes, re-deploy, and re-authorize.

- **Incorrect Load Order**  
  Prefix `.gs` filenames with `0_`, `1_`, etc., or consolidate initialization into `mainSetup.gs`.

---

## 6. Special Considerations for Google Sheet Projects

- Apps Script’s server‐side `.gs` files load alphabetically—order your code accordingly.
- Container‐bound scripts share the sheet’s properties; use `DocumentProperties` for sheet‐specific settings.
- HTML service enforces sandboxing: all client → server calls must use `google.script.run`.
- Keep UI responsive: show spinners (`spinnerUi.html`) during long‐running operations.
- Regularly check **Executions** (in the left menu) for performance and errors.
- Consider publishing as an **Add-on** for broader distribution (requires Google review).