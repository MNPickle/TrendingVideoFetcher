# ScriptBuilderProject

A Google Sheets automation tool that fetches trending YouTube videos via the YouTube Data API v3, stores them in a dedicated sheet, and provides management features through a custom menu and UI dashboards.

---

## Table of Contents

1. [Project Overview](#project-overview)  
2. [Architecture & Flow](#architecture--flow)  
3. [Features](#features)  
4. [Installation](#installation)  
5. [Configuration](#configuration)  
6. [Usage](#usage)  
7. [Components](#components)  
8. [Triggers & Custom Menu](#triggers--custom-menu)  
9. [Fail-Safes & Quota Management](#fail-safes--quota-management)  
10. [Developer Notes](#developer-notes)  
11. [Dependencies](#dependencies)  
12. [License](#license)  

---

## Project Overview

**ScriptBuilderProject** is a Google Apps Script project bound to a Google Sheet. It:

- Fetches the latest "most popular" YouTube videos by region and category.  
- Appends new videos (no duplicates) to a `VIDEOS` sheet.  
- Tracks fetch status and errors in a `LOGS` sheet.  
- Provides a statistics dashboard and error-log viewer.  
- Supports both daily automatic triggers and on-demand manual fetches via a custom menu.

---

## Architecture & Flow

Client-server architecture inside Google Sheets:

1. **User Interface**  
   - Configuration dialog  
   - Statistics dashboard  
   - Logs viewer  
   - Persistent sidebar  

2. **Backend (Apps Script)**  
   - `mainSetup.gs`: Initializes menu & triggers  
   - `youtubeApiClient.gs`: Builds API requests, parses responses  
   - `sharedUtilities.gs`: HTTP wrapper, config management, logging  

3. **Data Storage (Sheets)**  
   - **CONFIG**: Stores API key & settings  
   - **VIDEOS**: Appended/updated video records  
   - **LOGS**: Timestamped log entries  

Flow:

1. User configures settings via UI → stored in `CONFIG`.  
2. A daily time-based trigger (4:00 UTC) and/or manual menu action invokes fetch.  
3. `youtubeApiClient.gs` calls YouTube Data API → processes JSON.  
4. New videos are appended; duplicates marked `DUPLICATE`.  
5. Statistics and logs are viewable via sidebars/dashboards.

---

## Features

- Custom menu with:
  - 🚀 Fetch Trending Videos Now  
  - ⚙️ Configure API Key & Settings  
  - 📊 View Statistics Dashboard  
  - 🔍 View Error Logs  
  - ♻️ Reset Last Fetch Timestamp  

- Automatic daily fetch at 04:00 UTC  
- Manual on-demand fetch  
- Duplicate detection by `video_id`  
- Detailed error logging (with stack traces & payloads)  
- Exponential back-off on HTTP errors (403/500)  
- Hard timeout of 30 seconds per request  
- Stats panel: total rows, new in last 24 h, error count  

---

## Installation

1. Create or open a Google Sheet.  
2. From the menu: **Extensions → Apps Script**.  
3. In the Apps Script editor, create the following files and copy the provided code:  
   - `mainSetup.gs`  
   - `youtubeApiClient.gs`  
   - `sharedUtilities.gs`  
   - `configDialogUi.html`  
   - `statsDashboardUi.html`  
   - `logsViewerUi.html`  
   - `sidebarUi.html`  
   - `spinnerUi.html`  
   - `uiComponentsStyles.html`  
4. Enable the **YouTube Data API v3** in the Google Cloud project linked to your script.  
5. Save all files and refresh the Google Sheet.  

---

## Configuration

1. In your Sheet: **Trending Video Fetcher → ⚙️ Configure API Key / Settings**  
2. In the modal, enter:
   - `YT_API_KEY` (YouTube Data API v3 key)  
   - `REGION` (e.g., `US`, `GB`, `IN`; default `US`)  
   - `CATEGORY_ID` (optional numeric category; leave blank for all)  
   - `MAX_RESULTS` (1–50; default 25)  
3. Click **Save**. The script will:
   - Store values in the `CONFIG` sheet.  
   - Install the daily trigger.  
   - Perform a test API call.  
   - Show success or detailed error.

---

## Usage

### Manual Fetch

From the Sheet menu:
```
Trending Video Fetcher → 🚀 Fetch Trending Videos Now
```
This invokes `runFetchNow()` and updates the `VIDEOS` sheet.

### View Statistics

```
Trending Video Fetcher → 📊 View Stats
```
Opens a sidebar with:
- Total videos fetched  
- New videos in last 24 h  
- Error count  

### View Error Logs

```
Trending Video Fetcher → 🔍 View Error Logs
```
Browse and clear log entries.

### Reset Last Fetch

```
Trending Video Fetcher → ♻️ Reset LAST_FETCH_UTC
```
Clears the timestamp in `CONFIG → LAST_FETCH_UTC`, forcing a full fetch on next run.

---

## Components

.gs Files:
- **mainSetup.gs**  
  Core initialization, custom menu creation, trigger setup, menu action handlers.  
- **youtubeApiClient.gs**  
  Builds YouTube API requests, parses responses, writes to `VIDEOS` sheet, marks duplicates.  
- **sharedUtilities.gs**  
  HTTP request wrapper with retry, config getters/setters, structured logging to `LOGS` sheet, date/time helpers.  

.html Files:
- **configDialogUi.html**  
  Modal form for configuration with client-side validation and test-API button.  
- **statsDashboardUi.html**  
  Displays charts and summary metrics.  
- **logsViewerUi.html**  
  Filterable table of log entries with expandable JSON payloads.  
- **sidebarUi.html**  
  Persistent navigation sidebar for quick access to fetch, stats, logs.  
- **spinnerUi.html**  
  Loading spinner component for async operations.  
- **uiComponentsStyles.html**  
  Centralized CSS for all UI dialogs and sidebars.  

---

## Triggers & Custom Menu

**Time-based Trigger**  
- Daily at 04:00 UTC → `fetchTrendingYouTubeVideos()`

**Manual Trigger**  
- Custom Menu → 🚀 Fetch Now

Menu items installed via `onOpen()` in `mainSetup.gs`.

---

## Fail-Safes & Quota Management

- Each API call ≈ 1 quota unit; target < 10 k/day.  
- Exponential back-off for HTTP 403/500.  
- Hard timeout of 30 s per request.  
- All exceptions logged with full stack trace.  

---

## Developer Notes

- Use ES5 syntax only (Apps Script restriction).  
- Always fetch API key & settings from `CONFIG` sheet—no hard-coding.  
- Avoid duplicates via `video_id` index on the `VIDEOS` sheet.  
- UI styling centralized in `uiComponentsStyles.html`.  
- Sidebar & spinner are implemented via their own HTML files.

---

## Dependencies

- Google Apps Script (bound to Google Sheets)  
- YouTube Data API v3 (enable in Google Cloud)  
- No external libraries; uses built-in `UrlFetchApp`, `SpreadsheetApp`, `HtmlService`.

---

## License

This project is released under the MIT License. See `LICENSE.md` for details.