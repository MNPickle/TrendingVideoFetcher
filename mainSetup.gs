function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.createMenu('YouTube Trends')
      .addItem('Fetch Trending Videos', 'fetchTrendingVideos')
      .addItem('Refresh Data', 'refreshData')
      .addItem('View Statistics', 'showStatsDashboard')
      .addItem('View Logs', 'showLogsViewer')
      .addItem('Configure Settings', 'showConfigDialog')
      .addItem('Manage Triggers', 'manageTriggers')
      .addToUi();
    
    manageTriggers(false);
  } catch (e) {
    console.error('Menu setup failed: ' + e.message);
  }
}

/**
 * Manages trigger installation for automatic fetching.
 * @param {boolean} showUI - Whether to show UI alerts (default: true)
 */
function manageTriggers(showUI) {
  if (typeof showUI === 'undefined') showUI = true;
  const ui = SpreadsheetApp.getUi();
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const existingTrigger = triggers.find(t => 
      t.getHandlerFunction() === 'fetchTrendingVideos' && 
      t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK &&
      t.getEventType() === ScriptApp.EventType.CLOCK
    );

    if (existingTrigger) {
      if (showUI) {
        const response = ui.alert(
          'Trigger Management', 
          'A daily trigger already exists.\n\n' +
          'Next run: ' + existingTrigger.getHandlerFunction() + 
          ' at ' + existingTrigger.getTriggerSource() + '\n\n' +
          'Would you like to remove it?',
          ui.ButtonSet.YES_NO
        );
        
        if (response === ui.Button.YES) {
          ScriptApp.deleteTrigger(existingTrigger);
          ui.alert('Trigger removed successfully.');
        }
      }
    } else if (showUI) {
      const response = ui.alert(
        'Add Daily Trigger',
        'Add a daily trigger to automatically fetch trending videos?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        ScriptApp.newTrigger('fetchTrendingVideos')
          .timeBased()
          .everyDays(1)
          .atHour(9)
          .create();
        ui.alert('Daily trigger added successfully.\n\nWill run every day at 9 AM.');
      }
    }
  } catch (e) {
    console.error('Trigger management failed: ' + e.message);
    if (showUI) {
      ui.alert('Trigger Operation Failed', 
        'An error occurred while managing triggers:\n\n' + e.message, 
        ui.ButtonSet.OK);
    }
  }
}

/**
 * Fetches trending videos and logs them to the sheet.
 */
function fetchTrendingVideos() {
  try {
    // TODO: Implement trending videos fetch logic using youtubeApiClient and sharedUtilities
    SpreadsheetApp.getUi().alert('fetchTrendingVideos is not yet implemented.');
  } catch (e) {
    console.error('fetchTrendingVideos failed: ' + e.message);
    SpreadsheetApp.getUi().alert('Error fetching trending videos: ' + e.message);
  }
}

/**
 * Refreshes data by re-fetching trending videos.
 */
function refreshData() {
  try {
    fetchTrendingVideos();
    SpreadsheetApp.getUi().alert('Data refreshed successfully.');
  } catch (e) {
    console.error('refreshData failed: ' + e.message);
    SpreadsheetApp.getUi().alert('Error refreshing data: ' + e.message);
  }
}

/**
 * Shows the configuration dialog.
 */
function showConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('configDialogUi')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Settings');
}

/**
 * Shows the statistics dashboard in a sidebar.
 */
function showStatsDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('statsDashboardUi')
    .setTitle('YouTube Trends Dashboard');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the logs viewer in a sidebar.
 */
function showLogsViewer() {
  const html = HtmlService.createHtmlOutputFromFile('logsViewerUi')
    .setTitle('Logs Viewer');
  SpreadsheetApp.getUi().showSidebar(html);
}