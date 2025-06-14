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
 * @param {boolean} showUI Whether to show UI alerts
 */
function manageTriggers(showUI) {
  if (typeof showUI === 'undefined') showUI = true;
  const ui = SpreadsheetApp.getUi();
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const existingTrigger = triggers.find(function(t) {
      return t.getHandlerFunction() === 'fetchTrendingVideos' &&
             t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK &&
             t.getEventType() === ScriptApp.EventType.CLOCK;
    });
    if (existingTrigger) {
      if (showUI) {
        const response = ui.alert('Trigger Management',
          'A daily trigger already exists.\n\nNext run: ' + existingTrigger.getHandlerFunction() +
          ' at ' + existingTrigger.getTriggerSource() + '\n\nWould you like to remove it?',
          ui.ButtonSet.YES_NO);
        if (response === ui.Button.YES) {
          ScriptApp.deleteTrigger(existingTrigger);
          ui.alert('Trigger removed successfully.');
        }
      }
    } else if (showUI) {
      const response = ui.alert('Add Daily Trigger',
        'Add a daily trigger to automatically fetch trending videos?',
        ui.ButtonSet.YES_NO);
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

