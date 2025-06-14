function fetchTrendingVideos() {
  var settings = config('settings') || {};
  if (!settings.apiKey) throw new Error('API key not configured');
  var request = {
    apiKey: settings.apiKey,
    endpoint: 'videos',
    part: 'snippet',
    chart: 'mostPopular',
    regionCode: settings.regionCode || 'US',
    maxResults: settings.maxResults || 25
  };
  var url = buildApiRequestUrl_(request);
  var response = fetchWithRetry(url, { muteHttpExceptions: true });
  var data = JSON.parse(response.getContentText());
  if (data.error) throw new Error(data.error.message || 'API error');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('VIDEOS');
  if (!sheet) {
    sheet = ss.insertSheet('VIDEOS');
    sheet.appendRow(['video_id','title','channel','published_at','status','fetched_at']);
  }
  var lastRow = sheet.getLastRow();
  var existingIds = [];
  if (lastRow > 1) {
    existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r){ return r[0]; });
  }
  var rows = processApiResponse_(data, ['video_id','title','channel','published_at','status','fetched_at'], existingIds);
  if (rows.length) {
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
  }
  logToSheet('LOGS',{level:'INFO',message:'Fetched '+rows.length+' videos'});
  return rows.length;
}

function refreshData() {
  return fetchTrendingVideos();
}

function showConfigDialog() {
  var settings = config('settings') || {};
  var template = HtmlService.createTemplateFromFile('configDialogUi');
  template.settings = settings;
  var html = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Settings');
}

function saveConfiguration(cfg) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    config('settings', cfg);
  } finally {
    lock.releaseLock();
  }
  return true;
}

function testApiConnection(apiKey) {
  var url = buildApiRequestUrl_({apiKey: apiKey, endpoint:'videos', part:'id', chart:'mostPopular', maxResults:1});
  var response = fetchWithRetry(url,{muteHttpExceptions:true});
  var data = JSON.parse(response.getContentText());
  if (data.error) throw new Error(data.error.message || 'API error');
  return true;
}

function showStatsDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('statsDashboardUi').setTitle('YouTube Trends Dashboard');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getDashboardData() {
  return {
    success: true,
    apiStatus: {healthy: true, message: 'OK'},
    dataStatus: {healthy: true, message: 'OK'},
    lastSync: new Date().toISOString(),
    recentActivity: [],
    viewsData: [],
    engagementData: [],
    growthData: []
  };
}

function showLogsViewer() {
  var html = HtmlService.createHtmlOutputFromFile('logsViewerUi').setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html,'Logs Viewer');
}

function getLogs(page) {
  page = page || 1;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOGS');
  if (!sheet) return {logs:[], totalPages:1};
  var values = sheet.getDataRange().getValues();
  var logs = [];
  for (var i=1;i<values.length;i++) {
    logs.push({
      id:i,
      timestamp: values[i][0],
      level: values[i][1],
      message: values[i][2],
      payload: JSON.parse(values[i][3] || '{}')
    });
  }
  var pageSize = 20;
  var totalPages = Math.max(1, Math.ceil(logs.length / pageSize));
  var start = (page-1)*pageSize;
  var slice = logs.slice(start,start+pageSize);
  return {logs:slice, totalPages: totalPages};
}

function clearLogs() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOGS');
    if (sheet) sheet.clear();
  } finally {
    lock.releaseLock();
  }
  return true;
}

