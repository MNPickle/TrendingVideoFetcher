function buildApiRequestUrl_(config) {
  if (!config || typeof config !== 'object') throw new Error("Invalid config object");
  if (!config.apiKey || typeof config.apiKey !== 'string') throw new Error("Missing or invalid API key in config");
  if (!config.endpoint || typeof config.endpoint !== 'string') throw new Error("Missing or invalid API endpoint in config");
  
  const baseUrl = 'https://www.googleapis.com/youtube/v3/';
  let url = `${baseUrl}${encodeURIComponent(config.endpoint)}?key=${encodeURIComponent(config.apiKey)}`;
  
  if (config.part) url += `&part=${encodeURIComponent(config.part)}`;
  if (config.chart) url += `&chart=${encodeURIComponent(config.chart)}`;
  if (config.regionCode) url += `&regionCode=${encodeURIComponent(config.regionCode)}`;
  if (config.maxResults) url += `&maxResults=${encodeURIComponent(config.maxResults)}`;
  
  try {
    new URL(url); // Validate constructed URL
    return url;
  } catch (e) {
    throw new Error(`Invalid URL constructed: ${e.message}`);
  }
}

/**
 * Processes API response into sheet-ready rows
 * @param {Object} response - YouTube API response
 * @param {Array.<string>} columns - Column headers in target sheet
 * @param {Array.<number>} videoIds - Existing video IDs for duplicate detection
 * @return {Array.<Array>} Rows ready for sheet insertion
 */
function processApiResponse_(response, columns, videoIds) {
  if (!response || typeof response !== 'object') throw new Error("Invalid API response");
  if (!Array.isArray(columns)) throw new Error("Invalid column definitions");
  
  if (response.error) {
    throw new Error(`YouTube API Error: ${response.error.message || 'Unknown error'}`);
  }
  
  const rows = [];
  const items = Array.isArray(response.items) ? response.items : [];
  const existingIds = new Set(Array.isArray(videoIds) ? videoIds : []);
  
  items.forEach(item => {
    if (!item || !item.id) return;
    
    const videoId = item.id;
    if (existingIds.has(videoId)) return;
    
    const row = columns.map(column => {
      try {
        switch(column) {
          case 'video_id': 
            return videoId;
          case 'title': 
            return (item.snippet && item.snippet.title) || '[No Title]';
          case 'channel': 
            return (item.snippet && item.snippet.channelTitle) || '[No Channel]';
          case 'published_at': 
            return (item.snippet && item.snippet.publishedAt) || new Date().toISOString();
          case 'status': 
            return 'New';
          case 'fetched_at': 
            return new Date().toISOString();
          default: 
            return '';
        }
      } catch (e) {
        return '[Error]';
      }
    });
    
    rows.push(row);
    existingIds.add(videoId);
  });
  
  return rows;
}

/**
 * Updates status columns for modified rows
 * @param {Range} range - Sheet range to update
 * @param {Array.<string>} statusUpdates - New status values
 * @param {number} statusColIndex - Column index for status
 * @throws {Error} If input parameters are invalid
 */
function updateStatusColumns_(range, statusUpdates, statusColIndex) {
  if (!range || typeof range.getValues !== 'function') throw new Error("Invalid range");
  if (!Array.isArray(statusUpdates)) throw new Error("Invalid status updates");
  if (typeof statusColIndex !== 'number' || statusColIndex < 0 || Math.floor(statusColIndex) !== statusColIndex) {
    throw new Error("Invalid status column index");
  }
  
  const values = range.getValues();
  if (statusUpdates.length !== values.length) {
    throw new Error("Range row count doesn't match status updates count");
  }
  
  const validValues = [];
  statusUpdates.forEach((status, i) => {
    if (i >= values.length) return;
    if (statusColIndex >= values[i].length) {
      throw new Error("Status column index exceeds available columns");
    }
    
    const row = values[i].slice();
    row[statusColIndex] = status;
    validValues.push(row);
  });
  
  range.setValues(validValues);
}