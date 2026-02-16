function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Tenure Rank')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * RUN THIS FUNCTION ONCE TO AUTHORIZE PERMISSIONS
 * This function makes dummy calls to force the "Review Permissions" dialog.
 */
function authorizeScript() {
  console.log("Attempting to trigger AdminDirectory auth...");
  if (typeof AdminDirectory !== 'undefined') {
    AdminDirectory.Users.list({customer: 'my_customer', maxResults: 1});
    console.log("AdminDirectory authorized.");
  } else {
    console.log("AdminDirectory service not valid.");
  }

  console.log("Attempting to trigger People API auth...");
  if (typeof People !== 'undefined') {
    People.People.listDirectoryPeople({
      readMask: 'photos', 
      sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'], 
      pageSize: 1
    });
    console.log("People API authorized.");
  } else {
    console.log("People API service not valid.");
  }
}

// Helper to manage chunked caching (max 100KB per key)
function getCache(key) {
  var cache = CacheService.getScriptCache();
  var json = cache.get(key);
  if (!json) return null;
  
  // Check if it's a chunked pointer
  if (json.indexOf("chunked:") === 0) {
    var chunksCount = parseInt(json.split(":")[1]);
    var fullString = "";
    for (var i = 0; i < chunksCount; i++) {
      var chunk = cache.get(key + "_" + i);
      if (!chunk) return null; // Missing a chunk, invalidate all
      fullString += chunk;
    }
    return JSON.parse(fullString);
  }
  
  return JSON.parse(json);
}

function putCache(key, data, time) {
  var cache = CacheService.getScriptCache();
  var json = JSON.stringify(data);
  var size = json.length;
  var limit = 90000; // Safe limit (under 100KB)
  
  if (size < limit) {
    cache.put(key, json, time);
  } else {
    // Split into chunks
    var chunks = [];
    var chunkId = 0;
    for (var i = 0; i < size; i += limit) {
      cache.put(key + "_" + chunkId, json.substring(i, i + limit), time);
      chunkId++;
    }
    // Store the pointer
    cache.put(key, "chunked:" + chunkId, time);
  }
}

function getData() {
  var cachedData = getCache("employeesData");
  
  if (cachedData) {
    // Sliding Expiration: Refresh the cache
    putCache("employeesData", cachedData, 21600);
    return {
      data: cachedData,
      logs: []
    };
  }

  // 1. OPEN THE SHEET
  var sheetId = 'YOUR-SHEET-ID-HERE';
  var tabName = 'TenureRank';
  
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(tabName);
    
    // Error check: Does tab exist?
    if (!sheet) {
      throw new Error('Tab named "' + tabName + '" was not found. Please check the spelling on the bottom tab of your spreadsheet.');
    }
    
    // 2. GET DATA
    var lastRow = sheet.getLastRow();
    // If no data beyond header (row 1), return empty
    if (lastRow < 2) return []; 
    
    // Get Columns A (Email) and B (Start Date) starting from row 2
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    // 3. PROCESS DATA (TEXT ONLY)
    var employees = values.map(function(row) {
      var email = row[0]; // Column A
      var rawDate = row[1]; // Column B
      
      // Basic Email Check
      if (!email || String(email).trim() === "") return null;

      // Extract Name
      var name = "Team Member";
      try {
        var userPart = email.split('@')[0];
        name = userPart.split('.').map(function(part) {
          return part.charAt(0).toUpperCase() + part.slice(1);
        }).join(' ');
      } catch (e) {
        name = email;
      }
      
      // Robust Date Parsing
      var parsedDate = null;
      var isValidDate = false;
      
      if (rawDate) {
        if (rawDate instanceof Date) {
          parsedDate = rawDate;
          isValidDate = true;
        } else {
          // Try to parse string date
          parsedDate = new Date(rawDate);
          isValidDate = !isNaN(parsedDate.getTime());
        }
      }

      var normalizedEmail = String(email).toLowerCase();
      return {
        email: normalizedEmail,
        name: name,
        startDate: parsedDate,
        isValidDate: isValidDate
      };
    });
    
    // Remove null rows
    employees = employees.filter(function(e) { return e !== null; });
    
    // 4. SORT
    employees.sort(function(a, b) {
      if (!a.isValidDate) return 1;
      if (!b.isValidDate) return -1;
      return a.startDate - b.startDate;
    });
    
    // 5. FORMAT FOR FRONTEND
    var now = new Date();
    
    var formattedData = employees.map(function(e, index) {
      var displayDate = "Unknown Date";
      var tenureString = "";
      
      if (e.isValidDate) {
        displayDate = Utilities.formatDate(e.startDate, Session.getScriptTimeZone(), "MMM d, yyyy");
        
        // Calculate Tenure
        var totalMonths = (now.getFullYear() - e.startDate.getFullYear()) * 12 + (now.getMonth() - e.startDate.getMonth());
        if (now.getDate() < e.startDate.getDate()) totalMonths--;
        
        var years = Math.floor(totalMonths / 12);
        var months = totalMonths % 12;
        
        var parts = [];
        if (years > 0) parts.push(years + " " + (years === 1 ? "year" : "years"));
        if (months > 0) parts.push(months + " " + (months === 1 ? "month" : "months"));
        
        if (parts.length === 0) {
           var diffTime = now - e.startDate;
           var days = Math.floor(diffTime / (1000 * 60 * 60 * 24));
           tenureString = days + " days";
        } else {
           tenureString = parts.join(", ");
        }
      }
      
      return {
        rank: index + 1,
        name: e.name,
        email: e.email,
        formattedDate: displayDate,
        tenure: tenureString
      };
    });
    
    // Cache the processed data
    putCache("employeesData", formattedData, 21600);

    return {
      data: formattedData,
      logs: []
    };
    
  } catch (e) {
    throw new Error(e.toString());
  }
}

/**
 * Fetch photos for a specific list of emails.
 * Checks Cache -> Bulk API -> Individual Fallback
 */
function fetchPhotos(emails) {
    if (!emails || emails.length === 0) return {};
    
    var photoMap = {};
    var debugLogs = [];
    var photosLoaded = false;
    
    // 1. Check Cache
    var cachedPhotos = getCache("photoMap");
    if (cachedPhotos) {
        photoMap = cachedPhotos;
        // Sliding Expiration
        putCache("photoMap", photoMap, 21600);
    }
    
    // If cache missed or empty, fetch
    if (Object.keys(photoMap).length === 0) {
        // 2. Admin Directory Bulk
        try {
          if (typeof AdminDirectory !== 'undefined') {
             var pageToken;
             do {
               var response = AdminDirectory.Users.list({
                 customer: 'my_customer',
                 maxResults: 500,
                 viewType: 'domain_public',
                 fields: 'users(primaryEmail,thumbnailPhotoUrl),nextPageToken',
                 pageToken: pageToken
               });
               if (response.users) {
                 response.users.forEach(function(user) {
                   if (user.primaryEmail && user.thumbnailPhotoUrl) {
                     photoMap[user.primaryEmail.toLowerCase()] = user.thumbnailPhotoUrl;
                   }
                 });
                 photosLoaded = true;
               }
               pageToken = response.nextPageToken;
             } while (pageToken);
          } else {
            debugLogs.push("AdminDirectory service is NOT enabled.");
          }
        } catch (e) {
          debugLogs.push("AdminDirectory failed: " + e.message);
        }

        // 3. People API Bulk
        if (!photosLoaded) {
          try {
            if (typeof People !== 'undefined') {
              var pageToken;
              do {
                var response = People.People.listDirectoryPeople({
                  readMask: 'emailAddresses,photos',
                  sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
                  pageSize: 1000,
                  pageToken: pageToken
                });
                
                if (response.people) {
                  response.people.forEach(function(person) {
                    if (person.emailAddresses && person.photos) {
                      var url = person.photos[0].url;
                      person.emailAddresses.forEach(function(emailObj) {
                        photoMap[emailObj.value.toLowerCase()] = url;
                      });
                    }
                  });
                  photosLoaded = true;
                }
                pageToken = response.nextPageToken;
              } while (pageToken);
            } else {
               debugLogs.push("People API not enabled.");
            }
          } catch (e) {
            debugLogs.push("People API failed: " + e.message);
          }
        }
        
        // Save bulk result to cache
        if (Object.keys(photoMap).length > 0) {
           putCache("photoMap", photoMap, 21600);
        }
    }
    
    // 4. Individual Fallback
    var updatedCache = false;
    emails.forEach(function(email) {
       var normEmail = email.toLowerCase();
       if (!photoMap[normEmail]) {
          // Try individual fetch
          try {
            if (typeof AdminDirectory !== 'undefined') {
              var user = AdminDirectory.Users.get(email, {viewType: 'domain_public'});
              if (user && user.thumbnailPhotoUrl) {
                photoMap[normEmail] = user.thumbnailPhotoUrl;
                updatedCache = true;
              }
            }
          } catch (e) {}
       }
    });
    
    if (updatedCache) {
       putCache("photoMap", photoMap, 21600);
    }

    return {
        photos: photoMap,
        logs: debugLogs
    };
}
    
