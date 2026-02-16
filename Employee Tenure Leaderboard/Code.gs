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
    var chunkId = 0;
    for (var i = 0; i < size; i += limit) {
      cache.put(key + "_" + chunkId, json.substring(i, i + limit), time);
      chunkId++;
    }
    // Store the pointer
    cache.put(key, "chunked:" + chunkId, time);
  }
}

/**
 * Combined endpoint: returns employee data AND photos in a single round-trip.
 * This eliminates a second google.script.run call (~2s saved).
 */
function getPageData() {
  var result = getData();
  var emails = (result.data || []).map(function(e) { return e.email; });

  // Wrap photo fetch so cards always render even if photos fail
  var photos = {};
  var photoLogs = [];
  try {
    var photoResult = fetchPhotosForEmails_(emails);
    photos = photoResult.photos || {};
    photoLogs = photoResult.logs || [];
  } catch (e) {
    photoLogs.push("Photo fetch failed: " + e.message);
  }

  return {
    data: result.data,
    photos: photos,
    logs: (result.logs || []).concat(photoLogs)
  };
}

function getData() {
  var cachedData = getCache("employeesData");
  
  if (cachedData) {
    return {
      data: cachedData,
      logs: []
    };
  }

  var sheetId = 'YOUR-SHEET-ID';
  var tabName = 'TenureRank';
  
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(tabName);
    
    if (!sheet) {
      throw new Error('Tab named "' + tabName + '" was not found.');
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { data: [], logs: [] }; 
    
    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    var employees = values.map(function(row) {
      var email = row[0];
      var rawDate = row[1];
      
      if (!email || String(email).trim() === "") return null;

      var name = "Team Member";
      try {
        var userPart = email.split('@')[0];
        name = userPart.split('.').map(function(part) {
          return part.charAt(0).toUpperCase() + part.slice(1);
        }).join(' ');
      } catch (e) {
        name = email;
      }
      
      var parsedDate = null;
      var isValidDate = false;
      
      if (rawDate) {
        if (rawDate instanceof Date) {
          parsedDate = rawDate;
          isValidDate = true;
        } else {
          parsedDate = new Date(rawDate);
          isValidDate = !isNaN(parsedDate.getTime());
        }
      }

      return {
        email: String(email).toLowerCase(),
        name: name,
        startDate: parsedDate,
        isValidDate: isValidDate
      };
    });
    
    employees = employees.filter(function(e) { return e !== null; });
    
    employees.sort(function(a, b) {
      if (!a.isValidDate) return 1;
      if (!b.isValidDate) return -1;
      return a.startDate - b.startDate;
    });
    
    var now = new Date();
    // Hoist timezone lookup out of the loop (was called once per employee before)
    var tz = Session.getScriptTimeZone();
    
    var formattedData = employees.map(function(e, index) {
      var displayDate = "Unknown Date";
      var tenureString = "";
      
      if (e.isValidDate) {
        displayDate = Utilities.formatDate(e.startDate, tz, "MMM d, yyyy");
        
        // Guard against future start dates
        if (e.startDate > now) {
          tenureString = "Starting soon";
        } else {
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
      }
      
      return {
        rank: index + 1,
        name: e.name,
        email: e.email,
        formattedDate: displayDate,
        tenure: tenureString
      };
    });
    
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
 * Internal photo fetcher — fetches bulk photos and filters to only requested emails.
 * No individual fallback loop (users without photos get the ui-avatars default).
 */
function fetchPhotosForEmails_(emails) {
    if (!emails || emails.length === 0) return { photos: {}, logs: [] };
    
    var fullPhotoMap = {};
    var debugLogs = [];
    var photosLoaded = false;
    
    // Build a quick lookup set for the emails we care about
    var emailSet = {};
    emails.forEach(function(e) { emailSet[e.toLowerCase()] = true; });
    
    // 1. Check Cache
    var cachedPhotos = getCache("photoMap");
    if (cachedPhotos) {
        fullPhotoMap = cachedPhotos;
    }
    
    // 2. If cache missed, fetch from APIs
    if (Object.keys(fullPhotoMap).length === 0) {
        // Admin Directory Bulk
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
                     fullPhotoMap[user.primaryEmail.toLowerCase()] = user.thumbnailPhotoUrl;
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

        // People API fallback (only if AdminDirectory didn't work)
        if (!photosLoaded) {
          try {
            if (typeof People !== 'undefined') {
              var pageToken2;
              do {
                var resp = People.People.listDirectoryPeople({
                  readMask: 'emailAddresses,photos',
                  sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],
                  pageSize: 1000,
                  pageToken: pageToken2
                });
                
                if (resp.people) {
                  resp.people.forEach(function(person) {
                    if (person.emailAddresses && person.photos) {
                      var url = person.photos[0].url;
                      person.emailAddresses.forEach(function(emailObj) {
                        fullPhotoMap[emailObj.value.toLowerCase()] = url;
                      });
                    }
                  });
                  photosLoaded = true;
                }
                pageToken2 = resp.nextPageToken;
              } while (pageToken2);
            } else {
               debugLogs.push("People API not enabled.");
            }
          } catch (e) {
            debugLogs.push("People API failed: " + e.message);
          }
        }
        
        // Cache the full domain map for next time
        if (Object.keys(fullPhotoMap).length > 0) {
           putCache("photoMap", fullPhotoMap, 21600);
        }
    }
    
    // 3. Filter to only the emails on the leaderboard (smaller payload)
    var filteredPhotos = {};
    for (var email in fullPhotoMap) {
      if (emailSet[email]) {
        filteredPhotos[email] = fullPhotoMap[email];
      }
    }

    return {
        photos: filteredPhotos,
        logs: debugLogs
    };
}

/** @deprecated Use getPageData() instead — kept for backward compatibility */
function fetchPhotos(emails) {
    return fetchPhotosForEmails_(emails);
}
