/**
 * Serves the HTML to the browser.
 */
function doGet() {
  // Matches the file name 'Index.html' exactly
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Google Group Inspector')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Main function called by the UI.
 * Recursively fetches owners, managers, nested groups, and all unique members.
 * FILTERS: Only returns emails ending in @companydomain.com for privacy.
 * COUNTS: Tracks how many external members were hidden.
 */
function getGroupLeadership(groupEmail) {
  try {
    var response = {
      success: true,
      owners: [],
      managers: [],
      nestedGroups: [],
      allMembers: [],
      hiddenCount: 0
    };
    
    // Sets to track unique values
    var processedGroups = new Set();
    var uniqueMembers = new Set();
    var nestedGroupsSet = new Set();
    var hiddenMembersSet = new Set(); // Track unique hidden members
    
    // 1. Fetch direct leadership (Owners/Managers) of the requested group
    try {
      var initialMembers = listMembers(groupEmail);
      
      initialMembers.forEach(function(member) {
        // SECURITY FILTER: Check for internal @companydomain.com addresses
        if (!member.email.endsWith('@companydomain.com')) {
          hiddenMembersSet.add(member.email);
          return;
        }

        if (member.role === 'OWNER') {
          response.owners.push(member.email);
        } else if (member.role === 'MANAGER') {
          response.managers.push(member.email);
        }
      });
    } catch (e) {
       return { success: false, error: "Group not found or access denied: " + groupEmail };
    }

    // 2. Recursive Fetch for All Members and Nested Groups
    fetchMembersRecursive(groupEmail, processedGroups, uniqueMembers, nestedGroupsSet, hiddenMembersSet);

    // Convert Sets to Arrays for the response
    response.allMembers = Array.from(uniqueMembers).sort();
    response.nestedGroups = Array.from(nestedGroupsSet).sort();
    response.hiddenCount = hiddenMembersSet.size; // Return count of unique hidden members
    response.owners.sort();
    response.managers.sort();
    
    return response;

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Helper function to recursively fetch members.
 */
function fetchMembersRecursive(groupEmail, processedGroups, uniqueMembers, nestedGroupsSet, hiddenMembersSet) {
  // Prevent infinite loops (circular dependencies)
  if (processedGroups.has(groupEmail)) return;
  processedGroups.add(groupEmail);

  try {
    var members = listMembers(groupEmail);

    if (members && members.length > 0) {
      members.forEach(function(member) {
        var email = member.email;
        var type = member.type; // 'USER', 'GROUP', or 'CUSTOMER'

        // SECURITY FILTER: Check for internal Company Domain addresses
        if (!email.endsWith('@copmanydomain.com')) {
           // If it's a user, count them as hidden
           if (type === 'USER') {
             hiddenMembersSet.add(email);
           }
           return;
        }

        if (type === 'GROUP') {
          nestedGroupsSet.add(email);
          // RECURSE: Dive into this nested group
          fetchMembersRecursive(email, processedGroups, uniqueMembers, nestedGroupsSet, hiddenMembersSet);
        } else if (type === 'USER') {
          uniqueMembers.add(email);
        }
      });
    }
  } catch (e) {
    console.warn("Could not fetch members for nested group: " + groupEmail + ". Error: " + e.toString());
    // We continue execution even if one nested group fails
  }
}

/**
 * Wrapper for AdminDirectory API to handle pagination.
 * REQUIRED: Enable "Admin SDK API" in Apps Script Services.
 */
function listMembers(groupKey) {
  var allMembers = [];
  var pageToken;
  
  do {
    var page = AdminDirectory.Members.list(groupKey, {
      pageToken: pageToken,
      maxResults: 200
    });
    
    if (page.members) {
      allMembers = allMembers.concat(page.members);
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  
  return allMembers;
}
