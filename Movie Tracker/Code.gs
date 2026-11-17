// ============================================================================
// MOVIE TRACKER - GOOGLE APPS SCRIPT
// ============================================================================
// Features:
// 1. Rotten Tomatoes Popcornmeter (General Audience Score) instead of Tomatometer
// 2. Full 4-Stage Automated Movie Lifecycle (Columns A–G only, zero extra columns):
//    - Stage 1 ("To be released" - Grey): Release Date > Today
//    - Stage 2 ("In Theaters" - Blue): Release Date <= Today (while in theatrical window)
//      * Daily check writes "VOD: MMM d, yyyy" into your Notes column.
//    - Stage 3 ("New Releases" - Green): Automatically promotes from "In Theaters"
//      the day the "VOD: <Date>" in Notes arrives (or after 60 days in theaters
//      if no VOD date was ever announced).
//    - Stage 4 (Permanent Genre - White): Automatically graduates from "New Releases"
//      to its primary Genre (Action, Comedy, etc.) 6 months after theatrical release.
// 3. Visual Modal Movie Match Fixer (3 easy ways to open, zero extra columns):
//    - Way 1 (Menu): Click any cell on the movie's row -> click
//      "🎬 Movie Tools -> 🎯 Fix Selected Row (Choose Correct Movie)"
//    - Way 2 (Quick '?' Suffix): Add "?" to the end of a movie title in Column A
//      (e.g. "Paul?") and hit Enter -> immediately pops open the visual selector!
//    - Way 3 (Col G Trigger): Type "fix" or "?" into Column G ("Errors") on any
//      row and hit Enter -> immediately pops open the visual selector!
// 4. "Won't Watch" Priority Tier:
//    - Automatically adds "Won't Watch" to your Priority column dropdown (if not already present)
//    - Selecting "Won't Watch" immediately moves the movie row to the bottom of the
//      existing list on the "Won't Watch" tab and removes it from "Movies".
// 5. Safe Linking ("Link Unlinked Movies Only"):
//    - Detects both =HYPERLINK() formulas AND Google Sheets Rich Text links
//    - Uses IMDb's "/suggestion/titles/" endpoint (never matches actors like "Paul")
//    - Never overwrites your typed movie titles in Column A
//    - Detects OMDb daily quota exhaustion ("Request limit reached!") immediately
// ============================================================================

// Store your OMDb API key in Apps Script Project Settings -> Script Properties as "OMDB_API_KEY"
function getOmdbApiKey() {
  try {
    return PropertiesService.getScriptProperties().getProperty("OMDB_API_KEY") || "";
  } catch (e) {
    return "";
  }
}
const apiKey = getOmdbApiKey();

// Track whether OMDb hit its 1,000/day free tier limit during the current execution
let omdbQuotaExceeded = false;

// ============================================================================
// 1. ON EDIT TRIGGER (Installable Trigger: attach to `onMovieEdit`)
// ============================================================================
function onMovieEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = String(sheet.getName() || "").trim();

  // If something is added or edited on the "History" tab, auto-sort newest watch date top-to-bottom
  if (sheetName.toLowerCase() === "history") {
    if (e.range.getRow() <= 1) return;
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return;
    try {
      const editRow = e.range.getRow();
      const titleVal = String(sheet.getRange(editRow, 1).getDisplayValue() || "").trim();
      const dateVal = sheet.getRange(editRow, 2).getValue();
      // If a title was added to History without a watch date, default Col B to today's date
      if (titleVal && (dateVal === "" || dateVal === null || dateVal === undefined)) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const dateCell = sheet.getRange(editRow, 2);
        dateCell.setValue(today);
        dateCell.setNumberFormat("MMMM d, yyyy");
      }
      sortHistorySheet(sheet);
    } finally {
      lock.releaseLock();
    }
    return;
  }

  if (sheetName !== "Movies") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();
  if (row <= 1) return;

  ensureHeaders(sheet);

  // If an AI assistant (like Muse.ai) or paste operation writes multiple columns/rows at once
  // starting at Column A, process any newly added unlinked rows and auto-sort!
  if (numRows > 1 || numCols > 1) {
    if (col === 1 && numRows <= 15) {
      processNewlyAddedRows(sheet);
    }
    return;
  }

  // MOVE TO "HISTORY" TAB WHEN PRIORITY TIER IS SET TO "Completed"
  if (col > 1 && e.value && isCompletedValue(e.value)) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return;
    try {
      moveRowToHistory(sheet, row);
    } finally {
      lock.releaseLock();
    }
    return;
  }

  // MOVE TO "WON'T WATCH" TAB WHEN PRIORITY TIER IS SET TO "Won't Watch"
  if (col > 1 && e.value && isWontWatchValue(e.value)) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return;
    try {
      moveRowToWontWatch(sheet, row);
    } finally {
      lock.releaseLock();
    }
    return;
  }

  // TRIGGER MODAL VIA COLUMN G ("Errors"):
  // If user types "fix" or "?" in Column G, open the visual movie picker modal for this row!
  if (col === 7 && e.value) {
    const cmd = String(e.value).trim().toLowerCase();
    if (cmd === "fix" || cmd === "?") {
      sheet.getRange(row, 7).setValue("");
      try {
        showMoviePickerModalForRow(sheet, row);
      } catch (err) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Click "🎬 Movie Tools -> 🎯 Fix Selected Row" in the top menu to open the movie picker for Row ${row}.`,
          "🎯 Open Movie Picker",
          6
        );
      }
      return;
    }
  }

  // COLUMN A EDIT (New or edited Movie Title)
  if (col === 1) {
    const rawTyped = String(e.value !== undefined ? e.value : e.range.getDisplayValue()).trim();
    if (!rawTyped) return;

    // TRIGGER MODAL VIA "?" SUFFIX IN COLUMN A (e.g. "Paul?"):
    // Strips the "?" and opens the visual movie picker modal immediately!
    if (rawTyped.endsWith("?")) {
      const cleanTyped = rawTyped.replace(/\?+\s*$/, "").trim();
      if (cleanTyped) {
        sheet.getRange(row, 1).setValue(cleanTyped);
        try {
          showMoviePickerModalForRow(sheet, row, cleanTyped);
          return;
        } catch (err) {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Click "🎬 Movie Tools -> 🎯 Fix Selected Row" in the top menu to choose the exact match for "${cleanTyped}".`,
            "🎯 Open Movie Picker",
            6
          );
        }
      }
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return;
    try {
      sheet.getRange(row, 1).clearNote();
      sheet.getRange(row, 7).setValue(""); // Clear previous error in Col G

      const result = updateMovieData(sheet, row, rawTyped, { preserveUserTitle: true });

      if (result.status === "NOT_FOUND") {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Could not find "${rawTyped}". Click "🎬 Movie Tools -> 🎯 Fix Selected Row" to search manually.`,
          "⚠️ Movie Lookup Failed",
          5
        );
      } else if (result.status === "FUZZY_MATCH") {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Linked "${rawTyped}" to "${result.matchedTitle}". Wrong film? Click "🎬 Movie Tools -> 🎯 Fix Selected Row".`,
          "✨ Matched Movie",
          5
        );
      }
      checkAndFormatSheet(sheet);
    } finally {
      lock.releaseLock();
    }
  }
}

// ============================================================================
// 1B. EXTERNAL / AI EDIT WATCHER (`onSheetChange` + `processNewlyAddedRows`)
//     - Google Sheets does NOT fire `onEdit` when third-party integrations or AI
//       tools (like Muse.ai) write via the Google Sheets API or insert rows.
//     - `onSheetChange` (installable `On change` trigger) and the 5-minute
//       background watcher call `processNewlyAddedRows()`, which detects any
//       newly added row missing its IMDb link & metadata, looks it up, and sorts!
// ============================================================================
function onSheetChange(e) {
  processNewlyAddedRows();
}

function processNewlyAddedRows(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  // Quick lock-free pre-check in memory (<50ms, 0 external API calls if nothing is pending)
  const range = sheet.getRange(1, 1, lastRow, 7);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();

  const pendingTitles = [];
  for (let i = 1; i < values.length; i++) {
    const rawTitle = extractTitle(values[i][0], formulas[i][0], richText[i][0]);
    if (!rawTitle) continue;

    const errText = String(values[i][6] || "").trim();
    if (errText.startsWith("❌") || errText.startsWith("⚠️")) continue;

    const hasImdbLink = Boolean(extractImdbId(formulas[i][0], richText[i][0]));
    const hasCategory = String(values[i][1] || "").trim() !== "";
    const hasDate = values[i][2] !== "" && values[i][2] !== null && values[i][2] !== undefined;

    // A newly added row from an AI tool / external API will have a title in Col A
    // but no IMDb hyperlink yet (and/or missing Category or Release Date)
    if (!hasImdbLink && (!hasCategory || !hasDate)) {
      pendingTitles.push({ expectedRow: i + 1, rawTitle: rawTitle });
    }
  }

  if (pendingTitles.length === 0) return 0;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 0;
  let processed = 0;
  try {
    ensureHeaders(sheet);
    for (let k = 0; k < pendingTitles.length; k++) {
      const item = pendingTitles[k];
      const liveRow = findCurrentRowForMovie(sheet, item.expectedRow, item.rawTitle);
      if (!liveRow) continue;

      // Double-check under lock that it still lacks an IMDb link
      const liveFormula = sheet.getRange(liveRow, 1).getFormula();
      const liveRich = sheet.getRange(liveRow, 1).getRichTextValue();
      if (extractImdbId(liveFormula, liveRich)) continue;

      sheet.getRange(liveRow, 1).clearNote();
      sheet.getRange(liveRow, 7).setValue("");
      updateMovieData(sheet, liveRow, item.rawTitle, { preserveUserTitle: true });
      SpreadsheetApp.flush();
      processed++;
    }

    if (processed > 0) {
      checkAndFormatSheet(sheet);
    }
  } finally {
    lock.releaseLock();
  }
  return processed;
}

// ============================================================================
// 2. VISUAL MODAL MOVIE PICKER (Menu, '?' Suffix, or 'fix' in Col G)
// ============================================================================
function fixSelectedRowDialog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || sheet.getName() !== "Movies") {
    SpreadsheetApp.getUi().alert("Please select a movie row on the 'Movies' sheet first.");
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("Please click on a movie row (Row 2 or below) first.");
    return;
  }
  showMoviePickerModalForRow(sheet, row);
}

// Menu: "➕ Add New Movie" — opens the same IMDb search modal in "add" mode
function addNewMovieDialog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Could not find a sheet named 'Movies'.");
    return;
  }
  showMoviePickerModalForRow(sheet, null, "", "add");
}

function showMoviePickerModalForRow(sheet, row, overrideQuery, mode) {
  const isAddMode = mode === "add";
  let rawTitle = overrideQuery || "";
  if (!isAddMode) {
    const titleCell = sheet.getRange(row, 1);
    rawTitle = overrideQuery || extractTitle(titleCell.getDisplayValue(), titleCell.getFormula(), titleCell.getRichTextValue());
    if (!rawTitle) {
      SpreadsheetApp.getUi().alert(`Row ${row} does not have a movie title in Column A.`);
      return;
    }
  }

  const { cleanTitle } = rawTitle ? sanitizeTitleAndYear(rawTitle) : { cleanTitle: "" };
  const candidates = cleanTitle ? fetchMovieCandidates(cleanTitle, 8) : [];
  const initialCandidatesJson = JSON.stringify(candidates);
  const safeRawTitleJson = JSON.stringify(rawTitle);
  const priorityOptions = isAddMode ? getPriorityOptions(sheet) : [];
  const priorityOptionsJson = JSON.stringify(priorityOptions);

  const headerText = isAddMode
    ? `Search IMDb and click <b>Add</b> to insert a new movie:`
    : `Matches for Row ${row} (<b>${escapeHtml(rawTitle)}</b>):`;
  const statusText = isAddMode
    ? `⏳ Adding movie (IMDb link, Priority, Release Date &amp; Popcornmeter)...`
    : `⏳ Updating Row ${row} (IMDb link, Release Date &amp; Popcornmeter)...`;

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;padding:4px;color:#202124;">
      <!-- Live Search Bar Inside Modal -->
      <div style="display:flex;gap:8px;margin-bottom:10px;">
        <input id="searchInput" type="text" value="${escapeHtml(cleanTitle)}" autofocus
               placeholder="Search movie title..."
               onkeydown="if(event.key==='Enter') runLiveSearch()"
               style="flex-grow:1;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;" />
        <button onclick="runLiveSearch()"
                style="background:#f1f3f4;color:#202124;border:1px solid #dadce0;padding:7px 14px;border-radius:6px;cursor:pointer;font-weight:600;font-size:13px;">
          🔍 Search
        </button>
      </div>

      <!-- Priority Selector for Add Mode -->
      <div style="display:${isAddMode ? "flex" : "none"};align-items:center;justify-content:space-between;margin-bottom:10px;background:#f8f9fa;padding:8px 10px;border-radius:6px;border:1px solid #dadce0;font-size:13px;">
        <label for="prioritySelect" style="font-weight:600;color:#202124;">Priority Tier:</label>
        <select id="prioritySelect" style="padding:5px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff;font-weight:500;">
          <option value="">(None / Blank)</option>
          ${priorityOptions.map(opt => `<option value="${escapeHtml(opt)}" ${opt.toLowerCase().includes("high") ? "selected" : ""}>${escapeHtml(opt)}</option>`).join("")}
        </select>
      </div>

      <!-- Option to keep user's typed title vs replace with official title (fix mode only) -->
      <div style="margin-bottom:8px;font-size:12px;color:#5f6368;display:flex;align-items:center;justify-content:space-between;">
        <span>${headerText}</span>
        <label style="display:${isAddMode ? "none" : "flex"};align-items:center;gap:4px;cursor:pointer;">
          <input type="checkbox" id="updateTitleCheckbox" checked />
          <span>Update Col A text to official title</span>
        </label>
      </div>

      <!-- Candidate List -->
      <div id="resultsContainer" style="max-height:260px;overflow-y:auto;border:1px solid #dadce0;border-radius:8px;background:#fff;"></div>

      <!-- Direct IMDb ID / URL Input -->
      <div style="margin-top:12px;padding-top:10px;border-top:1px solid #dadce0;display:flex;gap:8px;align-items:center;">
        <input id="customImdb" type="text" placeholder="Or paste any IMDb URL or ID (e.g. tt1092026)"
               onkeydown="if(event.key==='Enter') selectCustom()"
               style="flex-grow:1;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;" />
        <button onclick="selectCustom()"
                style="background:#1e8e3e;color:white;border:none;padding:7px 14px;border-radius:6px;cursor:pointer;font-weight:600;font-size:13px;">
          Use IMDb ID
        </button>
      </div>

      <div id="status" style="margin-top:8px;font-size:12px;color:#1a73e8;font-weight:500;display:none;">
        ${statusText}
      </div>
    </div>

    <script>
      const rowNum = ${row === null || row === undefined ? "null" : row};
      const isAddMode = ${isAddMode ? "true" : "false"};
      const originalTitle = ${safeRawTitleJson};
      let currentCandidates = ${initialCandidatesJson};

      function renderCandidates(list) {
        const container = document.getElementById('resultsContainer');
        if (!list || list.length === 0) {
          container.innerHTML = isAddMode && !document.getElementById('searchInput').value.trim()
            ? '<div style="padding:16px;text-align:center;color:#5f6368;font-size:13px;">Type a movie title above and press Enter to search IMDb.</div>'
            : '<div style="padding:16px;text-align:center;color:#5f6368;font-size:13px;">No matches found. Try a different search above or paste an IMDb ID below.</div>';
          return;
        }
        container.innerHTML = list.map((c, idx) => {
          const poster = c.poster || 'https://via.placeholder.com/44x64?text=No+Img';
          const yearStr = c.year ? ' (' + c.year + ')' : '';
          const typeBadge = c.type && c.type !== 'movie' ? ' [' + c.type + ']' : '';
          return '<div style="display:flex;align-items:center;justify-content:space-between;padding:8px 10px;border-bottom:1px solid #f1f3f4;gap:10px;">' +
            '<img src="' + poster + '" style="width:42px;height:62px;object-fit:cover;border-radius:4px;flex-shrink:0;background:#f1f3f4;" />' +
            '<div style="flex-grow:1;min-width:0;">' +
              '<div style="font-weight:600;font-size:13px;color:#202124;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' +
                escapeHtmlJs(c.title) + yearStr +
              '</div>' +
              '<div style="color:#5f6368;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' +
                escapeHtmlJs(c.stars || '') + escapeHtmlJs(typeBadge) +
              '</div>' +
              '<a href="https://www.imdb.com/title/' + c.imdbID + '/" target="_blank" style="font-size:11px;color:#1a73e8;text-decoration:none;">' +
                'IMDb: ' + c.imdbID + ' ↗' +
              '</a>' +
            '</div>' +
            '<button onclick="pickCandidate(' + idx + ')" ' +
                    'style="background:#1a73e8;color:white;border:none;padding:7px 14px;border-radius:6px;cursor:pointer;font-weight:600;font-size:12px;flex-shrink:0;">' +
              (isAddMode ? '➕ Add' : 'Select') +
            '</button>' +
          '</div>';
        }).join('');
      }

      function escapeHtmlJs(s) {
        return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
      }

      function runLiveSearch() {
        const q = document.getElementById('searchInput').value.trim();
        if (!q) return;
        document.getElementById('resultsContainer').innerHTML = '<div style="padding:16px;text-align:center;color:#5f6368;font-size:13px;">🔍 Searching IMDb...</div>';
        google.script.run
          .withSuccessHandler(list => {
            currentCandidates = list || [];
            renderCandidates(currentCandidates);
          })
          .withFailureHandler(err => alert(err.message))
          .fetchMovieCandidates(q, 8);
      }

      function getSelectedPriority() {
        const sel = document.getElementById('prioritySelect');
        return sel ? sel.value : '';
      }

      function pickCandidate(idx) {
        const c = currentCandidates[idx];
        if (!c) return;
        const replaceTitle = isAddMode || document.getElementById('updateTitleCheckbox').checked;
        const finalTitle = replaceTitle ? c.title : originalTitle;
        const prio = isAddMode ? getSelectedPriority() : '';
        applySelection(c.imdbID, finalTitle, c.year ? String(c.year) : '', prio);
      }

      function selectCustom() {
        const val = document.getElementById('customImdb').value.trim();
        const m = val.match(/(tt\\d{6,10})/i);
        if (!m) {
          alert('Please enter a valid IMDb ID or URL containing tt followed by digits (e.g. tt1092026).');
          return;
        }
        const replaceTitle = isAddMode || document.getElementById('updateTitleCheckbox').checked;
        const prio = isAddMode ? getSelectedPriority() : '';
        applySelection(m[1].toLowerCase(), replaceTitle ? '' : originalTitle, '', prio);
      }

      function applySelection(imdbId, titleToUse, yearToUse, priorityToUse) {
        document.getElementById('status').style.display = 'block';
        const runner = google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler(err => {
            alert(err.message);
            document.getElementById('status').style.display = 'none';
          });
        if (isAddMode) {
          runner.addNewMovieFromDialog(imdbId, titleToUse, yearToUse, priorityToUse);
        } else {
          runner.applyDialogMovieSelection(rowNum, imdbId, titleToUse, yearToUse);
        }
      }

      renderCandidates(currentCandidates);
      document.getElementById('searchInput').focus();
    </script>
  `).setWidth(520).setHeight(510);

  SpreadsheetApp.getUi().showModalDialog(html, isAddMode ? "➕ Add New Movie" : `🎯 Choose Correct Movie (Row ${row})`);
}

// Appends a new movie selected from the "➕ Add New Movie" modal, fills in its data, then sorts it into place
function addNewMovieFromDialog(imdbId, chosenTitle, chosenYear, chosenPriority) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) throw new Error("Could not find a sheet named 'Movies'.");

  const cleanImdbId = String(imdbId || "").trim().toLowerCase();
  if (!/^tt\d+$/.test(cleanImdbId)) throw new Error("Invalid IMDb ID.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ensureHeaders(sheet);

    // Prevent duplicates: check whether this IMDb ID is already on the Movies list
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const formulasA = sheet.getRange(2, 1, lastRow - 1, 1).getFormulas();
      const richA = sheet.getRange(2, 1, lastRow - 1, 1).getRichTextValues();
      for (let r = 0; r < formulasA.length; r++) {
        if (extractImdbId(formulasA[r][0], richA[r][0]) === cleanImdbId) {
          const existing = extractTitle(sheet.getRange(r + 2, 1).getDisplayValue(), formulasA[r][0], richA[r][0]);
          throw new Error(`"${existing}" is already on your Movies list (Row ${r + 2}).`);
        }
      }
    }

    // Resolve the official title & year if not provided (e.g. pasted IMDb ID)
    let resolvedTitle = chosenTitle;
    let resolvedYear = chosenYear;
    if (!resolvedTitle || !resolvedYear) {
      const idDetails = fetchImdbDetailsById(cleanImdbId);
      if (idDetails) {
        if (!resolvedTitle && idDetails.title) resolvedTitle = idDetails.title;
        if (!resolvedYear && idDetails.year) resolvedYear = String(idDetails.year);
      }
    }
    if (!resolvedTitle) throw new Error(`Could not find a title for IMDb ID ${cleanImdbId}.`);

    // Insert at the first empty row below the last movie title in Column A
    const newRow = findBottomMovieRow(sheet, 1) + 1;
    const safeTitle = resolvedTitle.replace(/"/g, '""');
    sheet.getRange(newRow, 1).setFormula(`=HYPERLINK("https://www.imdb.com/title/${cleanImdbId}/", "${safeTitle}")`);

    // Write priority if provided
    if (chosenPriority) {
      const priorityCol = findPriorityColumn(sheet);
      sheet.getRange(newRow, priorityCol).setValue(chosenPriority);
    }

    updateMovieData(sheet, newRow, resolvedTitle, {
      forcedImdbId: cleanImdbId,
      forcedTitle: resolvedTitle,
      forcedYear: resolvedYear,
      preserveUserTitle: true,
      manualOverride: true
    });
    checkAndFormatSheet(sheet);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Added "${resolvedTitle}"${resolvedYear ? ` (${resolvedYear})` : ""}${chosenPriority ? ` [${chosenPriority}]` : ""} and sorted it into place.`,
      "➕ Movie Added",
      5
    );
  } finally {
    lock.releaseLock();
  }
}

function applyDialogMovieSelection(row, imdbId, chosenTitle, chosenYear) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cleanImdbId = String(imdbId || "").trim().toLowerCase();
    const titleCell = sheet.getRange(row, 1);
    const rawTitle = extractTitle(titleCell.getDisplayValue(), titleCell.getFormula(), titleCell.getRichTextValue());

    // If user pasted a custom IMDb ID (or title/year weren't provided), look up exact title & year by IMDb ID
    let resolvedTitle = chosenTitle;
    let resolvedYear = chosenYear;
    if (!resolvedTitle || !resolvedYear) {
      const idDetails = fetchImdbDetailsById(cleanImdbId);
      if (idDetails) {
        if (!resolvedTitle && idDetails.title) resolvedTitle = idDetails.title;
        if (!resolvedYear && idDetails.year) resolvedYear = String(idDetails.year);
      }
    }

    const displayTitle = resolvedTitle || rawTitle;
    const safeTitle = displayTitle.replace(/"/g, '""');

    // Clear any stale Rich Text link and write the updated =HYPERLINK formula for the title
    titleCell.clearContent();
    titleCell.setFormula(`=HYPERLINK("https://www.imdb.com/title/${cleanImdbId}/", "${safeTitle}")`);
    sheet.getRange(row, 7).setValue("");

    updateMovieData(sheet, row, displayTitle, {
      forcedImdbId: cleanImdbId,
      forcedTitle: displayTitle,
      forcedYear: resolvedYear,
      preserveUserTitle: true,
      manualOverride: true
    });
    checkAndFormatSheet(sheet);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 3. MENU TOOL: UPDATE MISSING DATES, SCORES & LINKS
// ============================================================================
function updateMissingDataOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return;

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  sheet.getRange(2, 1, lastRow - 1, 1).clearNote();

  const range = sheet.getRange(1, 1, lastRow, 7);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();
  const notesCol = findNotesColumn(sheet);

  let updates = 0;
  let notFoundCount = 0;
  const startTime = new Date().getTime();

  for (let i = 1; i < values.length; i++) {
    if (new Date().getTime() - startTime > 270000) {
      checkAndFormatSheet(sheet);
      SpreadsheetApp.getUi().alert(
        `Paused at row ${i + 1} to prevent script timeout.\nUpdated: ${updates} | Unfound: ${notFoundCount}\nRun "Update Missing Dates, Scores & Links" again to continue!`
      );
      return;
    }

    const rawTitle = extractTitle(values[i][0], formulas[i][0], richText[i][0]);
    if (!rawTitle) continue;

    const category = String(values[i][1]).toLowerCase().trim();
    const rawReleaseDate = values[i][2];
    const recDate = sanitizeAndRecoverDate(rawReleaseDate);
    const isCorrupt1905 = (rawReleaseDate instanceof Date && rawReleaseDate.getFullYear() === 1905);
    const rtScore = values[i][3];
    const rtFormula = String(formulas[i][3] || "");
    const errorCol = String(values[i][6] || "");
    const existingImdbId = extractImdbId(formulas[i][0], richText[i][0]);

    if (errorCol.startsWith("❌")) continue;

    const hasValidDate = !!(recDate && recDate.date && !isCorrupt1905);
    const hasRtLink = rtFormula.toUpperCase().startsWith("=HYPERLINK") && /rottentomatoes\.com/i.test(rtFormula);

    const needsLink = !existingImdbId;
    const needsDate = !hasValidDate || isCorrupt1905;
    const needsScore = category !== "to be released" && (rtScore === "" || rtScore === "N/A" || !hasRtLink);
    const hadError = errorCol.startsWith("⚠️");

    // FAST PATH: Row already has IMDb link & Release Date, only needs RT Popcornmeter + RT Hyperlink
    if (!needsLink && !needsDate && !hadError && needsScore) {
      const { cleanTitle, year } = sanitizeTitleAndYear(rawTitle);
      const searchYear = year || (recDate ? recDate.year : null);
      const rtData = fetchRottenTomatoesData(cleanTitle, searchYear);
      if (rtData) {
        const liveRow = findCurrentRowForMovie(sheet, i + 1, rawTitle);
        if (liveRow) {
          if (rtData.popcornmeter) {
            writeRtScoreCell(sheet.getRange(liveRow, 4), rtData.popcornmeter, rtData.url);
            updates++;
          }
          if (rtData.isNetflix) {
            const curNote = String(sheet.getRange(liveRow, notesCol).getValue() || "").trim();
            sheet.getRange(liveRow, notesCol).setValue(upsertNetflixNote(curNote));
          }
          SpreadsheetApp.flush();
        }
      }
      continue;
    }

    if (needsLink || needsDate || needsScore || hadError) {
      const liveRow = findCurrentRowForMovie(sheet, i + 1, rawTitle);
      if (!liveRow) continue;
      const res = updateMovieData(sheet, liveRow, rawTitle, { preserveUserTitle: true });
      SpreadsheetApp.flush();
      if (res.status === "SUCCESS" || res.status === "FUZZY_MATCH") updates++;
      if (res.status === "NOT_FOUND") notFoundCount++;
    }
  }

  checkAndFormatSheet(sheet);
  SpreadsheetApp.getUi().alert(
    `Finished checking sheet!\nUpdated: ${updates}\nUnfound (logged in Col G): ${notFoundCount}` +
    (omdbQuotaExceeded ? `\n\n⚠️ Note: Your free OMDb API key reached its 1,000/day limit, so Rotten Tomatoes & IMDb fallback was used.` : "")
  );
}

// ============================================================================
// 4. FAST MENU TOOL: STRICTLY LINKS ANY UNLINKED MOVIES
//    (Never overwrites titles, dates, categories, or scores, and uses 0 OMDb calls!)
// ============================================================================
function linkUnlinkedMoviesOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return;

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  sheet.getRange(2, 1, lastRow - 1, 1).clearNote();

  const values = sheet.getRange(1, 1, lastRow, 7).getValues();
  const formulas = sheet.getRange(1, 1, lastRow, 7).getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();
  const notesCol = findNotesColumn(sheet);

  let linkedCount = 0;
  let rtLinkedCount = 0;
  let convertedRichLinks = 0;
  let notFoundCount = 0;
  const startTime = new Date().getTime();
  const today = new Date();

  for (let i = 1; i < values.length; i++) {
    if (new Date().getTime() - startTime > 270000) {
      checkAndFormatSheet(sheet);
      SpreadsheetApp.getUi().alert(`Paused at row ${i + 1} to prevent timeout.\nIMDb Linked: ${linkedCount} | RT Linked: ${rtLinkedCount}\nRun again to finish!`);
      return;
    }

    const rawTitle = extractTitle(values[i][0], formulas[i][0], richText[i][0]);
    if (!rawTitle) continue;

    const errorCol = String(values[i][6] || "");
    if (errorCol.startsWith("❌")) continue;

    const { cleanTitle, year } = sanitizeTitleAndYear(rawTitle);
    const category = String(values[i][1]).toLowerCase().trim();
    const recDate = sanitizeAndRecoverDate(values[i][2], year);
    const isToBeReleased = category === "to be released";

    let searchYear = year || (recDate ? recDate.year : null);

    // 1. Check / Add IMDb Hyperlink in Column A
    const hasHyperlinkFormula = formulas[i][0] && formulas[i][0].toUpperCase().startsWith("=HYPERLINK");
    if (!hasHyperlinkFormula) {
      const existingImdbId = extractImdbId(formulas[i][0], richText[i][0]);
      if (existingImdbId) {
        const liveRow = findCurrentRowForMovie(sheet, i + 1, rawTitle);
        if (liveRow) {
          const safeTitle = rawTitle.replace(/"/g, '""');
          sheet.getRange(liveRow, 1).setFormula(`=HYPERLINK("https://www.imdb.com/title/${existingImdbId}/", "${safeTitle}")`);
          convertedRichLinks++;
        }
      } else {
        const match = findImdbIdWithTypoTolerance(
          cleanTitle,
          searchYear,
          isToBeReleased ? today.getFullYear() : null
        );

        const liveRow = findCurrentRowForMovie(sheet, i + 1, rawTitle);
        if (liveRow) {
          if (match && match.imdbID) {
            // IMPORTANT: Always keep `rawTitle` (never overwrite "Paul" with another movie's title!)
            const safeTitle = rawTitle.replace(/"/g, '""');
            sheet.getRange(liveRow, 1).setFormula(`=HYPERLINK("https://www.imdb.com/title/${match.imdbID}/", "${safeTitle}")`);
            sheet.getRange(liveRow, 7).setValue("");
            SpreadsheetApp.flush();
            linkedCount++;
          } else {
            sheet.getRange(liveRow, 7).setValue(`❌ Movie not found for "${cleanTitle}"`);
            SpreadsheetApp.flush();
            notFoundCount++;
          }
        }
      }
    }

    // 2. Check / Add Rotten Tomatoes Popcornmeter Hyperlink in Column D (for released movies)
    const rtFormula = String(formulas[i][3] || "");
    const hasRtLink = rtFormula.toUpperCase().startsWith("=HYPERLINK") && /rottentomatoes\.com/i.test(rtFormula);
    if (!isToBeReleased && !hasRtLink) {
      const rtData = fetchRottenTomatoesData(cleanTitle, searchYear);
      if (rtData) {
        const liveRow = findCurrentRowForMovie(sheet, i + 1, rawTitle);
        if (liveRow) {
          if (rtData.popcornmeter || values[i][3]) {
            const scoreText = rtData.popcornmeter || String(values[i][3]).trim();
            writeRtScoreCell(sheet.getRange(liveRow, 4), scoreText, rtData.url);
            rtLinkedCount++;
          }
          if (rtData.isNetflix) {
            const curNote = String(sheet.getRange(liveRow, notesCol).getValue() || "").trim();
            sheet.getRange(liveRow, notesCol).setValue(upsertNetflixNote(curNote));
          }
          SpreadsheetApp.flush();
        }
      }
    }
  }

  checkAndFormatSheet(sheet);
  SpreadsheetApp.getUi().alert(
    `Finished linking!\nNewly Linked IMDb: ${linkedCount}\nNewly Linked RT Scores: ${rtLinkedCount}\nPreserved Existing Rich Links: ${convertedRichLinks}\nUnfound (in Col G): ${notFoundCount}`
  );
}

// ============================================================================
// 4B. MENU TOOL: SCAN, ADD & REPAIR RT SCORES / LINKS ONLY (Column D)
//     - Automatically detects & repairs any shifted/wrong RT links from nearby rows
//     - Verifies the live movie title on the row right before every write (row-shift proof)
//     - Uses parallel mini-batches of 5 movies + live flush so it runs ~5x faster
// ============================================================================
function updateMissingRtScoresAndLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return;

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const range = sheet.getRange(1, 1, lastRow, 7);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const formulas = range.getFormulas();
  const richTextA = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();

  // Pre-extract clean titles and existing RT URLs so we can detect duplicate or shifted RT links
  const rowMeta = [];
  const rtUrlCounts = {};
  for (let i = 1; i < values.length; i++) {
    const rawTitle = extractTitle(values[i][0], formulas[i][0], richTextA[i][0]);
    const { cleanTitle, year } = sanitizeTitleAndYear(rawTitle);
    const rtFormula = String(formulas[i][3] || "");
    const existingRtUrl = extractRtUrlFromFormula(rtFormula);
    if (existingRtUrl) {
      const normUrl = existingRtUrl.toLowerCase();
      rtUrlCounts[normUrl] = (rtUrlCounts[normUrl] || 0) + 1;
    }
    rowMeta[i] = {
      row: i + 1,
      rawTitle: rawTitle,
      cleanTitle: cleanTitle,
      year: year,
      existingRtUrl: existingRtUrl
    };
  }

  // Build work queue of rows that need a score, need a link, or have a shifted/wrong RT link
  const workQueue = [];
  for (let i = 1; i < values.length; i++) {
    const meta = rowMeta[i];
    if (!meta || !meta.rawTitle) continue;

    const category = String(values[i][1]).toLowerCase().trim();
    if (category === "to be released") continue;

    const rtDisplay = String(displayValues[i][3] || "").trim();
    const hasScore = rtDisplay !== "" && rtDisplay !== "N/A";
    const hasRtLink = Boolean(meta.existingRtUrl);

    // Check if an existing RT link on this row actually belongs to another movie (e.g. shifted from rows above)
    let hasMismatchedLink = false;
    if (hasRtLink) {
      const normUrl = meta.existingRtUrl.toLowerCase();
      const matchesOwnTitle = doesRtSlugMatchTitle(meta.existingRtUrl, meta.cleanTitle);
      const isDuplicateUrl = rtUrlCounts[normUrl] > 1;

      // Also check if the slug matches a neighbor row (1–10 rows above or below) that has a different title
      let matchesNeighborAboveOrBelow = false;
      for (let offset = -10; offset <= 10; offset++) {
        if (offset === 0) continue;
        const neighbor = rowMeta[i + offset];
        if (!neighbor || !neighbor.cleanTitle) continue;
        if (normalizeForCompare(neighbor.cleanTitle) === normalizeForCompare(meta.cleanTitle)) continue;
        if (doesRtSlugMatchTitle(meta.existingRtUrl, neighbor.cleanTitle) && !matchesOwnTitle) {
          matchesNeighborAboveOrBelow = true;
          break;
        }
      }

      if (!matchesOwnTitle || isDuplicateUrl || matchesNeighborAboveOrBelow) {
        hasMismatchedLink = true;
      }
    }

    // Skip rows that already have a valid, matching RT hyperlink and score
    if (hasScore && hasRtLink && !hasMismatchedLink) continue;

    const recDate = sanitizeAndRecoverDate(values[i][2], meta.year);
    let searchYear = meta.year || (recDate ? recDate.year : null);

    workQueue.push({
      row: meta.row,
      rawTitle: meta.rawTitle,
      cleanTitle: meta.cleanTitle,
      searchYear: searchYear,
      rtDisplay: rtDisplay,
      hasScore: hasScore,
      hasMismatchedLink: hasMismatchedLink
    });
  }

  const notesCol = findNotesColumn(sheet);
  let updatedScores = 0;
  let repairedShifted = 0;
  let clearedShiftedNoScore = 0;
  let linkedExisting = 0;
  let skippedNoRtScore = 0;
  const startTime = new Date().getTime();
  const BATCH_SIZE = 5;

  for (let b = 0; b < workQueue.length; b += BATCH_SIZE) {
    if (new Date().getTime() - startTime > 270000) {
      const nextRow = workQueue[b] ? workQueue[b].row : "";
      SpreadsheetApp.getUi().alert(
        `Paused at row ${nextRow} to prevent script timeout.\n` +
        `New Scores: ${updatedScores} | Fixed Wrong/Shifted Links: ${repairedShifted} | Cleared Unscored Shifts: ${clearedShiftedNoScore}\n` +
        `Run "🍿 Update Missing RT Scores & Links" again to continue!`
      );
      return;
    }

    const batch = workQueue.slice(b, b + BATCH_SIZE);
    const rtResults = fetchRottenTomatoesDataBatch(batch);

    // Read current live Column A titles right before writing this mini-batch so row shifts can NEVER happen
    const currentLastRow = sheet.getLastRow();
    const liveColA = currentLastRow > 1
      ? sheet.getRange(2, 1, currentLastRow - 1, 1).getDisplayValues()
      : [];

    for (let k = 0; k < batch.length; k++) {
      const item = batch[k];
      const rtData = rtResults[k];
      const liveRow = resolveRowFromLiveColA(liveColA, item.row, item.rawTitle);
      if (!liveRow) continue; // Movie was deleted or moved to "Won't Watch"

      if (rtData && rtData.isNetflix) {
        const curNote = String(sheet.getRange(liveRow, notesCol).getValue() || "").trim();
        sheet.getRange(liveRow, notesCol).setValue(upsertNetflixNote(curNote));
      }

      if (rtData && (rtData.popcornmeter || (item.hasScore && !item.hasMismatchedLink))) {
        const scoreToWrite = rtData.popcornmeter || item.rtDisplay;
        writeRtScoreCell(sheet.getRange(liveRow, 4), scoreToWrite, rtData.url);
        if (item.hasMismatchedLink) {
          repairedShifted++;
        } else if (!item.hasScore && rtData.popcornmeter) {
          updatedScores++;
        } else {
          linkedExisting++;
        }
      } else if (item.hasMismatchedLink) {
        // This row had a wrong/shifted RT score from another movie, and the actual movie has no RT score yet
        sheet.getRange(liveRow, 4).setValue("");
        clearedShiftedNoScore++;
      } else {
        skippedNoRtScore++;
      }
    }

    SpreadsheetApp.flush(); // Commit & display this mini-batch live on the sheet immediately
  }

  SpreadsheetApp.getUi().alert(
    `Finished Rotten Tomatoes scan & repair (Column D)!\n` +
    `• New RT Scores + Links Added: ${updatedScores}\n` +
    `• Fixed Wrong/Shifted RT Scores: ${repairedShifted}\n` +
    `• Cleared Shifted Scores (No RT Score Yet): ${clearedShiftedNoScore}\n` +
    `• Existing Scores Linked to RT: ${linkedExisting}\n` +
    `• No RT Audience Score Available Yet: ${skippedNoRtScore}`
  );
}

// ============================================================================
// 5. DAILY VOD RELEASE DATE & LIFECYCLE AUTOMATION
//    - Promotes "To be released" -> "New Releases" (if Netflix) or "In Theaters" on release date
//    - Writes "Netflix" in Notes for Netflix exclusives (upcoming or already released)
//    - Looks up VOD release dates for "In Theaters" movies & writes "VOD: <Date>" in Notes
//    - Promotes "In Theaters" -> "New Releases" when VOD date arrives (or after 60 days)
//    - Graduates "New Releases" -> Genre after 6 months
// ============================================================================
function dailyVodAndActiveMoviesCheck() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return;

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  // First, link & populate any newly added unlinked rows (e.g. from AI/external tools),
  // then run a quick pass of checkAndFormatSheet so any "To be released" movies
  // whose release date arrived today become "In Theaters" (or "New Releases" if Netflix),
  // and any "In Theaters" movies whose existing "VOD: <Date>" in Notes arrived today become "New Releases".
  processNewlyAddedRows(sheet);
  checkAndFormatSheet(sheet);

  const numCols = Math.max(7, sheet.getLastColumn());
  const notesCol = findNotesColumn(sheet); // Dynamically finds "Notes" header (defaults to Col F / 6)

  const values = sheet.getRange(1, 1, lastRow, numCols).getValues();
  const formulas = sheet.getRange(1, 1, lastRow, numCols).getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sixMonthsAgo = getSixMonthsAgo(today);
  const startTime = new Date().getTime();
  const cache = CacheService.getScriptCache();

  // Build bulk VOD Release Date calendar map from DVDsReleaseDates (current + upcoming month)
  const vodCalendar = fetchDigitalReleaseCalendarMap();
  let vodUpdates = 0;

  const rtUrlBackfillQueue = [];

  for (let i = 1; i < values.length; i++) {
    if (new Date().getTime() - startTime > 240000) break;

    const rawTitle = extractTitle(values[i][0], formulas[i][0], richText[i][0]);
    if (!rawTitle) continue;

    const category = String(values[i][1]).toLowerCase().trim();
    const { cleanTitle, year } = sanitizeTitleAndYear(rawTitle);
    const recDate = sanitizeAndRecoverDate(values[i][2], year);
    const isCorrupt1905 = (values[i][2] instanceof Date && values[i][2].getFullYear() === 1905);
    const existingDate = recDate ? recDate.date : null;
    const hasValidDate = !!(existingDate && !isCorrupt1905);
    const existingNoteText = String(values[i][notesCol - 1] || "").trim();

    // 1. If a movie in "New Releases" (or "In Theaters") is now older than 6 months,
    //    run updateMovieData so it graduates to its permanent Genre!
    if ((category === "new releases" || category === "in theaters") && hasValidDate && existingDate <= sixMonthsAgo) {
      updateMovieData(sheet, i + 1, rawTitle, { preserveUserTitle: true });
      Utilities.sleep(120);
      continue;
    }

    const searchYear = year || (recDate ? recDate.year : null) || (hasValidDate ? existingDate.getFullYear() : today.getFullYear());

    // 2. For "To be released" movies:
    //    - If release date is unannounced (e.g. 12/31 placeholder or year-only), ensure Notes says "Release Date Unannounced"
    //    - Check Rotten Tomatoes (and OMDb if unannounced) for Netflix status & newly announced exact release dates
    if (category === "to be released") {
      const isCurrentlyUnannounced = !hasValidDate || (recDate && recDate.isYearOnly) || isDec31Date(existingDate) || isUnannouncedDateNote(existingNoteText);
      let workingNoteText = existingNoteText;

      if (isCurrentlyUnannounced && !isUnannouncedDateNote(workingNoteText)) {
        workingNoteText = upsertUnannouncedDateNote(workingNoteText);
        sheet.getRange(i + 1, notesCol).setValue(workingNoteText);
        vodUpdates++;
      }

      // ALWAYS re-check every "To be released" movie daily (no cache skip), so release date changes
      // (delays to a later date/year, or moving earlier) are picked up automatically.
      const imdbId = extractImdbId(formulas[i][0], richText[i][0]);
      const rtData = fetchRottenTomatoesData(cleanTitle, searchYear);

      if (rtData && rtData.isNetflix) {
        const withNetflix = upsertNetflixNote(workingNoteText);
        if (withNetflix !== workingNoteText) {
          workingNoteText = withNetflix;
          sheet.getRange(i + 1, notesCol).setValue(workingNoteText);
          vodUpdates++;
        }
      }

      // Candidate exact dates from Rotten Tomatoes (US theatrical / streaming) and OMDb/IMDb (exact IMDb ID match)
      let rtDateObj = null;
      if (rtData) {
        const rtDateStr = (rtData.isNetflix && rtData.vodDateStr)
          ? rtData.vodDateStr
          : (rtData.theatersDateStr || rtData.vodDateStr);
        const parsedRt = rtDateStr ? parseSafeLocalDate(rtDateStr) : null;
        // Ignore older namesake films (e.g. an older movie with the same title)
        if (parsedRt && !isNaN(parsedRt.getTime()) && !isDec31Date(parsedRt) && parsedRt > sixMonthsAgo) {
          rtDateObj = parsedRt;
        }
      }

      let omdbDateObj = null;
      let omdbYear = null;
      if (imdbId && !omdbQuotaExceeded) {
        const omdbData = fetchOmdbJson(`https://www.omdbapi.com/?i=${imdbId}&apikey=${apiKey}`);
        if (omdbData && omdbData.Response === "True") {
          if (omdbData.Released && omdbData.Released !== "N/A") {
            const parsedOmdb = parseSafeLocalDate(omdbData.Released);
            if (parsedOmdb && !isNaN(parsedOmdb.getTime()) && !isDec31Date(parsedOmdb)) {
              omdbDateObj = parsedOmdb;
            }
          }
          const ym = String(omdbData.Year || "").match(/\d{4}/);
          if (ym) omdbYear = parseInt(ym[0], 10);
        }
      }

      // Prefer RT's US release date; fall back to OMDb's exact IMDb-ID release date
      const latestExactDate = rtDateObj || omdbDateObj;
      const dateCell = sheet.getRange(i + 1, 3);

      if (latestExactDate) {
        const changed = !hasValidDate || isCurrentlyUnannounced ||
          !existingDate || existingDate.getTime() !== latestExactDate.getTime();
        if (changed) {
          dateCell.setValue(latestExactDate);
          dateCell.setNumberFormat("MMMM d, yyyy");
          vodUpdates++;
        }
        const cleanedNote = removeUnannouncedDateNote(workingNoteText);
        if (cleanedNote !== workingNoteText) {
          workingNoteText = cleanedNote;
          sheet.getRange(i + 1, notesCol).setValue(workingNoteText);
        }
      } else if (omdbYear && omdbYear >= today.getFullYear() && (!existingDate || existingDate.getFullYear() !== omdbYear)) {
        // No exact date yet, but the release YEAR changed (e.g. delayed from 2026 to 2027):
        // switch to a year-only 12/31 placeholder and mark "Release Date Unannounced"
        dateCell.setValue(new Date(omdbYear, 11, 31));
        dateCell.setNumberFormat("yyyy");
        const withUnannounced = upsertUnannouncedDateNote(workingNoteText);
        if (withUnannounced !== workingNoteText) {
          workingNoteText = withUnannounced;
          sheet.getRange(i + 1, notesCol).setValue(workingNoteText);
        }
        vodUpdates++;
      }

      Utilities.sleep(100);
      continue;
    }

    // 3. For movies currently "In Theaters", check daily for Netflix status, VOD release date & Popcornmeter
    if (category === "in theaters") {
      const imdbId = extractImdbId(formulas[i][0], richText[i][0]);
      let vodDateStr = (imdbId && vodCalendar.byImdb[imdbId]) || vodCalendar.byTitle[normalizeForCompare(cleanTitle)] || null;

      const rtData = fetchRottenTomatoesData(cleanTitle, searchYear);
      if (rtData && rtData.popcornmeter) {
        writeRtScoreCell(sheet.getRange(i + 1, 4), rtData.popcornmeter, rtData.url);
      }

      if ((rtData && rtData.isNetflix) || isNetflixNote(existingNoteText)) {
        const updatedNoteText = upsertNetflixNote(existingNoteText);
        if (updatedNoteText !== existingNoteText) {
          sheet.getRange(i + 1, notesCol).setValue(updatedNoteText);
          vodUpdates++;
        }
        // Netflix exclusives go directly to streaming -> move to "New Releases"
        if (!hasValidDate || existingDate <= today) {
          sheet.getRange(i + 1, 2).setValue("New Releases");
        }
      } else {
        if (!vodDateStr && rtData && rtData.vodDateStr) {
          vodDateStr = rtData.vodDateStr;
        }
        if (vodDateStr) {
          const updatedNoteText = upsertVodNote(existingNoteText, vodDateStr);
          if (updatedNoteText !== existingNoteText) {
            sheet.getRange(i + 1, notesCol).setValue(updatedNoteText);
            vodUpdates++;
          }
          // If the newly discovered VOD date is today or in the past, promote immediately to "New Releases"!
          const vodDateObj = parseSafeLocalDate(vodDateStr);
          if (vodDateObj && vodDateObj <= today) {
            sheet.getRange(i + 1, 2).setValue("New Releases");
          }
        }
      }
      Utilities.sleep(120);
      continue;
    }

    // 4. For already-released movies (New Releases / Genres):
    //    - If RT score (Column D) is blank even though the movie is released, look up & populate RT score + link
    //    - Queue existing RT links that aren't marked "Netflix" yet for fast parallel Netflix check
    const rtScoreVal = String(values[i][3] || "").trim();
    const rtUrl = extractRtUrlFromFormula(formulas[i][3]);
    if (!rtScoreVal || rtScoreVal === "N/A") {
      const rtData = fetchRottenTomatoesData(cleanTitle, searchYear);
      if (rtData) {
        if (rtData.popcornmeter) {
          writeRtScoreCell(sheet.getRange(i + 1, 4), rtData.popcornmeter, rtData.url);
          vodUpdates++;
        }
        if (rtData.isNetflix && !isNetflixNote(existingNoteText)) {
          const updatedNoteText = upsertNetflixNote(existingNoteText);
          if (updatedNoteText !== existingNoteText) {
            sheet.getRange(i + 1, notesCol).setValue(updatedNoteText);
            vodUpdates++;
          }
        }
      }
      Utilities.sleep(100);
    } else if (!isNetflixNote(existingNoteText) && rtUrl) {
      const urlCacheKey = `nf_url_${rtUrl.replace(/[^a-zA-Z0-9_-]/g, "").slice(-60)}`;
      if (!cache || !cache.get(urlCacheKey)) {
        rtUrlBackfillQueue.push({
          row: i + 1,
          rtUrl: rtUrl,
          existingNoteText: existingNoteText,
          cacheKey: urlCacheKey
        });
      }
    }
  }

  // Execute fast parallel check on existing released movies with RT links to tag any Netflix exclusives in Column F
  const NF_BATCH_SIZE = 8;
  for (let b = 0; b < rtUrlBackfillQueue.length; b += NF_BATCH_SIZE) {
    if (new Date().getTime() - startTime > 250000) break;
    const batch = rtUrlBackfillQueue.slice(b, b + NF_BATCH_SIZE);
    try {
      const reqs = batch.map(item => ({
        url: item.rtUrl,
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
      }));
      const resps = UrlFetchApp.fetchAll(reqs);
      for (let j = 0; j < batch.length; j++) {
        const res = resps[j];
        if (!res || res.getResponseCode() !== 200) continue;
        const parsed = parseRtMoviePageHtml(res.getContentText(), { url: batch[j].rtUrl });
        if (cache) {
          try { cache.put(batch[j].cacheKey, "1", 21600); } catch (e) {}
        }
        if (parsed && parsed.isNetflix) {
          const curNote = String(sheet.getRange(batch[j].row, notesCol).getValue() || "").trim();
          const updatedNote = upsertNetflixNote(curNote);
          if (updatedNote !== curNote) {
            sheet.getRange(batch[j].row, notesCol).setValue(updatedNote);
            vodUpdates++;
          }
        }
      }
    } catch (e) {}
  }

  // Final sort & color pass (promotes any newly-reached VOD dates to "New Releases")
  checkAndFormatSheet(sheet);
  return vodUpdates;
}

function dailyVodCheckMenuWrapper() {
  const count = dailyVodAndActiveMoviesCheck();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Checked movies for VOD dates, Netflix exclusives & category transitions. Updated ${count || 0} note(s).`,
    "📺 VOD & Category Check Complete",
    5
  );
}

// Weekly / Manual Active Movies Refresh ("To be released", "In Theaters" & "New Releases")
function weeklyActiveMoviesRefresh() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return;

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const values = sheet.getRange(1, 1, lastRow, 7).getValues();
  const formulas = sheet.getRange(1, 1, lastRow, 1).getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();

  for (let i = 1; i < values.length; i++) {
    const rawTitle = extractTitle(values[i][0], formulas[i][0], richText[i][0]);
    const category = String(values[i][1]).toLowerCase().trim();

    if (rawTitle && (category === "to be released" || category === "in theaters" || category === "new releases")) {
      updateMovieData(sheet, i + 1, rawTitle, { preserveUserTitle: true });
      Utilities.sleep(150);
    }
  }

  dailyVodAndActiveMoviesCheck();
}

// ============================================================================
// 6. CORE MOVIE LOOKUP FUNCTION (Popcornmeter + VOD + Safe Title Handling)
// ============================================================================
function updateMovieData(sheet, rowNum, titleText, options) {
  options = options || {};
  const titleCell = sheet.getRange(rowNum, 1);
  const errorCell = sheet.getRange(rowNum, 7);
  titleCell.clearNote();

  try {
    const existingFormula = titleCell.getFormula();
    const existingRichText = titleCell.getRichTextValue();
    const rawDisplayTitle = extractTitle(titleText, existingFormula, existingRichText);
    const { cleanTitle, year } = sanitizeTitleAndYear(rawDisplayTitle);

    const existingImdbId = options.forcedImdbId || extractImdbId(existingFormula, existingRichText);

    const currentCategory = sheet.getRange(rowNum, 2).getValue();
    const existingDateRaw = sheet.getRange(rowNum, 3).getValue();
    const existingRecovered = sanitizeAndRecoverDate(existingDateRaw, options.forcedYear || year);
    const existingDate = existingRecovered ? existingRecovered.date : null;
    const existingIsCorrupt1905 = (existingDateRaw instanceof Date && existingDateRaw.getFullYear() === 1905);
    const notesCol = findNotesColumn(sheet);
    const existingNoteText = String(sheet.getRange(rowNum, notesCol).getValue() || "").trim();

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const sixMonthsAgo = getSixMonthsAgo(today);
    const sixtyDaysAgo = new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000);

    const isToBeReleased = String(currentCategory).toLowerCase().trim() === "to be released";
    let searchYear = options.forcedYear || year;
    if (!searchYear && existingRecovered && existingRecovered.year && existingRecovered.year >= 1920) {
      searchYear = existingRecovered.year;
    }

    let data = null;
    let fuzzyMatchedTitle = null;
    let matchedImdbId = existingImdbId || null;

    // 1. Query OMDb (unless quota already exceeded during this run)
    if (!omdbQuotaExceeded) {
      // FAST PATH: Exact IMDb ID lookup (?i=)
      if (matchedImdbId) {
        const idParsed = fetchOmdbJson(`https://www.omdbapi.com/?i=${matchedImdbId}&plot=full&apikey=${apiKey}`);
        if (idParsed && idParsed.Response === "True") data = idParsed;
      }

      // ATTEMPT 1: Exact OMDb Title Lookup (?t=)
      if (!data && !omdbQuotaExceeded) {
        let url = `https://www.omdbapi.com/?t=${encodeURIComponent(cleanTitle)}&type=movie&plot=full&apikey=${apiKey}`;
        if (searchYear) url += `&y=${searchYear}`;
        const parsed = fetchOmdbJson(url);
        if (parsed && parsed.Response === "True") data = parsed;
      }

      // Retry without year constraint ONLY if user didn't explicitly specify a year
      if (!data && searchYear && !year && !options.forcedYear && !omdbQuotaExceeded) {
        const url = `https://www.omdbapi.com/?t=${encodeURIComponent(cleanTitle)}&type=movie&plot=full&apikey=${apiKey}`;
        const parsed = fetchOmdbJson(url);
        if (parsed && parsed.Response === "True") data = parsed;
      }
    }

    // 2. IMDb Native Titles Suggestion API (Free, unlimited, typo-tolerant)
    let suggestion = null;
    if (!matchedImdbId) {
      suggestion = findImdbIdWithTypoTolerance(
        cleanTitle,
        searchYear,
        isToBeReleased ? today.getFullYear() : null
      );
      if (suggestion && suggestion.imdbID) {
        matchedImdbId = suggestion.imdbID;
        fuzzyMatchedTitle = suggestion.matchedTitle;
        if (!searchYear && suggestion.year) searchYear = suggestion.year;

        if (!data && !omdbQuotaExceeded) {
          const idParsed = fetchOmdbJson(`https://www.omdbapi.com/?i=${matchedImdbId}&plot=full&apikey=${apiKey}`);
          if (idParsed && idParsed.Response === "True") {
            data = idParsed;
            fuzzyMatchedTitle = idParsed.Title || suggestion.matchedTitle;
          }
        }
      }
    } else if (!data) {
      // We already have an exact IMDb ID (either existing or from Row Fixer),
      // so look up its exact title & year by IMDb ID without guessing by title!
      const idDetails = fetchImdbDetailsById(matchedImdbId);
      if (idDetails) {
        fuzzyMatchedTitle = idDetails.title || options.forcedTitle || null;
        if (!searchYear && idDetails.year) searchYear = idDetails.year;
      }
    }

    if (data && data.imdbID) {
      matchedImdbId = data.imdbID;
      if (!searchYear && data.Year) {
        const ym = String(data.Year).match(/\d{4}/);
        if (ym) searchYear = parseInt(ym[0], 10);
      }
    }

    // 3. Query Rotten Tomatoes for Popcornmeter (Audience Score), Streaming/VOD Date, & Fallback Metadata
    const rawLookupTitleForRt = options.forcedTitle || (data && data.Title) || fuzzyMatchedTitle || cleanTitle;
    const { cleanTitle: lookupTitleForRt, year: rtTitleYear } = sanitizeTitleAndYear(rawLookupTitleForRt);
    const rtSearchYear = searchYear || (rtTitleYear ? parseInt(rtTitleYear, 10) : null);
    const rtData = fetchRottenTomatoesData(lookupTitleForRt, rtSearchYear);

    // If neither OMDb, IMDb Suggestion, nor Rotten Tomatoes found the movie:
    if (!data && !matchedImdbId && !rtData) {
      errorCell.setValue(`❌ Movie not found for "${cleanTitle}"`);
      return { status: "NOT_FOUND" };
    }

    // 4. Write IMDb Hyperlink to Column A (Always updates the hyperlink URL; updates title text on manual fix or if requested)
    if (matchedImdbId) {
      const displayTitle = options.forcedTitle
        ? options.forcedTitle
        : ((!options.preserveUserTitle && fuzzyMatchedTitle) ? fuzzyMatchedTitle : rawDisplayTitle);
      const safeTitle = displayTitle.replace(/"/g, '""');
      titleCell.setFormula(`=HYPERLINK("https://www.imdb.com/title/${matchedImdbId}/", "${safeTitle}")`);
    }

    // Detect whether the movie is a Netflix exclusive (via Rotten Tomatoes, OMDb Production, or existing Notes)
    const isNetflix = Boolean(
      (rtData && rtData.isNetflix) ||
      (data && data.Production && /\bnetflix\b/i.test(data.Production)) ||
      isNetflixNote(existingNoteText)
    );

    // 5. Determine Release Date:
    //    - For Netflix exclusives, prefer the RT Streaming release date (since it goes directly to streaming)
    //    - Otherwise: OMDb Released -> RT Theatrical Date -> RT Streaming Date -> Year fallback
    let apiDate = null;
    let isYearOnly = false;
    if (isNetflix && rtData && rtData.vodDateStr) {
      apiDate = parseSafeLocalDate(rtData.vodDateStr);
    }
    if (!apiDate && data && data.Released && data.Released !== "N/A") {
      apiDate = parseSafeLocalDate(data.Released);
    }
    if (!apiDate && rtData && rtData.theatersDateStr) {
      apiDate = parseSafeLocalDate(rtData.theatersDateStr);
    }
    if (!apiDate && rtData && rtData.vodDateStr) {
      apiDate = parseSafeLocalDate(rtData.vodDateStr);
    }

    if (apiDate instanceof Date && apiDate.getFullYear() === 1905) {
      const sanitizedApi = sanitizeAndRecoverDate(apiDate, searchYear);
      apiDate = sanitizedApi ? sanitizedApi.date : null;
    }

    let finalCategory = currentCategory;
    let finalDate = existingDate;

    // Determine fallback year from options, API responses, user title, or recovered existing cell
    const rawFallbackYear = options.forcedYear || (data && data.Year) || (suggestion && suggestion.year) || year || (existingRecovered ? existingRecovered.year : null) || searchYear;
    const fallbackYearMatch = String(rawFallbackYear || "").match(/\d{4}/);
    const fallbackYear = fallbackYearMatch ? parseInt(fallbackYearMatch[0], 10) : null;

    if (apiDate && !isNaN(apiDate.getTime())) {
      finalDate = apiDate;
      isYearOnly = false;
    } else if (!existingDate || existingIsCorrupt1905 || (fallbackYear && (!existingRecovered || existingRecovered.year !== fallbackYear))) {
      // If no exact API date, or previous date was corrupted 1905, use fallback year!
      if (fallbackYear) {
        finalDate = (fallbackYear >= today.getFullYear()) ? new Date(fallbackYear, 11, 31) : new Date(fallbackYear, 0, 1);
        isYearOnly = true;
      }
    } else if (existingRecovered && existingRecovered.isYearOnly) {
      isYearOnly = true;
    }

    // Prevent overwriting a known future date on an announced / "To be released" movie with an older namesake's date
    const isAnnouncedFuture = (fallbackYear && fallbackYear >= today.getFullYear()) || isToBeReleased;
    if (isAnnouncedFuture && (!apiDate || apiDate < sixMonthsAgo)) {
      if (fallbackYear && fallbackYear >= today.getFullYear()) {
        finalDate = new Date(fallbackYear, 11, 31);
        isYearOnly = !apiDate;
      } else if (existingDate && existingDate > today) {
        finalDate = existingDate;
      }
    }

    const hasValidDate = finalDate instanceof Date && !isNaN(finalDate.getTime());

    // Resolve VOD Date (for non-Netflix theatrical movies)
    let vodDateStr = (!isNetflix && rtData) ? rtData.vodDateStr : null;
    if (!isNetflix && !vodDateStr && hasValidDate && !isYearOnly && finalDate <= today && finalDate > sixMonthsAgo) {
      const vodCal = fetchDigitalReleaseCalendarMap();
      vodDateStr = (matchedImdbId && vodCal.byImdb[matchedImdbId]) || vodCal.byTitle[normalizeForCompare(cleanTitle)] || null;
    }
    const vodDateObj = isNetflix ? null : ((vodDateStr ? parseSafeLocalDate(vodDateStr) : null) || extractVodDateFromNote(existingNoteText));

    // Determine primary genre from OMDb or Rotten Tomatoes, with "Time Travel" override if movie involves time travel
    let primaryGenre = null;
    if (data && data.Genre && data.Genre !== "N/A") {
      primaryGenre = data.Genre.split(",")[0].trim();
    } else if (rtData && rtData.genre) {
      primaryGenre = rtData.genre;
    }

    const isTimeTravel = isMovieTimeTravel({
      currentCategory: currentCategory,
      imdbId: matchedImdbId,
      plotText: (data && data.Plot) || "",
      synopsisText: (rtData && rtData.synopsis) || "",
      noteText: existingNoteText
    });
    if (isTimeTravel) {
      primaryGenre = "Time Travel";
    }

    if (hasValidDate) {
      if (finalDate > today || (isYearOnly && fallbackYear && fallbackYear >= today.getFullYear())) {
        finalCategory = "To be released";
      } else if (finalDate <= today && finalDate > sixMonthsAgo && !isYearOnly) {
        // Transition rules between "In Theaters" and "New Releases":
        // - Netflix exclusives go directly to streaming -> "New Releases" (skipping "In Theaters")
        // - Theatrical movies transition from "In Theaters" to "New Releases" once VOD arrives (or after 60 days)
        if (isNetflix) {
          finalCategory = "New Releases";
        } else if (vodDateObj) {
          finalCategory = vodDateObj <= today ? "New Releases" : "In Theaters";
        } else if (finalDate <= sixtyDaysAgo) {
          finalCategory = "New Releases";
        } else if (isToBeReleased || String(currentCategory).toLowerCase().trim() === "in theaters" || !currentCategory) {
          finalCategory = "In Theaters";
        } else {
          finalCategory = "New Releases";
        }
      } else {
        finalCategory = primaryGenre || currentCategory;
      }
    } else if (isTimeTravel) {
      const curCatNorm = String(currentCategory || "").toLowerCase().trim();
      if (curCatNorm !== "to be released" && curCatNorm !== "in theaters" && curCatNorm !== "new releases") {
        finalCategory = "Time Travel";
      }
    }

    sheet.getRange(rowNum, 2).setValue(finalCategory);

    const dateCell = sheet.getRange(rowNum, 3);
    if (hasValidDate) {
      dateCell.setValue(finalDate);
      if (isYearOnly) {
        dateCell.setNumberFormat("yyyy");
      } else {
        dateCell.setNumberFormat("MMMM d, yyyy");
      }
    } else {
      dateCell.setValue("");
    }

    // 6. Rotten Tomatoes Popcornmeter (General Audience Score) + RT URL Hyperlink in Column D (Col 4)
    let popcornScore = "";
    const finalCatLower = String(finalCategory).toLowerCase().trim();
    if (finalCatLower !== "to be released" && rtData && rtData.popcornmeter) {
      popcornScore = rtData.popcornmeter;
    }
    writeRtScoreCell(sheet.getRange(rowNum, 4), popcornScore, rtData ? rtData.url : null);

    // 7. Write/Update Notes in Column F:
    //    - If Netflix exclusive -> write "Netflix" (for both upcoming and already released movies, replacing any VOD mention)
    //    - Otherwise -> write "VOD: <Date>" if in "In Theaters" or "New Releases"
    //    - If "To be released" with an unannounced release date (12/31 placeholder or year-only) -> write "Release Date Unannounced"
    //      until an exact release date is announced (at which point "Release Date Unannounced" is automatically removed)
    let updatedNoteText = existingNoteText;
    if (isNetflix) {
      updatedNoteText = upsertNetflixNote(updatedNoteText);
    } else if (vodDateStr && (finalCatLower === "in theaters" || finalCatLower === "new releases")) {
      updatedNoteText = upsertVodNote(updatedNoteText, vodDateStr);
    }

    const hasAnnouncedExactApiDate = Boolean(apiDate && !isNaN(apiDate.getTime()) && !isDec31Date(apiDate));
    const isUnannouncedRelease = (finalCatLower === "to be released") &&
      !hasAnnouncedExactApiDate &&
      (isYearOnly || !hasValidDate || isDec31Date(finalDate));

    if (isUnannouncedRelease) {
      updatedNoteText = upsertUnannouncedDateNote(updatedNoteText);
    } else {
      updatedNoteText = removeUnannouncedDateNote(updatedNoteText);
    }

    if (updatedNoteText !== existingNoteText) {
      sheet.getRange(rowNum, notesCol).setValue(updatedNoteText);
    }

    // 8. Clear Column G ("Errors") on success — only real errors are ever written to Column G
    const matchedTitle = (data && data.Title) || fuzzyMatchedTitle || (rtData && rtData.title) || cleanTitle;
    errorCell.setValue("");

    return normalizeForCompare(matchedTitle) !== normalizeForCompare(cleanTitle)
      ? { status: "FUZZY_MATCH", matchedTitle: matchedTitle }
      : { status: "SUCCESS", matchedTitle: matchedTitle };
  } catch (e) {
    errorCell.setValue(`⚠️ Lookup Error: ${e.message}`);
    return { status: "ERROR" };
  }
}

function writeRtScoreCell(cell, scoreText, rtUrl) {
  const cleanScore = String(scoreText || "").trim();
  if (!cleanScore) {
    cell.setValue("");
    return;
  }
  if (rtUrl) {
    const safeScore = cleanScore.replace(/"/g, '""');
    cell.setFormula(`=HYPERLINK("${rtUrl}", "${safeScore}")`);
  } else {
    cell.setValue(cleanScore);
  }
}

// ============================================================================
// 7. ROTTEN TOMATOES POPCORNMETER & STREAMING/VOD SCRAPER
// ============================================================================
function fetchRottenTomatoesData(title, targetYear) {
  try {
    const { cleanTitle, year: extractedYear } = sanitizeTitleAndYear(title);
    const clean = String(cleanTitle || "").trim();
    if (!clean) return null;
    const effectiveYear = targetYear || (extractedYear ? parseInt(extractedYear, 10) : null);

    const searchUrl = `https://www.rottentomatoes.com/search?search=${encodeURIComponent(clean)}`;
    const res = UrlFetchApp.fetch(searchUrl, {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
    });
    if (res.getResponseCode() !== 200) return null;

    const candidates = selectRtCandidates(res.getContentText(), clean, effectiveYear);
    if (!candidates || candidates.length === 0) return null;

    const chosen = candidates[0];
    const pageRes = UrlFetchApp.fetch(chosen.url, {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
    });
    if (pageRes.getResponseCode() !== 200) return null;

    const parsed = parseRtMoviePageHtml(pageRes.getContentText(), chosen);

    // If the top candidate page is an unscored duplicate stub (e.g. /m/paddington_2_2018)
    // and there is a second matching candidate (e.g. /m/paddington_2), try the second candidate!
    if ((!parsed || !parsed.popcornmeter) && candidates.length > 1) {
      const alt = candidates[1];
      const altRes = UrlFetchApp.fetch(alt.url, {
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
      });
      if (altRes.getResponseCode() === 200) {
        const altParsed = parseRtMoviePageHtml(altRes.getContentText(), alt);
        if (altParsed && altParsed.popcornmeter) {
          return altParsed;
        }
      }
    }

    return parsed;
  } catch (e) {
    return null;
  }
}

// Fetches Rotten Tomatoes data for a mini-batch of movies in parallel via UrlFetchApp.fetchAll
function fetchRottenTomatoesDataBatch(batchItems) {
  const results = new Array(batchItems.length).fill(null);
  if (!batchItems || batchItems.length === 0) return results;

  try {
    const sanitizedItems = batchItems.map(item => {
      const { cleanTitle, year } = sanitizeTitleAndYear(item.cleanTitle || item.rawTitle || "");
      return {
        cleanTitle: cleanTitle,
        searchYear: item.searchYear || (year ? parseInt(year, 10) : null)
      };
    });

    const searchRequests = sanitizedItems.map(item => ({
      url: `https://www.rottentomatoes.com/search?search=${encodeURIComponent(String(item.cleanTitle || "").trim())}`,
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
    }));

    const searchResponses = UrlFetchApp.fetchAll(searchRequests);
    const pageQueue = [];

    for (let i = 0; i < sanitizedItems.length; i++) {
      const res = searchResponses[i];
      if (!res || res.getResponseCode() !== 200) continue;
      const candidates = selectRtCandidates(res.getContentText(), sanitizedItems[i].cleanTitle, sanitizedItems[i].searchYear);
      if (candidates && candidates.length > 0 && candidates[0].url) {
        pageQueue.push({
          index: i,
          chosen: candidates[0],
          fallback: candidates.length > 1 ? candidates[1] : null
        });
      }
    }

    if (pageQueue.length > 0) {
      const pageRequests = pageQueue.map(entry => ({
        url: entry.chosen.url,
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
      }));
      const pageResponses = UrlFetchApp.fetchAll(pageRequests);
      const fallbackQueue = [];

      for (let j = 0; j < pageQueue.length; j++) {
        const pageRes = pageResponses[j];
        const entry = pageQueue[j];
        if (pageRes && pageRes.getResponseCode() === 200) {
          results[entry.index] = parseRtMoviePageHtml(pageRes.getContentText(), entry.chosen);
        }
        if ((!results[entry.index] || !results[entry.index].popcornmeter) && entry.fallback && entry.fallback.url) {
          fallbackQueue.push({ index: entry.index, chosen: entry.fallback });
        }
      }

      if (fallbackQueue.length > 0) {
        const fbRequests = fallbackQueue.map(fb => ({
          url: fb.chosen.url,
          muteHttpExceptions: true,
          headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
        }));
        const fbResponses = UrlFetchApp.fetchAll(fbRequests);
        for (let k = 0; k < fallbackQueue.length; k++) {
          const fbRes = fbResponses[k];
          if (!fbRes || fbRes.getResponseCode() !== 200) continue;
          const fbParsed = parseRtMoviePageHtml(fbRes.getContentText(), fallbackQueue[k].chosen);
          if (fbParsed && fbParsed.popcornmeter) {
            results[fallbackQueue[k].index] = fbParsed;
          }
        }
      }
    }

    return results;
  } catch (e) {
    // Fallback to sequential lookup if fetchAll encounters any unexpected error
    for (let i = 0; i < batchItems.length; i++) {
      results[i] = fetchRottenTomatoesData(batchItems[i].cleanTitle, batchItems[i].searchYear);
    }
    return results;
  }
}

function selectBestRtCandidate(searchHtml, cleanTitle, targetYear) {
  const ranked = selectRtCandidates(searchHtml, cleanTitle, targetYear);
  return ranked && ranked.length > 0 ? ranked[0] : null;
}

function selectRtCandidates(searchHtml, rawTitleInput, targetYear) {
  if (!searchHtml || !rawTitleInput) return [];
  const { cleanTitle, year: extractedYear } = sanitizeTitleAndYear(rawTitleInput);
  const effectiveYear = targetYear || extractedYear;
  const rowRegex = /<search-page-media-row([\s\S]*?)<\/search-page-media-row>/gi;
  const candidates = [];
  let match;
  let orderIdx = 0;

  while ((match = rowRegex.exec(searchHtml)) !== null) {
    const fullRow = match[0];
    const block = match[1];
    const hrefMatch = block.match(/href="(https:\/\/www\.rottentomatoes\.com\/m\/[^"]+)"/i);
    if (!hrefMatch) continue; // Only movie (/m/) links

    const yearMatch = fullRow.match(/release-year="(\d{4})"/i) || fullRow.match(/releaseyear="(\d{4})"/i);
    const nameMatch = block.match(/data-qa="info-name"[^>]*>([\s\S]*?)<\/a>/i);
    const candTitle = nameMatch ? decodeHtmlEntities(nameMatch[1].replace(/<[^>]+>/g, "").trim()) : "";
    const candYear = yearMatch ? parseInt(yearMatch[1], 10) : null;
    const hasScore = /tomatometer-score="\d+"/i.test(fullRow) || /audiencescore="\d+"/i.test(fullRow);

    candidates.push({
      title: candTitle,
      year: candYear,
      url: hrefMatch[1],
      hasScore: hasScore,
      orderIdx: orderIdx++
    });
  }

  if (candidates.length === 0) return [];

  const normTarget = normalizeForCompare(cleanTitle);
  const numYear = effectiveYear ? parseInt(String(effectiveYear).match(/\d{4}/)?.[0] || "", 10) : null;

  const scoredCandidates = [];
  for (const c of candidates) {
    const normCand = normalizeForCompare(c.title);
    const isExactTitle = (normCand === normTarget);
    const isOverlapTitle = isExactTitle ||
      normCand.includes(normTarget) ||
      normTarget.includes(normCand) ||
      doesRtSlugMatchTitle(c.url, cleanTitle);

    if (!isOverlapTitle) continue; // Never select unrelated movies

    const yearDiff = (numYear && c.year) ? Math.abs(c.year - numYear) : null;
    const withinOneYear = (yearDiff !== null && yearDiff <= 1);

    let tier = 50;
    if (numYear) {
      if (isExactTitle && withinOneYear) {
        // Scored candidate within ±1 year ALWAYS beats an unscored duplicate stub (like /m/paddington_2_2018)!
        if (c.hasScore && yearDiff === 0) tier = 1;
        else if (c.hasScore && yearDiff === 1) tier = 2;
        else if (yearDiff === 0) tier = 3;
        else tier = 4;
      } else if (isExactTitle) {
        tier = c.hasScore ? 10 : 11;
      } else if (withinOneYear) {
        if (c.hasScore && yearDiff === 0) tier = 20;
        else if (c.hasScore && yearDiff === 1) tier = 21;
        else if (yearDiff === 0) tier = 22;
        else tier = 23;
      } else {
        tier = c.hasScore ? 30 : 31;
      }
    } else {
      if (isExactTitle) {
        tier = c.hasScore ? 5 : 6;
      } else {
        tier = c.hasScore ? 25 : 26;
      }
    }

    scoredCandidates.push({
      candidate: c,
      tier: tier,
      orderIdx: c.orderIdx
    });
  }

  scoredCandidates.sort((a, b) => {
    if (a.tier !== b.tier) return a.tier - b.tier;
    return a.orderIdx - b.orderIdx;
  });

  return scoredCandidates.map(s => s.candidate);
}

function parseRtMoviePageHtml(pageHtml, chosen) {
  if (!pageHtml || !chosen) return null;
  let popcornmeter = null;
  let criticsScore = null;

  // Extract JSON from <script id="media-scorecard-json">
  const scorecardMatch = pageHtml.match(/<script[^>]*id="media-scorecard-json"[^>]*>([\s\S]*?)<\/script>/i);
  if (scorecardMatch) {
    try {
      const scJson = JSON.parse(scorecardMatch[1].trim());
      if (scJson.audienceScore && scJson.audienceScore.scorePercent) {
        popcornmeter = scJson.audienceScore.scorePercent;
      } else if (scJson.audienceScore && scJson.audienceScore.score) {
        popcornmeter = `${scJson.audienceScore.score}%`;
      }
      if (scJson.criticsScore && scJson.criticsScore.scorePercent) {
        criticsScore = scJson.criticsScore.scorePercent;
      } else if (scJson.criticsScore && scJson.criticsScore.score) {
        criticsScore = `${scJson.criticsScore.score}%`;
      }
    } catch (e) {}
  }

  // Fallback: if a movie is listed on RT with a Tomatometer score but fewer than 50 audience ratings, use criticsScore
  if (!popcornmeter && criticsScore) {
    popcornmeter = criticsScore;
  }

  // Extract Release Date (Streaming) -> VOD Date
  let vodDateStr = null;
  const vodMatch = pageHtml.match(/Release Date \(Streaming\)<\/rt-text>\s*<\/dt>\s*<dd[^>]*>\s*<rt-text[^>]*>([^<]+)<\/rt-text>/i);
  if (vodMatch) {
    vodDateStr = vodMatch[1].replace(/,.*$/, "").trim();
  }

  // Extract Release Date (Theaters)
  let theatersDateStr = null;
  const theatersMatch = pageHtml.match(/Release Date \(Theaters\)<\/rt-text>\s*<\/dt>\s*<dd[^>]*>\s*<rt-text[^>]*>([^<]+)<\/rt-text>/i);
  if (theatersMatch) {
    theatersDateStr = theatersMatch[1].replace(/,\s*(Wide|Limited|Original).*$/i, "").trim();
  }

  // Extract Genre from <script id="media-hero-json">
  let genre = null;
  const heroMatch = pageHtml.match(/<script[^>]*id="media-hero-json"[^>]*>([\s\S]*?)<\/script>/i);
  if (heroMatch) {
    try {
      const heroJson = JSON.parse(heroMatch[1].trim());
      if (heroJson.content && Array.isArray(heroJson.content.metadataGenres) && heroJson.content.metadataGenres.length > 0) {
        genre = heroJson.content.metadataGenres[0];
      }
    } catch (e) {}
  }

  // Extract Synopsis
  let synopsis = null;
  const synopsisMatch = pageHtml.match(/data-qa=["']synopsis-value["'][^>]*>([\s\S]*?)<\/rt-text>/i);
  if (synopsisMatch) {
    synopsis = decodeHtmlEntities(synopsisMatch[1].replace(/<[^>]+>/g, "").trim());
  }
  if (!synopsis) {
    const descJsonMatch = pageHtml.match(/data-AdsMediaScorecardManager=["']description["'][^>]*>[\s\S]*?"description":\s*"([^"]+)"/i);
    if (descJsonMatch) {
      try {
        synopsis = JSON.parse(`"${descJsonMatch[1]}"`);
      } catch (e) {
        synopsis = descJsonMatch[1];
      }
    }
  }

  // Extract Distributor & Production Co to detect Netflix exclusives
  let distributor = null;
  const distMatch = pageHtml.match(/Distributor<\/rt-text>\s*<\/dt>\s*<dd[^>]*>([\s\S]*?)<\/dd>/i);
  if (distMatch) {
    distributor = decodeHtmlEntities(distMatch[1].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim());
  }

  let productionCo = null;
  const prodMatch = pageHtml.match(/Production Co<\/rt-text>\s*<\/dt>\s*<dd[^>]*>([\s\S]*?)<\/dd>/i);
  if (prodMatch) {
    productionCo = decodeHtmlEntities(prodMatch[1].replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim());
  }

  const isNetflix = Boolean(
    (distributor && /\bnetflix\b/i.test(distributor)) ||
    (!distributor && productionCo && /\bnetflix\b/i.test(productionCo))
  );

  return {
    url: chosen.url,
    title: chosen.title,
    year: chosen.year,
    popcornmeter: popcornmeter,
    criticsScore: criticsScore,
    vodDateStr: vodDateStr,
    theatersDateStr: theatersDateStr,
    genre: genre,
    synopsis: synopsis,
    distributor: distributor,
    isNetflix: isNetflix
  };
}

function extractRtUrlFromFormula(formulaVal) {
  if (!formulaVal) return null;
  const m = String(formulaVal).match(/https?:\/\/(?:www\.)?rottentomatoes\.com\/m\/[^"'?#\s)]+/i);
  return m ? m[0] : null;
}

function doesRtSlugMatchTitle(rtUrl, title) {
  if (!rtUrl || !title) return false;
  const m = String(rtUrl).match(/rottentomatoes\.com\/m\/([^"'?#\s/)]+)/i);
  if (!m) return false;

  // Strip optional leading numeric ID (e.g. "1000013-12_angry_men") and optional trailing year (e.g. "_2024")
  const rawSlug = m[1]
    .replace(/^\d+[-_]/, "")
    .replace(/[-_](19|20)\d{2}$/, "");

  const normSlug = normalizeForCompare(rawSlug);
  const normTitle = normalizeForCompare(title);
  if (!normSlug || !normTitle) return false;

  if (normSlug === normTitle) return true;

  const stripArticle = s => s.replace(/^(the|a|an)/, "");
  const slugNoArt = stripArticle(normSlug);
  const titleNoArt = stripArticle(normTitle);
  if (slugNoArt && titleNoArt && slugNoArt === titleNoArt) return true;

  // Substring match when the shorter normalized string is at least 4 chars
  const minLen = Math.min(slugNoArt.length, titleNoArt.length);
  if (minLen >= 4 && (slugNoArt.includes(titleNoArt) || titleNoArt.includes(slugNoArt))) {
    return true;
  }

  // Word-level overlap check
  const stopWords = new Set(["the", "and", "for", "with", "from", "into", "part", "movie", "film", "vol", "volume"]);
  const slugWords = rawSlug
    .toLowerCase()
    .split(/[-_]+/)
    .map(w => w.replace(/[^a-z0-9]/g, ""))
    .filter(w => w.length >= 2 && !stopWords.has(w));
  const titleWords = String(title)
    .toLowerCase()
    .replace(/&/g, "and")
    .split(/[^a-z0-9]+/)
    .filter(w => w.length >= 2 && !stopWords.has(w));

  if (titleWords.length > 0 && slugWords.length > 0) {
    const slugWordSet = new Set(slugWords);
    let matches = 0;
    for (const tw of titleWords) {
      if (
        slugWordSet.has(tw) ||
        slugWords.some(sw => sw.length >= 4 && tw.length >= 4 && (sw.includes(tw) || tw.includes(sw)))
      ) {
        matches++;
      }
    }
    if (matches >= Math.ceil(titleWords.length * 0.5)) return true;
  }

  return false;
}

// Verifies the current live row of a movie right before writing so row deletions/moves never shift scores
function findCurrentRowForMovie(sheet, expectedRow, expectedRawTitle) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const liveColA = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  return resolveRowFromLiveColA(liveColA, expectedRow, expectedRawTitle);
}

function resolveRowFromLiveColA(liveColA, expectedRow, expectedRawTitle) {
  const targetNorm = normalizeForCompare(expectedRawTitle);
  if (!targetNorm || !liveColA || liveColA.length === 0) return null;

  // 1. Check expectedRow first (index = expectedRow - 2)
  const expectedIdx = expectedRow - 2;
  if (expectedIdx >= 0 && expectedIdx < liveColA.length) {
    if (normalizeForCompare(liveColA[expectedIdx][0]) === targetNorm) {
      return expectedRow;
    }
  }

  // 2. If rows shifted (e.g. a row above was deleted or moved to "Won't Watch"), find closest matching row
  let bestRow = null;
  let bestDist = Infinity;
  for (let idx = 0; idx < liveColA.length; idx++) {
    if (normalizeForCompare(liveColA[idx][0]) === targetNorm) {
      const candidateRow = idx + 2;
      const dist = Math.abs(candidateRow - expectedRow);
      if (dist < bestDist) {
        bestDist = dist;
        bestRow = candidateRow;
      }
    }
  }
  return bestRow;
}

// ============================================================================
// 8. DIGITAL / VOD RELEASE CALENDAR SCRAPER (DVDsReleaseDates.com)
// ============================================================================
let cachedVodCalendar = null;

function fetchDigitalReleaseCalendarMap() {
  if (cachedVodCalendar) return cachedVodCalendar;
  const byImdb = {};
  const byTitle = {};

  try {
    const baseUrl = "https://www.dvdsreleasedates.com";
    const mainRes = UrlFetchApp.fetch(`${baseUrl}/digital-releases/`, {
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
    });

    if (mainRes.getResponseCode() === 200) {
      const mainHtml = mainRes.getContentText();
      parseDvdReleaseDatesHtml(mainHtml, byImdb, byTitle);

      // Also grab the next upcoming month link from the page navigation
      const monthLinks = [];
      const linkRegex = /href=['"](\/digital-releases\/\d{4}\/\d{1,2}\/[^'"]+)['"]/gi;
      let m;
      while ((m = linkRegex.exec(mainHtml)) !== null) {
        if (!monthLinks.includes(m[1])) monthLinks.push(m[1]);
      }
      for (let i = 0; i < Math.min(2, monthLinks.length); i++) {
        const subRes = UrlFetchApp.fetch(`${baseUrl}${monthLinks[i]}`, {
          muteHttpExceptions: true,
          headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
        });
        if (subRes.getResponseCode() === 200) {
          parseDvdReleaseDatesHtml(subRes.getContentText(), byImdb, byTitle);
        }
      }
    }
  } catch (e) {}

  cachedVodCalendar = { byImdb, byTitle };
  return cachedVodCalendar;
}

function parseDvdReleaseDatesHtml(html, byImdb, byTitle) {
  const sections = html.split(/class=['"]reldate[^'"]*['"][^>]*>(?:<a[^>]*><\/a>)?/i);
  for (let i = 1; i < sections.length; i++) {
    const sec = sections[i];
    const dateEndIdx = sec.indexOf("<");
    if (dateEndIdx === -1) continue;

    const rawDateHeader = sec.substring(0, dateEndIdx).replace(/^(Monday|Tuesday|Wednesday|Friday|Saturday|Sunday|Thursday)\s+/i, "").trim();
    const parsedDate = parseSafeLocalDate(rawDateHeader);
    const formattedDate = parsedDate
      ? Utilities.formatDate(parsedDate, Session.getScriptTimeZone() || "America/New_York", "MMM d, yyyy")
      : rawDateHeader;

    const cells = sec.split(/<td class=['"]dvdcell['"]>/i);
    for (let c = 1; c < cells.length; c++) {
      const cellHtml = cells[c];
      const imdbMatch = cellHtml.match(/imdb\.com\/title\/(tt\d+)/i);
      const titleMatch = cellHtml.match(/<a style=['"]color:#000;['"] href=['"]\/movies\/\d+\/[^'"]+['"]>([^<]+)<\/a>/i);

      if (imdbMatch) {
        byImdb[imdbMatch[1].toLowerCase()] = formattedDate;
      }
      if (titleMatch) {
        const norm = normalizeForCompare(decodeHtmlEntities(titleMatch[1]));
        if (norm) byTitle[norm] = formattedDate;
      }
    }
  }
}

// ============================================================================
// 9. IMDB SUGGESTION API & CANDIDATE SEARCH HELPERS
// ============================================================================
function findImdbIdWithTypoTolerance(query, targetYear, lookForwardYear) {
  const candidates = fetchMovieCandidates(query, 10);
  if (candidates.length === 0) return null;

  const normQuery = normalizeForCompare(query);

  // 1. Exact title + exact year match
  if (targetYear) {
    const exactBoth = candidates.find(
      c => String(c.year) === String(targetYear) && normalizeForCompare(c.title) === normQuery
    );
    if (exactBoth) return { imdbID: exactBoth.imdbID, matchedTitle: exactBoth.title, year: exactBoth.year };

    // 2. Exact year match ONLY if title also overlaps (prevents matching "Paul" to another random 2011 film!)
    const yearAndPartialTitle = candidates.find(
      c => String(c.year) === String(targetYear) && (
        normalizeForCompare(c.title).includes(normQuery) || normQuery.includes(normalizeForCompare(c.title))
      )
    );
    if (yearAndPartialTitle) {
      return { imdbID: yearAndPartialTitle.imdbID, matchedTitle: yearAndPartialTitle.title, year: yearAndPartialTitle.year };
    }
  }

  // 3. Look-forward year for "To be released" movies
  if (lookForwardYear) {
    const futureExact = candidates.find(
      c => (!c.year || c.year >= lookForwardYear) && normalizeForCompare(c.title) === normQuery
    );
    if (futureExact) return { imdbID: futureExact.imdbID, matchedTitle: futureExact.title, year: futureExact.year };

    const futureAny = candidates.find(c => !c.year || c.year >= lookForwardYear);
    if (futureAny) return { imdbID: futureAny.imdbID, matchedTitle: futureAny.title, year: futureAny.year };
  }

  // 4. Exact title match regardless of year
  const exactTitle = candidates.find(c => normalizeForCompare(c.title) === normQuery);
  if (exactTitle) return { imdbID: exactTitle.imdbID, matchedTitle: exactTitle.title, year: exactTitle.year };

  // 5. Default to top-ranked movie candidate
  return { imdbID: candidates[0].imdbID, matchedTitle: candidates[0].title, year: candidates[0].year };
}

// Uses IMDb's `/suggestion/titles/` endpoint so actor names (like "Paul") never crowd out movies
function fetchMovieCandidates(query, maxResults) {
  maxResults = maxResults || 8;
  const results = [];
  const seenIds = new Set();

  try {
    const cleanQuery = String(query || "").toLowerCase().replace(/[^a-z0-9\s]/g, "").trim();
    if (!cleanQuery) return results;
    const firstChar = cleanQuery.charAt(0);

    // `/suggestion/titles/` guarantees title results (`tt...`) instead of actors (`nm...`)
    const url = `https://v3.sg.media-imdb.com/suggestion/titles/${firstChar}/${encodeURIComponent(cleanQuery)}.json`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      if (json.d && Array.isArray(json.d)) {
        // Prioritize feature films / movies over TV episodes or video games
        const sorted = json.d
          .filter(item => item.id && item.id.startsWith("tt"))
          .sort((a, b) => {
            const aIsMovie = (a.qid === "movie" || a.qid === "tvMovie" || !a.qid) ? 0 : 1;
            const bIsMovie = (b.qid === "movie" || b.qid === "tvMovie" || !b.qid) ? 0 : 1;
            return aIsMovie - bIsMovie;
          });

        for (const item of sorted) {
          if (seenIds.has(item.id)) continue;
          seenIds.add(item.id);
          results.push({
            imdbID: item.id,
            title: item.l || query,
            year: item.y || null,
            type: item.qid || item.q || "movie",
            stars: item.s || "",
            poster: (item.i && item.i.imageUrl) ? item.i.imageUrl : ""
          });
          if (results.length >= maxResults) break;
        }
      }
    }
  } catch (e) {}

  return results;
}

// Looks up exact movie title & year by IMDb ID (e.g. "tt1092026") using IMDb's free suggestion endpoint
function fetchImdbDetailsById(imdbId) {
  try {
    const cleanId = String(imdbId || "").trim().toLowerCase();
    if (!/^tt\d+$/.test(cleanId)) return null;
    const url = `https://v3.sg.media-imdb.com/suggestion/t/${encodeURIComponent(cleanId)}.json`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const json = JSON.parse(res.getContentText());
      if (json.d && Array.isArray(json.d)) {
        const exact = json.d.find(item => String(item.id || "").toLowerCase() === cleanId) || json.d[0];
        if (exact) {
          return {
            imdbID: cleanId,
            title: exact.l || null,
            year: exact.y || null,
            stars: exact.s || "",
            poster: (exact.i && exact.i.imageUrl) ? exact.i.imageUrl : ""
          };
        }
      }
    }
  } catch (e) {}
  return null;
}

function fetchOmdbJson(url) {
  if (omdbQuotaExceeded || !apiKey) return null;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const text = res.getContentText();
    if (!text) return null;
    const parsed = JSON.parse(text);
    if (parsed && parsed.Error && String(parsed.Error).toLowerCase().includes("limit reached")) {
      omdbQuotaExceeded = true;
      return null;
    }
    return parsed;
  } catch (e) {
    return null;
  }
}

// ============================================================================
// 10. SHEET FORMATTING, SORTING & LIFECYCLE PROMOTION HELPERS
// ============================================================================
function checkAndFormatSheet(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return;

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const numCols = Math.max(7, sheet.getLastColumn());
  const notesColIdx = findNotesColumn(sheet) - 1; // 0-based index for Notes column
  const priorityColIdx = findPriorityColumn(sheet) - 1; // 0-based index for Priority column
  const range = sheet.getRange(1, 1, lastRow, numCols);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();

  // Clear legacy cell notes on Column A
  sheet.getRange(2, 1, lastRow - 1, 1).clearNote();

  const rows = [];
  const wontWatchRows = [];
  const completedRows = [];
  const completedTimestamp = new Date();
  completedTimestamp.setHours(0, 0, 0, 0);

  for (let r = 1; r < values.length; r++) {
    const rawColA = String(values[r][0] || "").trim();
    const rawColB = String(values[r][1] || "").trim();
    // Skip empty rows (no title, no category, and no formula in Column A)
    if (!rawColA && !rawColB && !formulas[r][0]) {
      continue;
    }

    const rowData = [];
    for (let c = 0; c < numCols; c++) {
      if (c === 0) {
        // Preserve Rich Text hyperlinks by converting them to =HYPERLINK() before sorting!
        if (formulas[r][0]) {
          rowData.push(formulas[r][0]);
        } else {
          const rtLink = richText[r][0] ? richText[r][0].getLinkUrl() : null;
          const cellText = String(values[r][0] || "").trim();
          if (rtLink && cellText) {
            rowData.push(`=HYPERLINK("${rtLink}", "${cellText.replace(/"/g, '""')}")`);
          } else {
            rowData.push(values[r][0]);
          }
        }
      } else if (c === 6) {
        // Column G ("Errors"): only keep real errors (starting with ❌ or ⚠️), strip any old log messages
        const errText = String(values[r][c] || "").trim();
        rowData.push((errText.startsWith("❌") || errText.startsWith("⚠️")) ? errText : "");
      } else {
        rowData.push(formulas[r][c] ? formulas[r][c] : values[r][c]);
      }
    }

    // If Priority (or Category) is set to "Completed", queue row to move to "History" tab (Col A = Title w/ IMDb link, Col B = Today's Date)
    if (isCompletedValue(values[r][priorityColIdx]) || isCompletedValue(values[r][1])) {
      completedRows.push([rowData[0], completedTimestamp]);
      continue;
    }

    // If Priority (or Category) is set to "Won't Watch", queue row to move to "Won't Watch" tab
    if (isWontWatchValue(values[r][priorityColIdx]) || isWontWatchValue(values[r][1])) {
      wontWatchRows.push(rowData);
      continue;
    }

    rows.push({
      data: rowData,
      rawDate: values[r][2]
    });
  }

  // Move any queued "Completed" rows to the bottom of the "History" tab
  if (completedRows.length > 0) {
    appendRowsToHistory(completedRows);
  }

  // Move any queued "Won't Watch" rows to the bottom of the "Won't Watch" tab
  if (wontWatchRows.length > 0) {
    appendRowsToWontWatch(wontWatchRows);
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sixtyDaysAgo = new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000);

  // Automatic Category Promotions (Zero API calls required!):
  // 1. "To be released" -> "New Releases" (if Netflix exclusive) or "In Theaters" once Release Date (Col C) <= today
  // 2. "In Theaters" -> "New Releases" once the "VOD: <Date>" in Notes <= today
  //    (or if Netflix exclusive, or if theatrical release date was 60+ days ago and no future VOD date is listed)
  for (let i = 0; i < rows.length; i++) {
    // Sanitize and recover any 1905 serial dates or raw year values in Column C
    const rawVal = rows[i].rawDate;
    const recovered = sanitizeAndRecoverDate(rawVal);
    let d = recovered ? recovered.date : (rawVal instanceof Date && !isNaN(rawVal.getTime()) ? rawVal : null);
    const isYearOnly = recovered ? recovered.isYearOnly : false;
    if (recovered && recovered.date) {
      rows[i].data[2] = recovered.date;
      rows[i].rawDate = recovered.date;
    }

    let cat = String(rows[i].data[1]).toLowerCase().trim();
    const hasValidDate = d instanceof Date && !isNaN(d.getTime());
    let noteText = String(rows[i].data[notesColIdx] || "");
    const isNetflix = isNetflixNote(noteText);
    const isUnannouncedDate = isYearOnly || !hasValidDate || isDec31Date(d);

    // If marked Netflix in Notes, ensure any old "VOD: ..." mention is replaced by "Netflix"
    if (isNetflix) {
      noteText = upsertNetflixNote(noteText);
    }

    // Only promote from "To be released" if it has an exact announced release date (not year-only or 12/31 placeholder):
    // - Netflix exclusives go directly to "New Releases" (skipping "In Theaters")
    // - Theatrical movies go to "In Theaters"
    if (cat === "to be released" && hasValidDate && !isUnannouncedDate && d <= today) {
      if (isNetflix) {
        rows[i].data[1] = "New Releases";
        cat = "new releases";
      } else {
        rows[i].data[1] = "In Theaters";
        cat = "in theaters";
      }
    }

    if (cat === "in theaters") {
      if (isNetflix) {
        rows[i].data[1] = "New Releases";
        cat = "new releases";
      } else {
        const vodDate = extractVodDateFromNote(noteText);
        if (vodDate && vodDate <= today) {
          rows[i].data[1] = "New Releases";
          cat = "new releases";
        } else if (!vodDate && hasValidDate && !isYearOnly && d <= sixtyDaysAgo) {
          rows[i].data[1] = "New Releases";
          cat = "new releases";
        }
      }
    }

    // Maintain "Release Date Unannounced" in Notes (Column F):
    // - If "To be released" with a 12/31 placeholder date (or year-only / blank date), set "Release Date Unannounced"
    // - Once an exact non-12/31 date is set (or movie moves out of "To be released"), remove "Release Date Unannounced"
    if (cat === "to be released" && isUnannouncedDate) {
      noteText = upsertUnannouncedDateNote(noteText);
    } else if (isUnannouncedDateNote(noteText)) {
      noteText = removeUnannouncedDateNote(noteText);
    }

    rows[i].data[notesColIdx] = noteText;
  }

  // Automatic "Time Travel" Genre Categorization:
  // For any movie in its genre stage (not "To be released", "In Theaters", or "New Releases"),
  // if it involves time travel in any way (via Wikidata, Wikipedia film categories, or notes),
  // set its Category (Column B) to "Time Travel".
  const genreImdbIds = [];
  for (let i = 0; i < rows.length; i++) {
    const c = String(rows[i].data[1] || "").toLowerCase().trim();
    if (c === "to be released" || c === "in theaters" || c === "new releases" || c === "time travel") continue;
    const imdbId = extractImdbId(rows[i].data[0], null);
    if (imdbId) genreImdbIds.push(imdbId);
  }
  const timeTravelMap = genreImdbIds.length > 0 ? fetchTimeTravelImdbSet(genreImdbIds) : {};

  for (let i = 0; i < rows.length; i++) {
    const c = String(rows[i].data[1] || "").toLowerCase().trim();
    if (c === "to be released" || c === "in theaters" || c === "new releases" || c === "time travel") continue;
    const imdbId = extractImdbId(rows[i].data[0], null);
    const noteText = String(rows[i].data[notesColIdx] || "");
    if ((imdbId && timeTravelMap[imdbId]) || isTimeTravelText(noteText)) {
      rows[i].data[1] = "Time Travel";
    }
  }

  const getCategoryRank = (cat) => {
    const c = String(cat).toLowerCase().trim();
    if (c === "to be released") return 1;
    if (c === "in theaters") return 2;
    if (c === "new releases") return 3;
    return 4;
  };

  // Priority order for genre movies: High -> Medium -> Low -> anything else -> blank
  const getPriorityRank = (val) => {
    const p = String(val || "").toLowerCase().trim();
    if (!p) return 5;
    if (p.includes("high")) return 1;
    if (p.includes("med")) return 2;
    if (p.includes("low")) return 3;
    return 4;
  };

  rows.sort((a, b) => {
    const rankA = getCategoryRank(a.data[1]);
    const rankB = getCategoryRank(b.data[1]);
    if (rankA !== rankB) return rankA - rankB;

    const dateA = a.rawDate instanceof Date ? a.rawDate.getTime() : new Date(a.rawDate).getTime();
    const dateB = b.rawDate instanceof Date ? b.rawDate.getTime() : new Date(b.rawDate).getTime();

    const timeA = isNaN(dateA) ? Infinity : dateA;
    const timeB = isNaN(dateB) ? Infinity : dateB;

    // Genre movies (after "New Releases"): sort by Priority (High -> Medium -> Low -> other/blank),
    // then newest to oldest within each priority (undated rows at the bottom of their priority group)
    if (rankA === 4) {
      const prioA = getPriorityRank(a.data[priorityColIdx]);
      const prioB = getPriorityRank(b.data[priorityColIdx]);
      if (prioA !== prioB) return prioA - prioB;

      const descA = isNaN(dateA) ? -Infinity : dateA;
      const descB = isNaN(dateB) ? -Infinity : dateB;
      if (descA === descB) return 0;
      return descB - descA;
    }

    // "To be released", "In Theaters", "New Releases": unchanged (oldest to newest)
    if (timeA === timeB) return 0;
    return timeA - timeB;
  });

  const sortedValues = rows.map(r => r.data);

  // If all rows were moved to "Won't Watch", clear remaining movie rows
  if (sortedValues.length === 0) {
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
    }
    return;
  }

  const dataRange = sheet.getRange(2, 1, sortedValues.length, numCols);
  dataRange.setValues(sortedValues);

  // Post a Discord "Now available" notification for any movie that just moved into "New Releases"
  try {
    processNewReleaseNotifications(sortedValues);
  } catch (e) {}

  // Apply number formats to Column C: "yyyy" for year-only dates, "MMMM d, yyyy" for exact dates
  const dateFormats = [];
  for (let i = 0; i < sortedValues.length; i++) {
    const val = sortedValues[i][2];
    const rec = sanitizeAndRecoverDate(val);
    if (rec && rec.isYearOnly) {
      dateFormats.push(["yyyy"]);
    } else if (val instanceof Date && !isNaN(val.getTime())) {
      dateFormats.push(["MMMM d, yyyy"]);
    } else {
      dateFormats.push(["@"]);
    }
  }
  if (dateFormats.length > 0) {
    sheet.getRange(2, 3, dateFormats.length, 1).setNumberFormats(dateFormats);
  }

  // Ensure there is ALWAYS 1 clean empty row at the bottom for easy movie entry on mobile/desktop
  const emptyRowIndex = sortedValues.length + 2;
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < emptyRowIndex) {
    sheet.insertRowsAfter(currentMaxRows, emptyRowIndex - currentMaxRows);
  }

  // Clear content, notes, and background on the designated empty row
  sheet.getRange(emptyRowIndex, 1, 1, numCols).clearContent().clearNote();
  sheet.getRange(emptyRowIndex, 1, 1, 3).setBackground(null);

  // If previous rows existed below the single empty row (from "Won't Watch" movements or stale entries), remove them
  const currentLastRow = sheet.getLastRow();
  if (currentLastRow > emptyRowIndex) {
    sheet.deleteRows(emptyRowIndex + 1, currentLastRow - emptyRowIndex);
  }

  // Ensure "Won't Watch" is included in the Priority column's dropdown options, covering the empty row
  ensureWontWatchInPriorityDropdown(sheet, sortedValues.length + 1);

  // Remove alternating row Bandings
  sheet.getBandings().forEach(b => b.remove());

  // Reset and apply solid category colors ONLY to Columns A–C (leaving Columns D+ untouched)
  const colorRange = sheet.getRange(2, 1, sortedValues.length, 3);
  colorRange.setBackground(null);
  SpreadsheetApp.flush();

  const backgrounds = [];
  for (let i = 0; i < sortedValues.length; i++) {
    const c = String(sortedValues[i][1]).toLowerCase().trim();
    let color = "#ffffff";
    if (c === "to be released") {
      color = "#e0e0e0"; // Solid Grey
    } else if (c === "in theaters") {
      color = "#d2e3fc"; // Solid Blue
    } else if (c === "new releases") {
      color = "#ceead6"; // Solid Green
    }
    backgrounds.push([color, color, color]);
  }

  colorRange.setBackgrounds(backgrounds);
}

// ============================================================================
// "COMPLETED" -> "HISTORY" & "WON'T WATCH" PRIORITY TIER / TAB MOVEMENT HELPERS
// ============================================================================
function isCompletedValue(val) {
  const norm = String(val || "").toLowerCase().trim();
  return norm === "completed" || norm === "complete";
}

function isWontWatchValue(val) {
  const norm = String(val || "").toLowerCase().replace(/['’`]/g, "").trim();
  return norm === "wont watch" || norm === "do not watch" || norm === "dont watch";
}

function findPriorityColumn(sheet) {
  const lastCol = Math.max(7, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || "").toLowerCase().trim();
    if (h.includes("priority") || h.includes("tier") || h.includes("interest")) {
      return c + 1;
    }
  }
  return 5; // Default to Column E (Col 5) if no explicit "Priority" header is found
}

function ensureWontWatchInPriorityDropdown(sheet, numDataRows) {
  try {
    const priorityCol = findPriorityColumn(sheet);
    const totalRows = numDataRows || Math.max(1, sheet.getLastRow() - 1);
    if (totalRows <= 1) return;

    const sourceCell = sheet.getRange(2, priorityCol);
    const sampleRule = sourceCell.getDataValidation();
    if (!sampleRule) return;

    // IMPORTANT: Never call setDataValidation(sampleRule.copy().requireValueInList(...).build())!
    // Apps Script's DataValidationBuilder strips Google Sheets' modern "Chip" display style
    // and resets all custom dropdown chip colors to gray "Arrow" dropdowns.
    // Instead, only copy validation via PASTE_DATA_VALIDATION to any bottom row missing it,
    // which preserves the exact "Chip" style and custom option colors.
    const bottomCell = sheet.getRange(totalRows + 1, priorityCol);
    if (!bottomCell.getDataValidation()) {
      sourceCell.copyTo(
        bottomCell,
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
        false
      );
    }
  } catch (e) {}
}

function getOrCreateHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (String(s.getName() || "").toLowerCase().trim() === "history") {
      return s;
    }
  }
  const created = ss.insertSheet("History");
  created.getRange(1, 1, 1, 2).setValues([["Movie", "Watch Date"]]);
  return created;
}

function appendRowsToHistory(historyRows) {
  if (!historyRows || historyRows.length === 0) return;
  const targetSheet = getOrCreateHistorySheet();

  // Ensure Row 1 exists before inserting at Row 2 (right below the header)
  if (targetSheet.getLastRow() === 0) {
    targetSheet.getRange(1, 1, 1, 2).setValues([["Movie", "Watch Date"]]);
  }

  // Insert new row(s) at the top below the header (Row 2)
  targetSheet.insertRowsAfter(1, historyRows.length);
  targetSheet.getRange(2, 1, historyRows.length, 2).setValues(historyRows);
  targetSheet.getRange(2, 2, historyRows.length, 1).setNumberFormat("MMMM d, yyyy");

  // Always keep the History tab sorted newest watch date top-to-bottom
  sortHistorySheet(targetSheet);
}

function sortHistorySheet(historySheet) {
  if (!historySheet) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    for (const s of sheets) {
      if (String(s.getName() || "").toLowerCase().trim() === "history") {
        historySheet = s;
        break;
      }
    }
  }
  if (!historySheet) return;

  const lastRow = historySheet.getLastRow();
  if (lastRow <= 2) return; // Nothing to sort if only header + at most 1 row

  const numCols = Math.max(2, historySheet.getLastColumn());
  const numDataRows = lastRow - 1;
  const range = historySheet.getRange(2, 1, numDataRows, numCols);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const richTextA = historySheet.getRange(2, 1, numDataRows, 1).getRichTextValues();

  const rows = [];
  for (let r = 0; r < numDataRows; r++) {
    const hasContent = values[r].some(c => String(c || "").trim() !== "") || Boolean(formulas[r][0]);
    if (!hasContent) continue;

    const rowData = [];
    for (let c = 0; c < numCols; c++) {
      if (c === 0) {
        if (formulas[r][0]) {
          rowData.push(formulas[r][0]);
        } else {
          let linkUrl = richTextA[r][0] ? richTextA[r][0].getLinkUrl() : null;
          if (!linkUrl && richTextA[r][0] && richTextA[r][0].getRuns) {
            const runs = richTextA[r][0].getRuns();
            for (let k = 0; k < runs.length; k++) {
              if (runs[k].getLinkUrl()) {
                linkUrl = runs[k].getLinkUrl();
                break;
              }
            }
          }
          const cellText = String(values[r][0] || "").trim();
          if (linkUrl && cellText) {
            rowData.push(`=HYPERLINK("${linkUrl}", "${cellText.replace(/"/g, '""')}")`);
          } else {
            rowData.push(values[r][0]);
          }
        }
      } else {
        rowData.push(formulas[r][c] ? formulas[r][c] : values[r][c]);
      }
    }

    const rawDate = values[r][1];
    let parsedTime = -Infinity;
    if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
      parsedTime = rawDate.getTime();
    } else if (rawDate) {
      const d = parseSafeLocalDate(String(rawDate));
      if (d && !isNaN(d.getTime())) {
        parsedTime = d.getTime();
        rowData[1] = d;
      }
    }

    rows.push({
      data: rowData,
      time: parsedTime,
      origIdx: r
    });
  }

  if (rows.length === 0) return;

  // Sort newest watch date to oldest (top to bottom); preserve top-to-bottom order on ties
  rows.sort((a, b) => {
    if (a.time !== b.time) return b.time - a.time;
    return a.origIdx - b.origIdx;
  });

  const sortedValues = rows.map(r => r.data);
  historySheet.getRange(2, 1, sortedValues.length, numCols).setValues(sortedValues);
  historySheet.getRange(2, 2, sortedValues.length, 1).setNumberFormat("MMMM d, yyyy");

  if (sortedValues.length < numDataRows) {
    historySheet.getRange(2 + sortedValues.length, 1, numDataRows - sortedValues.length, numCols).clearContent();
  }
}

function moveRowToHistory(sheet, rowNum) {
  const rowRange = sheet.getRange(rowNum, 1, 1, 1);
  const valueA = rowRange.getValue();
  const formulaA = rowRange.getFormula();
  const richTextA = rowRange.getRichTextValue();

  const movieTitle = extractTitle(valueA, formulaA, richTextA) || "Movie";

  let titleCellToCopy = valueA;
  if (formulaA) {
    titleCellToCopy = formulaA;
  } else {
    const rtLink = richTextA ? richTextA.getLinkUrl() : null;
    const cellText = String(valueA || "").trim();
    if (rtLink && cellText) {
      titleCellToCopy = `=HYPERLINK("${rtLink}", "${cellText.replace(/"/g, '""')}")`;
    } else if (richTextA && richTextA.getRuns) {
      const runs = richTextA.getRuns();
      for (let i = 0; i < runs.length; i++) {
        const u = runs[i].getLinkUrl();
        if (u && cellText) {
          titleCellToCopy = `=HYPERLINK("${u}", "${cellText.replace(/"/g, '""')}")`;
          break;
        }
      }
    }
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  appendRowsToHistory([[titleCellToCopy, today]]);
  sheet.deleteRow(rowNum);

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Added "${movieTitle}" to the "History" tab with today's date.`,
      "✅ Completed",
      4
    );
  } catch (e) {}
}

function getOrCreateWontWatchSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (isWontWatchValue(s.getName())) {
      return s;
    }
  }
  return ss.insertSheet("Won't Watch");
}

function findBottomMovieRow(targetSheet, numCols) {
  const lastRow = targetSheet.getLastRow();
  if (lastRow === 0) return 0;
  const checkCols = Math.max(1, Math.min(numCols || 7, targetSheet.getLastColumn() || 1));
  const data = targetSheet.getRange(1, 1, lastRow, checkCols).getDisplayValues();
  for (let r = data.length - 1; r >= 0; r--) {
    if (data[r].some(cell => String(cell || "").trim() !== "")) {
      return r + 1;
    }
  }
  return 0;
}

function appendRowsToWontWatch(rowsArray) {
  if (!rowsArray || rowsArray.length === 0) return;
  const targetSheet = getOrCreateWontWatchSheet();
  const numCols = rowsArray[0].length;
  const bottomRow = findBottomMovieRow(targetSheet, numCols);
  const startRow = bottomRow + 1;

  targetSheet.getRange(startRow, 1, rowsArray.length, numCols).setValues(rowsArray);
}

function moveRowToWontWatch(sheet, rowNum) {
  const numCols = Math.max(7, sheet.getLastColumn());
  const rowRange = sheet.getRange(rowNum, 1, 1, numCols);
  const values = rowRange.getValues()[0];
  const formulas = rowRange.getFormulas()[0];
  const richText = sheet.getRange(rowNum, 1).getRichTextValue();

  const movieTitle = extractTitle(values[0], formulas[0], richText) || "Movie";

  const rowToMove = [];
  for (let c = 0; c < numCols; c++) {
    if (c === 0) {
      if (formulas[0]) {
        rowToMove.push(formulas[0]);
      } else {
        const rtLink = richText ? richText.getLinkUrl() : null;
        const cellText = String(values[0] || "").trim();
        if (rtLink && cellText) {
          rowToMove.push(`=HYPERLINK("${rtLink}", "${cellText.replace(/"/g, '""')}")`);
        } else {
          rowToMove.push(values[0]);
        }
      }
    } else {
      rowToMove.push(formulas[c] ? formulas[c] : values[c]);
    }
  }

  appendRowsToWontWatch([rowToMove]);
  sheet.deleteRow(rowNum);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Moved "${movieTitle}" to the bottom of the "Won't Watch" tab.`,
    "🚫 Moved to Won't Watch",
    4
  );
}

function checkAndFormatSheetWrapper() {
  processNewlyAddedRows();
  checkAndFormatSheet();
  sortHistorySheet();
}

function ensureHeaders(sheet) {
  const headerG = sheet.getRange(1, 7).getValue();
  if (headerG !== "Errors") {
    sheet.getRange(1, 7).setValue("Errors");
  }
}

function findNotesColumn(sheet) {
  const lastCol = Math.max(7, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c] || "").toLowerCase().trim();
    if (h === "notes" || h === "note") {
      return c + 1;
    }
  }
  return 6; // Default to Column F (Col 6) if no "Notes" header is found
}

function isNetflixNote(noteText) {
  return /\bnetflix\b/i.test(String(noteText || ""));
}

function isTimeTravelText(text) {
  if (!text) return false;
  return /\b(time[\s-]+travel(?:ler|er|ing|led|s)?|time[\s-]+loop(?:s)?|time[\s-]+machine(?:s)?|time[\s-]+portal(?:s)?|time[\s-]+vortex|time[\s-]+warp|time[\s-]+paradox(?:es)?|temporal\s+(?:paradox|loop|anomaly|displacement|inversion|agent)|time\s+inversion|travel(?:s|ed|ing|led)?\s+(?:back|forward)\s+in\s+time|travel(?:s|ed|ing|led)?\s+through\s+time|sent\s+back\s+in\s+time|go(?:es|ing)?\s+back\s+in\s+time|reliv(?:e|es|ing)\s+the\s+same\s+day|stuck\s+in\s+a\s+time\s+loop)\b/i.test(String(text));
}

function isMovieTimeTravel(opts) {
  if (!opts) return false;
  if (String(opts.currentCategory || "").toLowerCase().trim() === "time travel") return true;
  if (isTimeTravelText(opts.plotText) || isTimeTravelText(opts.synopsisText) || isTimeTravelText(opts.noteText)) return true;
  if (opts.imdbId) {
    const cleanId = String(opts.imdbId).trim().toLowerCase();
    const map = fetchTimeTravelImdbSet([cleanId]);
    if (map && map[cleanId]) return true;
  }
  return false;
}

/**
 * Given an array of IMDb IDs (e.g. ["tt0088763", "tt0816692"]), returns a map `{ [imdbId]: true }`
 * for every movie that involves time travel in any way (via Wikidata P136/P921 or English Wikipedia
 * time-travel / time-loop / temporal-paradox film categories). Results are cached in ScriptCache.
 */
function fetchTimeTravelImdbSet(imdbIds) {
  const result = {};
  if (!imdbIds || imdbIds.length === 0) return result;

  const uniqueIds = [];
  const seen = {};
  for (let i = 0; i < imdbIds.length; i++) {
    const id = String(imdbIds[i] || "").trim().toLowerCase();
    if (/^tt\d+$/.test(id) && !seen[id]) {
      seen[id] = true;
      uniqueIds.push(id);
    }
  }
  if (uniqueIds.length === 0) return result;

  let cache = null;
  try {
    cache = CacheService.getScriptCache();
  } catch (e) {}

  const uncachedIds = [];
  for (let i = 0; i < uniqueIds.length; i++) {
    const id = uniqueIds[i];
    const cachedVal = cache ? cache.get(`tt_cat_${id}`) : null;
    if (cachedVal === "1") {
      result[id] = true;
    } else if (cachedVal === null) {
      uncachedIds.push(id);
    }
  }

  if (uncachedIds.length === 0) return result;

  const SPARQL_BATCH = 60;
  const wikiTitleToImdb = {};
  const wikiTitles = [];

  for (let b = 0; b < uncachedIds.length; b += SPARQL_BATCH) {
    const batch = uncachedIds.slice(b, b + SPARQL_BATCH);
    const valuesClause = batch.map(id => `"${id}"`).join(" ");
    const sparql = `SELECT ?imdb ?isTT ?articleTitle WHERE {
      VALUES ?imdb { ${valuesClause} }
      ?item wdt:P345 ?imdb .
      BIND(EXISTS {
        ?item (wdt:P136|wdt:P921) ?topic .
        VALUES ?topic { wd:Q182154 wd:Q2750812 wd:Q1047698 }
      } AS ?isTT)
      OPTIONAL {
        ?article schema:about ?item ;
                 schema:isPartOf <https://en.wikipedia.org/> ;
                 schema:name ?articleTitle .
      }
    }`;

    try {
      const wdUrl = `https://query.wikidata.org/sparql?format=json&query=${encodeURIComponent(sparql)}`;
      const wdRes = UrlFetchApp.fetch(wdUrl, {
        muteHttpExceptions: true,
        headers: {
          "User-Agent": "MovieTrackerSpreadsheet/1.0 (GoogleAppsScript)",
          "Accept": "application/sparql-results+json"
        }
      });

      if (wdRes && wdRes.getResponseCode() === 200) {
        const wdJson = JSON.parse(wdRes.getContentText());
        const bindings = (wdJson && wdJson.results && wdJson.results.bindings) || [];
        for (let j = 0; j < bindings.length; j++) {
          const row = bindings[j];
          const id = row.imdb && row.imdb.value ? String(row.imdb.value).toLowerCase() : null;
          if (!id) continue;
          const isTT = row.isTT && row.isTT.value === "true";
          if (isTT) {
            result[id] = true;
          } else if (row.articleTitle && row.articleTitle.value) {
            const wTitle = row.articleTitle.value;
            if (!wikiTitleToImdb[wTitle]) {
              wikiTitleToImdb[wTitle] = id;
              wikiTitles.push(wTitle);
            }
          }
        }
      }
    } catch (e) {}
  }

  // Check English Wikipedia categories in parallel batches of 5 titles (prevents cllimit=500 truncation)
  if (wikiTitles.length > 0) {
    const WIKI_BATCH = 5;
    const wikiReqs = [];
    for (let i = 0; i < wikiTitles.length; i += WIKI_BATCH) {
      const chunk = wikiTitles.slice(i, i + WIKI_BATCH);
      const url = `https://en.wikipedia.org/w/api.php?action=query&prop=categories&cllimit=max&format=json&titles=${encodeURIComponent(chunk.join("|"))}`;
      wikiReqs.push({
        url: url,
        muteHttpExceptions: true,
        headers: { "User-Agent": "MovieTrackerSpreadsheet/1.0 (GoogleAppsScript)" }
      });
    }

    try {
      const wikiResps = UrlFetchApp.fetchAll(wikiReqs);
      const ttCatRegex = /\b(films about time travel|time travel films|films about time loops|time loop films|fiction about temporal paradoxes|time-travel films)\b/i;

      for (let r = 0; r < wikiResps.length; r++) {
        const res = wikiResps[r];
        if (!res || res.getResponseCode() !== 200) continue;
        const wJson = JSON.parse(res.getContentText());
        const pages = (wJson && wJson.query && wJson.query.pages) || {};
        const pageKeys = Object.keys(pages);
        for (let k = 0; k < pageKeys.length; k++) {
          const p = pages[pageKeys[k]];
          const id = p && p.title ? wikiTitleToImdb[p.title] : null;
          if (!id || !p.categories) continue;
          for (let c = 0; c < p.categories.length; c++) {
            const catTitle = String(p.categories[c].title || "");
            if (ttCatRegex.test(catTitle)) {
              result[id] = true;
              break;
            }
          }
        }
      }
    } catch (e) {}
  }

  if (cache) {
    const toCache = {};
    for (let i = 0; i < uncachedIds.length; i++) {
      const id = uncachedIds[i];
      toCache[`tt_cat_${id}`] = result[id] ? "1" : "0";
    }
    try {
      cache.putAll(toCache, 21600);
    } catch (e) {}
  }

  return result;
}

function isDec31Date(dateObj) {
  return Boolean(
    dateObj instanceof Date &&
    !isNaN(dateObj.getTime()) &&
    dateObj.getMonth() === 11 &&
    dateObj.getDate() === 31
  );
}

function isUnannouncedDateNote(noteText) {
  return /\brelease\s+date\s+unannounced\b/i.test(String(noteText || ""));
}

function upsertUnannouncedDateNote(existingNote) {
  const tag = "Release Date Unannounced";
  const raw = String(existingNote || "").trim();
  if (!raw) return tag;
  const parts = raw
    .split("|")
    .map(p => p.trim())
    .filter(p => p && !/^release\s+date\s+unannounced$/i.test(p));
  if (parts.length === 0) return tag;
  return `${parts.join(" | ")} | ${tag}`;
}

function removeUnannouncedDateNote(existingNote) {
  const raw = String(existingNote || "").trim();
  if (!raw) return "";
  const parts = raw
    .split("|")
    .map(p => p.trim())
    .filter(p => p && !/^release\s+date\s+unannounced$/i.test(p));
  return parts.join(" | ");
}

function upsertNetflixNote(existingNote) {
  const raw = String(existingNote || "").trim();
  if (!raw) return "Netflix";
  const parts = raw
    .split("|")
    .map(p => p.trim())
    .filter(p => p && !/^VOD:\s*/i.test(p) && !/^netflix$/i.test(p));
  if (parts.length === 0) return "Netflix";
  return `${parts.join(" | ")} | Netflix`;
}

function upsertVodNote(existingNote, vodDateStr) {
  if (isNetflixNote(existingNote)) {
    return upsertNetflixNote(existingNote);
  }
  const tag = `VOD: ${vodDateStr}`;
  if (!existingNote) return tag;
  if (/VOD:\s*[^|\n]+/i.test(existingNote)) {
    return existingNote.replace(/VOD:\s*[^|\n]+/i, tag);
  }
  return `${existingNote} | ${tag}`;
}

function extractVodDateFromNote(noteText) {
  if (!noteText) return null;
  const match = String(noteText).match(/VOD:\s*([^|\n]+)/i);
  if (!match) return null;
  return parseSafeLocalDate(match[1].trim());
}

function getSixMonthsAgo(today) {
  const d = new Date(today.getTime());
  d.setDate(1);
  d.setMonth(d.getMonth() - 6);
  const maxDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(today.getDate(), maxDay));
  return d;
}

function extractTitle(displayVal, formulaVal, richTextVal) {
  if (formulaVal && formulaVal.toString().toUpperCase().startsWith("=HYPERLINK")) {
    const match = formulaVal.toString().match(/=HYPERLINK\("[^"]*"\s*[,;]\s*"([\s\S]*)"\)/i);
    if (match) return match[1].replace(/""/g, '"').trim();
  }
  if (displayVal !== undefined && displayVal !== null && String(displayVal).trim() !== "") {
    return String(displayVal).trim();
  }
  if (richTextVal && richTextVal.getText()) {
    return String(richTextVal.getText()).trim();
  }
  return "";
}

function extractImdbId(formulaVal, richTextVal) {
  if (formulaVal) {
    const match = formulaVal.toString().match(/imdb\.com\/title\/(tt\d+)/i);
    if (match) return match[1].toLowerCase();
  }
  if (richTextVal) {
    const directUrl = richTextVal.getLinkUrl();
    if (directUrl) {
      const m = directUrl.match(/imdb\.com\/title\/(tt\d+)/i);
      if (m) return m[1].toLowerCase();
    }
    const runs = richTextVal.getRuns ? richTextVal.getRuns() : [];
    for (let i = 0; i < runs.length; i++) {
      const u = runs[i].getLinkUrl();
      if (u) {
        const m = u.match(/imdb\.com\/title\/(tt\d+)/i);
        if (m) return m[1].toLowerCase();
      }
    }
  }
  return null;
}

function sanitizeTitleAndYear(rawTitle) {
  const str = String(rawTitle || "");
  const yearMatch = str.match(/\((\d{4})\)\s*$/);
  const year = yearMatch ? yearMatch[1] : null;

  const cleanTitle = str
    .replace(/\s*\([^)]*\)\s*$/g, "")
    .replace(/\s+season\s+\d+/i, "")
    .trim();

  return { cleanTitle: cleanTitle || str.trim(), year: year };
}

function normalizeForCompare(str) {
  return String(str || "")
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]/g, "");
}

function parseSafeLocalDate(dateStr) {
  if (!dateStr || dateStr === "N/A") return null;
  const s = String(dateStr).trim();
  const yearOnly = s.match(/^(\d{4})$/);
  if (yearOnly) {
    const y = parseInt(yearOnly[1], 10);
    const todayY = (new Date()).getFullYear();
    return y >= todayY ? new Date(y, 11, 31) : new Date(y, 0, 1);
  }
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
  }
  const d = new Date(s);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Normalizes, sanitizes, and recovers date values from Google Sheets cells.
 * Specifically fixes:
 * 1. 1905 serial date bug: In Google Sheets, entering a 4-digit year like 2026 or 2027
 *    into a date-formatted cell converts the number into day #2026 or #2027 since 1899-12-30,
 *    resulting in July 18, 1905 or July 19, 1905.
 * 2. 4-digit year strings or numbers (e.g. "2026", 2026) -> local Date Dec 31 (future) or Jan 1 (past).
 * 3. Standard Date objects -> returns valid Date or null if invalid.
 */
function sanitizeAndRecoverDate(val, fallbackYear) {
  const todayY = (new Date()).getFullYear();

  if (val === null || val === undefined || val === "") {
    if (fallbackYear) {
      const ym = String(fallbackYear).match(/\d{4}/);
      if (ym) {
        const y = parseInt(ym[0], 10);
        if (y >= 1900 && y <= 2150) {
          return { date: (y >= todayY ? new Date(y, 11, 31) : new Date(y, 0, 1)), isYearOnly: true, year: y };
        }
      }
    }
    return null;
  }

  // If number like 2026
  if (typeof val === "number" && !isNaN(val)) {
    if (val >= 1900 && val <= 2150) {
      const y = Math.floor(val);
      return { date: (y >= todayY ? new Date(y, 11, 31) : new Date(y, 0, 1)), isYearOnly: true, year: y };
    }
  }

  // If string
  if (typeof val === "string") {
    const trimmed = val.trim();
    if (!trimmed || trimmed === "N/A") return null;
    const yearMatch = trimmed.match(/^(\d{4})$/);
    if (yearMatch) {
      const y = parseInt(yearMatch[1], 10);
      if (y >= 1900 && y <= 2150) {
        return { date: (y >= todayY ? new Date(y, 11, 31) : new Date(y, 0, 1)), isYearOnly: true, year: y };
      }
    }
    const parsed = parseSafeLocalDate(trimmed);
    if (parsed && !isNaN(parsed.getTime())) {
      val = parsed;
    }
  }

  if (val instanceof Date && !isNaN(val.getTime())) {
    // Detect the 1905 Google Sheets serial-number conversion
    if (val.getFullYear() === 1905) {
      // If user/API explicitly specified a fallback year (e.g. 2026 for Babies), prioritize that
      if (fallbackYear) {
        const ym = String(fallbackYear).match(/\d{4}/);
        if (ym) {
          const y = parseInt(ym[0], 10);
          if (y >= 1900 && y <= 2150) {
            return { date: (y >= todayY ? new Date(y, 11, 31) : new Date(y, 0, 1)), isYearOnly: true, year: y };
          }
        }
      }
      const epoch = new Date(1899, 11, 30);
      const serialDays = Math.round((val.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000));
      if (serialDays >= 1920 && serialDays <= 2150) {
        return { date: (serialDays >= todayY ? new Date(serialDays, 11, 31) : new Date(serialDays, 0, 1)), isYearOnly: true, year: serialDays };
      }
    }

    return { date: val, isYearOnly: false, year: val.getFullYear() };
  }

  return null;
}

function decodeHtmlEntities(str) {
  return String(str || "")
    .replace(/&#039;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// ============================================================================
// 10B. ⚡ QUICK PRIORITIZER (HTML interface in PriorityPicker.html)
//      - Steps through every movie with a blank Priority cell
//      - One-click (or number-key) priority buttons, auto-advances to the next movie
//      - Shows poster, RT Popcornmeter score, and a one-click YouTube trailer search
// ============================================================================
function openPriorityPicker() {
  const html = HtmlService.createHtmlOutputFromFile("PriorityPicker")
    .setWidth(980)
    .setHeight(660);
  SpreadsheetApp.getUi().showModalDialog(html, "⚡ Quick Prioritizer");
}

function getPriorityOptions(sheet) {
  const col = findPriorityColumn(sheet);
  let opts = [];
  try {
    const rule = sheet.getRange(2, col).getDataValidation();
    if (rule) {
      const type = rule.getCriteriaType();
      const args = rule.getCriteriaValues();
      if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        opts = (Array.isArray(args[0]) ? args[0] : []).map(String);
      } else if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && args[0]) {
        opts = args[0].getDisplayValues().reduce((acc, r) => acc.concat(r), []).map(String);
      }
    }
  } catch (e) {}

  opts = opts.map(s => s.trim()).filter(s => s !== "");
  opts = opts.filter((s, i) => opts.indexOf(s) === i); // de-dupe
  if (opts.length === 0) opts = ["High", "Medium", "Low"];
  if (!opts.some(o => isCompletedValue(o))) opts.push("Completed");
  if (!opts.some(o => isWontWatchValue(o))) opts.push("Won't Watch");

  // Keep "Completed" and "Won't Watch" as the last options
  const completed = opts.filter(o => isCompletedValue(o));
  const wont = opts.filter(o => isWontWatchValue(o));
  return opts
    .filter(o => !isCompletedValue(o) && !isWontWatchValue(o))
    .concat(completed.slice(0, 1))
    .concat(wont.slice(0, 1));
}

// Called once when the prioritizer opens: returns the priority options + every un-prioritized movie
function getPriorityQueue() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) throw new Error("Could not find a sheet named 'Movies'.");

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  const options = getPriorityOptions(sheet);
  if (lastRow <= 1) return { options: options, movies: [] };

  const numCols = Math.max(7, sheet.getLastColumn());
  const range = sheet.getRange(1, 1, lastRow, numCols);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const formulas = range.getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();
  const priorityIdx = findPriorityColumn(sheet) - 1;
  const notesIdx = findNotesColumn(sheet) - 1;

  const movies = [];
  for (let i = 1; i < values.length; i++) {
    const title = extractTitle(display[i][0], formulas[i][0], richText[i][0]);
    if (!title) continue;
    if (String(display[i][priorityIdx] || "").trim() !== "") continue; // Already prioritized

    const { cleanTitle, year } = sanitizeTitleAndYear(title);
    const recDate = sanitizeAndRecoverDate(values[i][2], year);
    const resolvedYear = recDate ? recDate.year : (year ? parseInt(year, 10) : null);
    const rtScore = String(display[i][3] || "").trim();

    let formattedDateStr = String(display[i][2] || "").trim();
    if (recDate && (recDate.isYearOnly || values[i][2] instanceof Date && values[i][2].getFullYear() === 1905)) {
      formattedDateStr = String(recDate.year);
    }

    movies.push({
      row: i + 1,
      title: title,
      cleanTitle: cleanTitle,
      year: resolvedYear,
      imdbId: extractImdbId(formulas[i][0], richText[i][0]),
      category: String(display[i][1] || "").trim(),
      dateStr: formattedDateStr,
      rtScore: rtScore && rtScore !== "N/A" ? rtScore : "",
      rtUrl: extractRtUrlFromFormula(formulas[i][3]),
      notes: String(display[i][notesIdx] || "").trim()
    });
  }

  return { options: options, movies: movies };
}

// Fetches poster + cast + plot summary for one movie (cached for 6 hours so revisits are instant)
function getMovieMediaForPicker(imdbId, title, year, rtUrl) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "pp_v2_" + (imdbId || (normalizeForCompare(title) + "_" + (year || ""))).substring(0, 200);
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      // Only return cached entry if it already has a plot or if we explicitly cached an empty plot
      if (parsed && parsed.plot) return parsed;
    } catch (e) {}
  }

  const result = { poster: "", stars: "", imdbId: imdbId || "", plot: "" };

  // 1. Poster + cast from IMDb (by exact IMDb ID, or best title match if the row isn't linked yet)
  let details = imdbId ? fetchImdbDetailsById(imdbId) : null;
  if (!details || !details.poster) {
    const cands = fetchMovieCandidates(sanitizeTitleAndYear(title).cleanTitle, 8);
    const pick = (imdbId && cands.find(c => c.imdbID === imdbId)) ||
      cands.find(c => normalizeForCompare(c.title) === normalizeForCompare(sanitizeTitleAndYear(title).cleanTitle) &&
                      (!year || !c.year || Math.abs(c.year - year) <= 1)) ||
      (!imdbId ? cands[0] : null);
    if (pick) {
      details = Object.assign({}, details || {}, {
        poster: pick.poster || (details && details.poster) || "",
        stars: (details && details.stars) || pick.stars || ""
      });
      if (!result.imdbId) result.imdbId = pick.imdbID;
    }
  }
  if (details) {
    // Request a smaller ~400px-wide poster from IMDb's image CDN
    result.poster = details.poster ? details.poster.replace(/\._V1_.*\.jpg$/i, "._V1_SX400_.jpg") : "";
    result.stars = details.stars || "";
  }

  // 2. Direct Rotten Tomatoes lookup if rtUrl is already known on this row
  if (rtUrl && !result.plot) {
    try {
      const pageRes = UrlFetchApp.fetch(rtUrl, {
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)" }
      });
      if (pageRes.getResponseCode() === 200) {
        const parsed = parseRtMoviePageHtml(pageRes.getContentText(), { url: rtUrl });
        if (parsed && parsed.synopsis) {
          result.plot = parsed.synopsis;
        }
      }
    } catch (e) {}
  }

  // 3. Fallback to Rotten Tomatoes search if rtUrl wasn't provided or didn't yield a synopsis
  if (!result.plot) {
    try {
      const cleanTitle = sanitizeTitleAndYear(title).cleanTitle;
      const rtData = fetchRottenTomatoesData(cleanTitle, year);
      if (rtData && rtData.synopsis) {
        result.plot = rtData.synopsis;
      }
    } catch (e) {}
  }

  // 4. Fallback to OMDb API
  if (!result.plot && !omdbQuotaExceeded) {
    let omdbQuery = "";
    if (result.imdbId) {
      omdbQuery = `https://www.omdbapi.com/?i=${result.imdbId}&apikey=${apiKey}`;
    } else {
      omdbQuery = `https://www.omdbapi.com/?t=${encodeURIComponent(sanitizeTitleAndYear(title).cleanTitle)}&type=movie&apikey=${apiKey}`;
      if (year) omdbQuery += `&y=${year}`;
    }
    const omdbData = fetchOmdbJson(omdbQuery);
    if (omdbData && omdbData.Plot && omdbData.Plot !== "N/A") {
      result.plot = omdbData.Plot;
    }
  }

  try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch (e) {}
  return result;
}

// Saves the chosen priority. Verifies the movie's live row by title first, so row shifts never mis-assign.
function savePriorityFromPicker(expectedRow, title, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) throw new Error("Could not find a sheet named 'Movies'.");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const liveRow = findCurrentRowForMovie(sheet, expectedRow, title);
    if (!liveRow) throw new Error(`Couldn't find "${title}" on the Movies tab (it may have been moved or deleted).`);

    const priorityCol = findPriorityColumn(sheet);
    sheet.getRange(liveRow, priorityCol).setValue(value || "");

    if (value && isCompletedValue(value)) {
      moveRowToHistory(sheet, liveRow);
      return { moved: true };
    }
    if (value && isWontWatchValue(value)) {
      moveRowToWontWatch(sheet, liveRow);
      return { moved: true };
    }
    return { moved: false, row: liveRow };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 10C. 🎲 RANDOM HIGH-PRIORITY MOVIE SELECTOR
//      - Randomly picks a movie with Priority "High"
//      - Excludes categories "To be released" and "In Theaters"
//      - Displays modal preview with poster, RT score, synopsis & YouTube search
// ============================================================================
function openRandomHighPriorityMovie() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Could not find a sheet named 'Movies'.");
    return;
  }

  ensureHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("No movies found on the 'Movies' sheet.");
    return;
  }

  const numCols = Math.max(7, sheet.getLastColumn());
  const range = sheet.getRange(1, 1, lastRow, numCols);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const formulas = range.getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();
  const priorityIdx = findPriorityColumn(sheet) - 1;
  const notesIdx = findNotesColumn(sheet) - 1;

  const eligible = [];
  for (let i = 1; i < values.length; i++) {
    const title = extractTitle(display[i][0], formulas[i][0], richText[i][0]);
    if (!title) continue;

    const prio = String(display[i][priorityIdx] || "").trim().toLowerCase();
    if (!prio.includes("high")) continue;

    const cat = String(display[i][1] || "").trim().toLowerCase();
    if (cat === "to be released" || cat === "in theaters") continue;

    const { cleanTitle, year } = sanitizeTitleAndYear(title);
    const recDate = sanitizeAndRecoverDate(values[i][2], year);
    const resolvedYear = recDate ? recDate.year : (year ? parseInt(year, 10) : null);
    const rtScore = String(display[i][3] || "").trim();

    let formattedDateStr = String(display[i][2] || "").trim();
    if (recDate && (recDate.isYearOnly || values[i][2] instanceof Date && values[i][2].getFullYear() === 1905)) {
      formattedDateStr = String(recDate.year);
    }

    eligible.push({
      row: i + 1,
      title: title,
      cleanTitle: cleanTitle,
      year: resolvedYear,
      imdbId: extractImdbId(formulas[i][0], richText[i][0]),
      category: String(display[i][1] || "").trim(),
      dateStr: formattedDateStr,
      rtScore: rtScore && rtScore !== "N/A" ? rtScore : "",
      rtUrl: extractRtUrlFromFormula(formulas[i][3]),
      notes: String(display[i][notesIdx] || "").trim()
    });
  }

  if (eligible.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No eligible High Priority movies found!\n\n" +
      "Make sure you have movies marked 'High' that are not in 'To be released' or 'In Theaters'."
    );
    return;
  }

  const html = HtmlService.createHtmlOutputFromFile("RandomMoviePicker")
    .setWidth(850)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "🎲 Random Movie Selector");
}

function getRandomHighPriorityPool() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) throw new Error("Could not find 'Movies' sheet.");

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const numCols = Math.max(7, sheet.getLastColumn());
  const range = sheet.getRange(1, 1, lastRow, numCols);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const formulas = range.getFormulas();
  const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();
  const priorityIdx = findPriorityColumn(sheet) - 1;
  const notesIdx = findNotesColumn(sheet) - 1;

  const eligible = [];
  for (let i = 1; i < values.length; i++) {
    const title = extractTitle(display[i][0], formulas[i][0], richText[i][0]);
    if (!title) continue;

    const prio = String(display[i][priorityIdx] || "").trim().toLowerCase();
    if (!prio.includes("high")) continue;

    const cat = String(display[i][1] || "").trim().toLowerCase();
    if (cat === "to be released" || cat === "in theaters") continue;

    const { cleanTitle, year } = sanitizeTitleAndYear(title);
    const recDate = sanitizeAndRecoverDate(values[i][2], year);
    const resolvedYear = recDate ? recDate.year : (year ? parseInt(year, 10) : null);
    const rtScore = String(display[i][3] || "").trim();

    let formattedDateStr = String(display[i][2] || "").trim();
    if (recDate && (recDate.isYearOnly || values[i][2] instanceof Date && values[i][2].getFullYear() === 1905)) {
      formattedDateStr = String(recDate.year);
    }

    eligible.push({
      row: i + 1,
      title: title,
      cleanTitle: cleanTitle,
      year: resolvedYear,
      imdbId: extractImdbId(formulas[i][0], richText[i][0]),
      category: String(display[i][1] || "").trim(),
      dateStr: formattedDateStr,
      rtScore: rtScore && rtScore !== "N/A" ? rtScore : "",
      rtUrl: extractRtUrlFromFormula(formulas[i][3]),
      notes: String(display[i][notesIdx] || "").trim()
    });
  }

  return eligible;
}

// ============================================================================
// 11. CUSTOM MENU
// ============================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🎬 Movie Tools")
    .addItem("🎲 Pick Random High Priority Movie", "openRandomHighPriorityMovie")
    .addSeparator()
    .addItem("➕ Add New Movie", "addNewMovieDialog")
    .addItem("⚡ Quick Prioritize Movies", "openPriorityPicker")
    .addItem("🎯 Fix Selected Row (Choose Correct Movie)", "fixSelectedRowDialog")
    .addSeparator()
    .addItem("🍿 Update Missing RT Scores & Links", "updateMissingRtScoresAndLinks")
    .addItem("🔍 Update Missing Dates, Scores & Links", "updateMissingDataOnly")
    .addItem("🔗 Link Unlinked Movies Only", "linkUnlinkedMoviesOnly")
    .addItem("📺 Check VOD Dates & Promote Categories Now", "dailyVodCheckMenuWrapper")
    .addItem("🔄 Refresh In Theaters & New Releases", "weeklyActiveMoviesRefresh")
    .addSeparator()
    .addItem("🎨 Sort & Format Sheet Now", "checkAndFormatSheetWrapper")
    .addToUi();
}


// ============================================================================
// 12. 🔔 DISCORD "NOW AVAILABLE" NOTIFICATIONS
//     - Posts "Now available: <Title>" (title hyperlinked to IMDb) to a Discord webhook
//       whenever a movie moves from "To be released" / "In Theaters" into "New Releases".
//     - The webhook URL is stored in Script Properties as "DISCORD_WEBHOOK_URL" (never in code).
//     - Tracks which movies were "To be released"/"In Theaters" at the last check in the
//       "NR_PREV_ACTIVE" Script Property, so each transition is announced exactly once and
//       movies added directly as "New Releases" (or existing ones) never spam the channel.
// ============================================================================
function getDiscordWebhookUrl() {
  try {
    return PropertiesService.getScriptProperties().getProperty("DISCORD_WEBHOOK_URL") || "";
  } catch (e) {
    return "";
  }
}

function getMovieNotifyKey(colAValue) {
  const imdbId = extractImdbId(String(colAValue || ""), null);
  if (imdbId) return imdbId;
  const title = extractTitle(colAValue, String(colAValue || ""), null);
  const norm = normalizeForCompare(sanitizeTitleAndYear(title).cleanTitle);
  return norm ? "t:" + norm : "";
}

function processNewReleaseNotifications(sortedValues) {
  if (!sortedValues || sortedValues.length === 0) return;

  {
    const props = PropertiesService.getScriptProperties();
    const prevRaw = props.getProperty("NR_PREV_ACTIVE");
    const isFirstRun = prevRaw === null;
    let prevActive = {};
    if (!isFirstRun) {
      try { prevActive = JSON.parse(prevRaw) || {}; } catch (e) { prevActive = {}; }
    }

    const webhookUrl = getDiscordWebhookUrl();
    const nextActive = {};
    const toNotify = [];

    for (let i = 0; i < sortedValues.length; i++) {
      const colA = sortedValues[i][0];
      const cat = String(sortedValues[i][1] || "").toLowerCase().trim();
      const key = getMovieNotifyKey(colA);
      if (!key) continue;

      if (cat === "to be released" || cat === "in theaters") {
        nextActive[key] = 1;
      } else if (cat === "new releases" && !isFirstRun && prevActive[key]) {
        const title = extractTitle(colA, String(colA || ""), null);
        const imdbId = extractImdbId(String(colA || ""), null);
        toNotify.push({ key: key, title: title, imdbId: imdbId });
      }
    }

    for (let n = 0; n < toNotify.length; n++) {
      const item = toNotify[n];
      const sent = webhookUrl ? postDiscordNowAvailable(webhookUrl, item.title, item.imdbId) : false;
      // If posting failed (or no webhook configured yet), keep it pending so it retries next run
      if (!sent) nextActive[item.key] = 1;
    }

    props.setProperty("NR_PREV_ACTIVE", JSON.stringify(nextActive));
  }
}

function postDiscordNowAvailable(webhookUrl, title, imdbId) {
  const safeTitle = String(title || "Movie").replace(/([\[\]\\*_~`|])/g, "\\$1");
  const imdbUrl = imdbId
    ? `https://www.imdb.com/title/${imdbId}/`
    : `https://www.imdb.com/find/?q=${encodeURIComponent(title || "")}`;
  const payload = {
    content: `🍿 **Now available:** [${safeTitle}](${imdbUrl})`,
    allowed_mentions: { parse: [] }
  };

  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const res = UrlFetchApp.fetch(webhookUrl, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code >= 200 && code < 300) return true;
      if (code === 429) {
        let waitMs = 1500;
        try {
          const body = JSON.parse(res.getContentText());
          if (body && body.retry_after) waitMs = Math.ceil(Number(body.retry_after) * 1000) + 250;
        } catch (e) {}
        Utilities.sleep(Math.min(waitMs, 10000));
        continue;
      }
      return false;
    } catch (e) {
      Utilities.sleep(1000);
    }
  }
  return false;
}


// ============================================================================
// 13. 🔐 PLEX WEBHOOK RECEIVER (Marks Watched Movies as "Completed")
//     - Publish via Apps Script: Deploy -> New deployment -> Web app
//       * Execute as: Me
//       * Who has access: Anyone
//     - Security layers (all stored in Script Properties, never in code):
//       1. PLEX_WEBHOOK_TOKEN (required): high-entropy secret passed via ?token=...
//          Verified with constant-time comparison; fails closed if unset.
//       2. PLEX_SERVER_UUID (optional): restricts requests to your Plex server UUID.
//       3. PLEX_ACCOUNT_TITLE (optional): restricts to a specific Plex username.
//     - Only acts on event === "media.scrobble" and Metadata.type === "movie".
//     - Matches by IMDb ID first (from Metadata.Guid / Metadata.guid), then Title+Year.
//     - Marks Priority = "Completed" and moves row to "History" (Col A = IMDb link, Col B = Date).
// ============================================================================
function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const expectedToken = props.getProperty("PLEX_WEBHOOK_TOKEN") || "";

    // Fail closed if no secret token has been configured in Script Properties
    if (!expectedToken || expectedToken.length < 16) {
      return jsonTextResponse({ ok: false });
    }

    const providedToken = (e && e.parameter && e.parameter.token) ? String(e.parameter.token).trim() : "";
    if (!constantTimeEquals(providedToken, expectedToken)) {
      return jsonTextResponse({ ok: false });
    }

    const payload = parsePlexWebhookPayload(e);
    if (!payload || typeof payload !== "object") {
      return jsonTextResponse({ ok: false });
    }

    // Optional defense-in-depth: verify Plex Server UUID if configured
    const expectedServerUuid = (props.getProperty("PLEX_SERVER_UUID") || "").trim();
    if (expectedServerUuid) {
      const actualUuid = (payload.Server && payload.Server.uuid) ? String(payload.Server.uuid).trim() : "";
      if (!constantTimeEquals(actualUuid, expectedServerUuid)) {
        return jsonTextResponse({ ok: false });
      }
    }

    // Optional defense-in-depth: restrict to a specific Plex user account (e.g. server owner)
    const expectedAccount = (props.getProperty("PLEX_ACCOUNT_TITLE") || "").trim().toLowerCase();
    if (expectedAccount) {
      const actualAccount = (payload.Account && payload.Account.title) ? String(payload.Account.title).trim().toLowerCase() : "";
      if (actualAccount !== expectedAccount) {
        return jsonTextResponse({ ok: true, ignored: true });
      }
    }

    // Only act when Plex marks a movie as watched/scrobbled
    const eventName = String(payload.event || "").trim().toLowerCase();
    const meta = payload.Metadata || {};
    const mediaType = String(meta.type || "").trim().toLowerCase();

    if (eventName !== "media.scrobble" || mediaType !== "movie") {
      return jsonTextResponse({ ok: true, ignored: true });
    }

    const plexImdbId = extractPlexImdbId(meta);
    const plexTitle = String(meta.title || "").trim();
    const plexOrigTitle = String(meta.originalTitle || "").trim();
    const plexYear = meta.year ? parseInt(meta.year, 10) : null;

    if (!plexImdbId && !plexTitle) {
      return jsonTextResponse({ ok: false });
    }

    const result = markMovieCompletedFromPlex(plexImdbId, plexTitle, plexOrigTitle, plexYear);
    return jsonTextResponse({ ok: true, matched: result.matched });
  } catch (err) {
    return jsonTextResponse({ ok: false });
  }
}

function jsonTextResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function constantTimeEquals(a, b) {
  const strA = String(a || "");
  const strB = String(b || "");
  let diff = strA.length ^ strB.length;
  const maxLen = Math.max(strA.length, strB.length);
  for (let i = 0; i < maxLen; i++) {
    const codeA = i < strA.length ? strA.charCodeAt(i) : 0;
    const codeB = i < strB.length ? strB.charCodeAt(i) : 0;
    diff |= (codeA ^ codeB);
  }
  return diff === 0 && strA.length > 0;
}

function parsePlexWebhookPayload(e) {
  if (!e) return null;
  // Plex sends multipart/form-data with a "payload" form field
  if (e.parameter && e.parameter.payload) {
    try {
      return JSON.parse(e.parameter.payload);
    } catch (err) {}
  }
  // Fallback for application/json POST bodies
  if (e.postData && e.postData.contents) {
    const raw = String(e.postData.contents).trim();
    if (raw.startsWith("{")) {
      try {
        return JSON.parse(raw);
      } catch (err) {}
    }
  }
  return null;
}

function extractPlexImdbId(meta) {
  if (!meta) return null;
  if (Array.isArray(meta.Guid)) {
    for (let i = 0; i < meta.Guid.length; i++) {
      const idStr = String((meta.Guid[i] && meta.Guid[i].id) || "");
      const m = idStr.match(/imdb:\/\/(tt\d+)/i) || idStr.match(/\b(tt\d+)\b/i);
      if (m) return m[1].toLowerCase();
    }
  }
  if (meta.guid) {
    const m = String(meta.guid).match(/imdb:\/\/(tt\d+)/i) || String(meta.guid).match(/\b(tt\d+)\b/i);
    if (m) return m[1].toLowerCase();
  }
  return null;
}

function markMovieCompletedFromPlex(plexImdbId, plexTitle, plexOrigTitle, plexYear) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Movies");
  if (!sheet) return { matched: false };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { matched: false };

    const values = sheet.getRange(1, 1, lastRow, 3).getValues();
    const formulas = sheet.getRange(1, 1, lastRow, 1).getFormulas();
    const richText = sheet.getRange(1, 1, lastRow, 1).getRichTextValues();

    const matchedRow = findMovieRowForPlex(values, formulas, richText, plexImdbId, plexTitle, plexOrigTitle, plexYear);
    if (!matchedRow) return { matched: false };

    const priorityCol = findPriorityColumn(sheet);
    sheet.getRange(matchedRow, priorityCol).setValue("Completed");
    moveRowToHistory(sheet, matchedRow);
    return { matched: true, row: matchedRow };
  } finally {
    lock.releaseLock();
  }
}

function findMovieRowForPlex(values, formulas, richText, plexImdbId, plexTitle, plexOrigTitle, plexYear) {
  // 1. Exact IMDb ID match (most reliable)
  if (plexImdbId) {
    for (let i = 1; i < values.length; i++) {
      const rowImdb = extractImdbId(formulas[i][0], richText[i][0]);
      if (rowImdb && rowImdb === plexImdbId) {
        return i + 1;
      }
    }
  }

  // 2. Normalized Title (+ optional Year) fallback
  const targetNorms = [];
  if (plexTitle) {
    const n1 = normalizeForCompare(sanitizeTitleAndYear(plexTitle).cleanTitle);
    if (n1) targetNorms.push(n1);
  }
  if (plexOrigTitle) {
    const n2 = normalizeForCompare(sanitizeTitleAndYear(plexOrigTitle).cleanTitle);
    if (n2 && targetNorms.indexOf(n2) === -1) targetNorms.push(n2);
  }
  if (targetNorms.length === 0) return null;

  let fallbackNoYearRow = null;
  for (let i = 1; i < values.length; i++) {
    const rawTitle = extractTitle(values[i][0], formulas[i][0], richText[i][0]);
    if (!rawTitle) continue;

    const { cleanTitle, year: titleYear } = sanitizeTitleAndYear(rawTitle);
    const rowNorm = normalizeForCompare(cleanTitle);
    if (targetNorms.indexOf(rowNorm) === -1) continue;

    if (plexYear) {
      const recDate = sanitizeAndRecoverDate(values[i][2], titleYear);
      const rowYear = (recDate && recDate.year) || (titleYear ? parseInt(titleYear, 10) : null);
      if (rowYear && Math.abs(rowYear - plexYear) <= 1) {
        return i + 1;
      }
      if (!rowYear && !fallbackNoYearRow) {
        fallbackNoYearRow = i + 1;
      }
    } else {
      return i + 1;
    }
  }

  return fallbackNoYearRow;
}

