# Movie Tracker (Google Apps Script)

Automated movie watchlist & release tracker for Google Sheets using IMDb, Rotten Tomatoes (Popcornmeter + Tomatometer fallback), OMDb, DVDsReleaseDates, Wikidata, and Wikipedia.

## Features

- **4-Stage Automated Lifecycle** (`To be released` → `In Theaters` → `New Releases` → Genre / `Time Travel`)
- **Rotten Tomatoes Popcornmeter & Links**: Populates Column D with audience scores (or critics score fallback) linked directly to each movie's RT page.
- **Netflix & VOD Tracking**: Detects Netflix exclusives (writing `Netflix` in Notes and skipping `In Theaters`), tracks VOD release dates for theatrical releases, and marks `"Release Date Unannounced"` for `12/31` / year-only placeholders until an exact date is announced.
- **Time Travel Auto-Category**: Automatically categorizes movies involving time travel as `Time Travel` once they graduate from `New Releases`.
- **Priority Tiers & Tab Routing**:
  - `Completed`: Removes the row from `Movies` and logs title (with IMDb link) + current date to the `History` tab.
  - `Won't Watch`: Moves the row to the `Won't Watch` tab.
  - Preserves Google Sheets **Chip** dropdown styling and custom colors.
- **Interactive Modals**:
  - **Add New Movie** (`addNewMovieDialog`) with IMDb search & priority selector
  - **Quick Prioritizer** (`PriorityPicker.html`) with posters, RT scores, plot summaries, and trailer search
  - **Random High Priority Movie Selector** (`RandomMoviePicker.html`)
  - **Fix Selected Row** (`fixSelectedRowDialog`)

## Setup

1. Copy `.clasp.json.example` to `.clasp.json` and replace `YOUR_APPS_SCRIPT_ID_HERE` with your Google Apps Script ID (keep `.clasp.json` gitignored).
2. Push the project with `clasp push`.
3. In the Apps Script editor, go to **Project Settings → Script Properties** and add:
   - `OMDB_API_KEY`: your free OMDb API key from <https://www.omdbapi.com/apikey.aspx>
4. Reload your Google Sheet and use the **🎬 Movie Tools** menu (optionally run **⏰ Enable Daily VOD & Category Automation** for daily updates).
