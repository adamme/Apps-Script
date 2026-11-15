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
   - `DISCORD_WEBHOOK_URL` *(optional)*: Discord webhook for "Now available" posts when a movie moves into `New Releases` (or set it via **🎬 Movie Tools → 🔔 Set Discord Webhook URL**)
   - `PLEX_WEBHOOK_TOKEN` *(optional)*: High-entropy secret token for Plex webhook authentication (generated via **🎬 Movie Tools → 🔐 Setup Plex Webhook Security**)
   - `PLEX_SERVER_UUID` / `PLEX_ACCOUNT_TITLE` *(optional)*: Extra defense-in-depth checks to restrict Plex scrobble webhooks to your server UUID and username
4. Reload your Google Sheet and use the **🎬 Movie Tools** menu (optionally run **⏰ Enable Daily VOD & Category Automation** for daily updates).

### Plex Auto-Complete Webhook Setup

1. In the Apps Script editor, click **Deploy → New deployment → Web app**:
   - **Execute as:** Me
   - **Who has access:** Anyone
   - Copy the generated `/exec` URL.
2. In your Google Sheet, run **🎬 Movie Tools → 🔐 Setup Plex Webhook Security** to generate a secret token (and optionally lock it to your Plex username and server UUID).
3. In Plex (**Settings → Webhooks → Add Webhook**), paste `<YOUR_WEB_APP_EXEC_URL>?token=<YOUR_SECRET_TOKEN>`.
4. Whenever Plex sends a `media.scrobble` event for a movie, it is matched by IMDb ID (or Title + Year), marked `Completed`, removed from `Movies`, and logged to `History`.
