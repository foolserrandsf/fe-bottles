# Staff Bottle App – template

A phone app for bar staff, built from one Google Sheet. It's the same app Fool's Errand (San Francisco) runs, with every bar-specific detail moved into one settings file and the sheet itself.

**What staff get on their phones**
- The whole bottle list, grouped by menu section, with label photos, chalk prices and tags (New, ** chilled red / skin contact, blue "Chillable", 86'd).
- Tap a bottle for its drawer: a "remember it by" hook, tasting notes, terroir, producer story and glossary words.
- Search in plain words ("salty white", "chillable red under 70"). A search never hides a bottle: matches outside the filter that's on are listed underneath.
- **Mark 86'd / Back in stock**, logged by name, day and time.
- **+ Add bottle** and **Edit bottle**. Before saving, the app checks for duplicates and for 86'd or sold-out bottles and asks what to do. Label photos can be taken with the camera or picked from the phone's photos.
- **Print menu**: rebuilds the folded, double-sided guest bottle list PDF and opens it on the phone to print.
- **Change log**: every change with who and when, plus an **Undo** button on each one.
- Everything is locked behind an app PIN, and each person has their own codes for 86'd, add/edit and print, so the logs always show the right name.

**What the manager keeps**
- The Google Sheet is the single source of truth. The app reads it live and writes back to it.
- A "Bottle List" menu in the sheet: *Make menu tab*, *Make folded PDF* and *Set up this sheet*.

---

## What's in this folder

| File | What it is |
|---|---|
| `apps-script/Code.gs` | The sheet's script: the backend for the app plus the menu/PDF builder |
| `site/index.html` | The phone app (one file) |
| `site/config.js` | **Your settings**: bar name, colors, web app URL |
| `site/logo.png` | **Your logo** (square PNG, about 512×512). Used on the opening screen, in the header and in the loading animation |
| `site/icon-192.png`, `icon-512.png`, `apple-touch-icon.png` | Home-screen icons (replace with your logo at those sizes) |
| `site/manifest.webmanifest` | The app name on the home screen |
| `site/photos/`, `site/thumbs/`, `photos.json` | Optional: label photos hosted with the app (`photos/<ID>.jpg`, `thumbs/<ID>.jpg`, and the IDs listed in `photos.json`). Photos taken in the app go to Google Drive instead, and you don't need these. |

---

## Setup (about 30 minutes, no coding)

### 1. Make the sheet
1. In Google Drive, create a new Google Sheet, e.g. "Bottle List – <Your Bar>".
2. Open **Extensions → Apps Script**.
3. Delete what's in `Code.gs`, paste the whole of `apps-script/Code.gs`, and click **Save**.
4. On the left, click **Services (+)**, choose **Google Slides API** and click **Add** (needed for the printed PDF).
5. Go back to the sheet and reload it. A **Bottle List** menu appears after a few seconds.
6. Click **Bottle List → Set up this sheet (first time)**. Google asks for permission the first time: click **Continue → Allow**. If it says the app isn't verified, click **Advanced → Go to (project)**.

This creates the tabs:
- **Bottles**, with one example row and dropdowns on the columns that need them
- **Setup**
- **Staff app**, with the app PIN and staff codes
- **Menu layout**
- **Glossary**, already filled with starter wine terms
- **86 log** and **Bottle log**

### 2. Fill in the sheet
- **Setup tab:** your bar name. The other rows come in step 5.
- **Staff app tab:**
  - Change the **App PIN** in B1.
  - Replace the example "Manager" row with your staff: one row per person, with their own 86'd, add/edit and print codes. Leave a code blank to take that permission away.
- **Menu layout tab:** which menu section goes on which page of the folded list. Page 1 is the cover, so content starts on page 2. The app uses this order too.
- **Bottles tab:** your bottles. You can also add them later from the app.
  - **Menu section** must match a section on the Menu layout tab exactly.
  - **Taste group** decides the color filter in the app. Pick it from the dropdown.
  - **MENU:** columns are the short wording printed on the guest list.
  - **ID** must be unique per bottle (b1, b2, …). Bottles added in the app get new1, new2, … automatically.

### 3. Put the app online (GitHub Pages, free)
1. Create a free GitHub account and a new **public** repository, e.g. `bottles`.
2. Upload everything inside `site/` to the repository (**Add file → Upload files**).
3. Go to **Settings → Pages → Build and deployment**, set **Source: Deploy from a branch**, **Branch: main / (root)**, and click **Save**.
4. After a minute your app is at `https://<your-github-name>.github.io/bottles/`.

### 4. Connect the app to the sheet
1. In Apps Script, click **Deploy → New deployment → Select type → Web app**.
2. Set **Execute as: Me** and **Who has access: Anyone**, then click **Deploy** and allow access.
3. Copy the **Web app URL** (it ends in `/exec`).
4. In GitHub, open `config.js`, click the pencil, paste the URL into `apiUrl`, set `barName`, `listName` and your colors, and click **Commit**.
5. Open the app on a phone, enter the App PIN, then **Share → Add to Home Screen**.

> **After any change to the script**, publish it again: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. The URL stays the same.

### 5. Turn on the printed bottle list (optional)
1. In Google Slides, create a blank presentation and choose **File → Page setup → Custom → 11 × 8.5 inches**. Copy the long ID from its address bar (between `/d/` and `/edit`) into **Page template ID** on the Setup tab.
2. Optional: upload your logo to Drive and copy its file ID into **Cover logo file ID**. Without it, the cover prints your bar name.
3. Optional: create a Drive folder for label photos and copy its ID into **Label photo folder ID**. If you leave it blank, photos go next to the sheet.
4. Try **Bottle List → Make folded PDF**. Print it double-sided, flip on short edge, at actual size, then fold.

---

## Branding
- **Name and colors:** `config.js` (barName, listName, accent, "remember it by" box colors, background, text).
- **Logo:** replace `logo.png` and the three icon files with your own, keeping the same file names.
- **Home-screen name:** edit `name` and `short_name` in `manifest.webmanifest`.
- After changing files, bump the number in **both** `version.json` and the line `const APP_V = '…'` near the bottom of `index.html` (make them the same, e.g. `2`). Phones then reload the new version by themselves.

## Good to know
- **Codes are private.** They live only on the Staff app tab. Never put real codes in the app files, because the GitHub repository is public.
- **Statuses** (first column of Bottles):
  - `menu`: on the guest list and in the app
  - `off-menu`: in the app as "Shelf only", not printed
  - `insert` / `needs-page`: printed and flagged New
  - `sold-out` / `removed`: hidden everywhere
- **Taste groups** drive the color chips. Reds: Light & chillable, Bright & silky, Savory & spicy, Firm & structured, Dark & structured. Also Bubbles, Pink, Orange & grippy, and the white groups. Reds in *Light & chillable reds* that aren't marked "Serve chilled" get the blue **Chillable** tag.
- **Undo** puts back exactly what that one change altered. Undoing a newly added bottle sets its status to `removed`; nothing is ever deleted.
- **Too many wrong codes** (12 in 10 minutes) lock code entry for 10 minutes.
- **The booklet columns** (remember it by, tasting, terroir, producer) are filled in on the sheet. Bottles added from the app are marked "Facts verified? NO" with a note, so you know which pages still need writing.
