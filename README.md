# Bills & Car app

An installable web app for the household bills forecast, bank reconciliation, utilities and car expenses.

- **The app** (this repo) is hosted on GitHub Pages and installs to your phone's home screen. It opens instantly
  and works offline from the last data it saw.
- **Your data** stays in your two Google Sheets. The app reads and writes them through the Apps Script in
  `apps-script/`, which only answers requests that carry your app key.
- **Nothing private is in this repo.** The script URL and the key are typed into the app once per device and
  are kept only on that device. The repo can be public.

```
index.html              the app
manifest.webmanifest    makes it installable
sw.js                   keeps the app on the device so it opens offline
icons/                  home-screen icons
apps-script/            paste these into the bills sheet's Apps Script project (not used by GitHub)
```

## 1. Apps Script (in the bills Google Sheet)

1. If you haven't already: upload `Bills Dashboard.xlsx` to Drive, open it, then **File → Save as Google Sheets**.
   On its **Settings** tab set **Opening Balance**, **Opening Date** and **Car Sheet ID** (paste the car sheet's link).
2. In the bills sheet open **Extensions → Apps Script** and make the project match `apps-script/`:
   - `Code.gs` → replace its contents.
   - `Car.gs` → **+ → Script**, name it `Car`, paste.
   - `Api.gs` → **+ → Script**, name it `Api`, paste.
   - `appsscript.json` → **Project Settings → Show "appsscript.json"**, then replace its contents.
   - If there's an `Index` HTML file from the earlier version, delete it.
3. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy, authorise (if Google warns the app is unverified: **Advanced → Go to …**), and copy the
     **Web app URL** (it ends in `/exec`).
   - If you deployed an earlier version as "Only myself", archive that one under **Manage deployments**.
4. Reload the sheet. A **Dashboard** menu appears. Choose **Dashboard → Show app key** (authorise once more).
   It shows the script URL and your key.

"Anyone" means the URL can be reached from the app; it doesn't open your data. Every request without the right
key is refused, and visiting the URL in a browser shows only a one-line message.

## 2. GitHub Pages

1. Create a new repository, e.g. `bills-app`. Public is fine.
2. **Add file → Upload files**, drag in everything from this folder (including `icons/` and `apps-script/`), commit.
3. **Settings → Pages**: Source **Deploy from a branch**, Branch **main**, folder **/ (root)**, **Save**.
4. After a minute the app is at `https://<your-username>.github.io/bills-app/`.

## 3. On each phone or computer

1. Open the app URL. The **Connect** screen appears: paste the script URL and the app key.
2. Install it:
   - **iPhone (Safari):** Share → **Add to Home Screen**. The home-screen app keeps its own storage, separate from
     Safari, so open it from the home screen and connect once more there.
   - **Android (Chrome):** menu → **Install app** (or **Add to Home screen**).
   - **Computer (Chrome/Edge):** the install icon in the address bar.
3. Retire the old car app: in the car sheet's Apps Script, **Deploy → Manage deployments → Archive**.

The ⚙ button in the app changes the connection or wipes the key and saved data from that device.

## Keeping it safe

- Treat the app key like a password. If a device is lost: **Dashboard → Make a new app key** in the bills sheet,
  then reconnect your other devices with the new key.
- Data and the key are stored in the browser for `<your-username>.github.io`. Every GitHub Pages site under that
  same username shares that storage, so don't host pages there that you didn't write yourself.
- Some Google Workspace (work/school) accounts block "Anyone" deployments. A personal Gmail account is fine.

## Updating the app later

1. Upload the changed files to the repo.
2. In `sw.js`, change `VERSION` (e.g. `'v1'` → `'v2'`) and upload it too.
3. Devices pick up the new version in the background and use it from the next time they open the app.

If you change anything in `apps-script/`, paste it into Apps Script and **Deploy → Manage deployments → Edit →
New version**. The URL and key stay the same.
