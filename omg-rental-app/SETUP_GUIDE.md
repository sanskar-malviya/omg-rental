# OMG Rental — Go-live setup (about 15 minutes)

Your data lives in the Google Sheet **OMG Rental Database**. The app on Netlify reads and saves everything through a small Google Apps Script inside that sheet.

```
Tablet / phone / laptop  →  Netlify (the app)  →  Apps Script (PIN login, saving, booking lock)  →  Google Sheet (the data)
```

## Step 1 — Add the script to the sheet
1. Open the sheet: https://docs.google.com/spreadsheets/d/1wKoSapOIGER0jh_6S7p_l3zcshfursnaJUoParbcEhw/edit
2. Menu **Extensions → Apps Script**.
3. Delete everything in `Code.gs`, then paste in the whole contents of `google-apps-script/Code.gs` from this folder. Press **Save** (disk icon).
4. At the top, choose the function **setup**, then press **Run**.
   - Google asks for permission: **Review permissions → choose your account → Advanced → Go to project (unsafe) → Allow**. (It is your own script; Google shows this for every new script.)
   - When it finishes, the sheet has 16 tabs (Items, Customers, Bookings, Payments…).

## Step 2 — Publish the script as the app's API
1. In Apps Script: **Deploy → New deployment**.
2. Click the gear icon → **Web app**.
3. **Execute as: Me** · **Who has access: Anyone**. (This allows the app to reach the script. The script still requires a staff PIN for every action.)
4. Press **Deploy** and copy the **Web app URL** (ends with `/exec`).

## Step 3 — Put the URL in the app
Open `netlify-site/config.js` in Notepad and paste the URL:
```js
window.OMG_API_URL = 'https://script.google.com/macros/s/XXXXXXXX/exec';
```
Save the file.

## Step 4 — Put the app on Netlify
1. Go to https://app.netlify.com → **Add new site → Deploy manually**.
2. Drag the **`netlify-site`** folder onto the page. You get a link like `https://omg-rental.netlify.app`.
3. Optional: **Site settings → Change site name**, or add your own domain.

## Step 5 — First login
1. Open the Netlify link, choose **Shop Owner**, and enter PIN **1234**.
2. **Straight away:** tap your name (top right) → **Change my PIN**.
3. **Settings → Users & roles → Add user** for each staff member (name, role, PIN).
4. Optional: **Settings → Connection → Import sample data into the sheet** to try the app with the demo Garba stock. For real use, add your own items instead (Inventory → Add item).
5. On the shop tablet, open the link in Chrome → menu → **Add to Home screen**.

## Updating the script later
If you change `Code.gs`: **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. The URL stays the same.

## Good to know
- **Don't edit booking or payment rows by hand** in the sheet while the shop is open. Use the app. Editing items, prices or customers in the sheet is fine; tablets pick up changes within about a minute (or tap **Saved → refresh**).
- **Never delete rows.** Cancelled or void bookings keep their history, which is the audit trail.
- **Backups:** Google Sheets keeps version history (File → Version history). Also use **Settings → Export & backup** weekly to download CSV copies.
- **Internet drops:** the top bar shows *Offline · not saved*. Keep working; changes save automatically when the connection returns. Don't close the tab until it shows *Saved*.
- **Limits:** Google Sheets suits one shop with a few tablets and a few thousand bookings. Each save takes 1–3 seconds. If you grow to multiple branches or many staff, move to a proper database. The app's data layer (`toTables` / `fromTables` / `api`) is the only part that needs to change.

## Photos on GitHub (item photos + damage photos)
Photos are compressed on the tablet (about 1280 px, 150–280 KB). The Apps Script uploads them to the **`photos` branch** of this repository (`sanskar-malviya/omg-rental`), and the raw URL is saved in the **Item_Photos** tab. Using a separate branch means uploads never trigger a Netlify rebuild and never mix with the code.

The repository is public, so photo links work, but anyone with a link can see that photo. Only upload item and damage photos. **Never upload customer ID proofs.**

### 1. Create a GitHub token (one time)
1. Open https://github.com/settings/personal-access-tokens/new (GitHub → your photo → Settings → Developer settings → Personal access tokens → **Fine-grained tokens** → Generate new token).
2. **Token name:** `OMG Rental photos` · **Expiration:** 1 year (set a reminder to renew it).
3. **Repository access:** *Only select repositories* → pick **omg-rental**.
4. **Permissions → Repository permissions → Contents → Read and write.** Leave everything else as it is.
5. Click **Generate token** and copy it (it starts with `github_pat_`). GitHub shows it only once.

### 2. Put the token in Apps Script (never in the app or on GitHub)
Apps Script → ⚙ **Project Settings** → scroll to **Script Properties** → **Add script property**:
- Property: `GITHUB_TOKEN` · Value: your token → **Save script properties**

(Optional: `GITHUB_REPO` and `GITHUB_BRANCH`, only if you ever want a different repository or branch. The defaults are `sanskar-malviya/omg-rental` and `photos`.)

### 3. Update the script
1. Paste the latest `Code.gs` → Save.
2. Run **setup** once and click **Allow**. Google asks for a new permission, *“Connect to an external service”*, which is needed to talk to GitHub.
3. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy.**

### 4. Check
**Settings → Connection** shows *Photo storage (GitHub): connected ✓*. Open an item → **Photos → Add photo**. The file appears at `github.com/sanskar-malviya/omg-rental/tree/photos/items/<CODE>/`.

Files are stored as `items/<ITEM CODE>/…jpg` and `damage/<BOOKING NO>/<ITEM CODE>-…jpg`. "Remove" hides a photo in the app. The file stays in the branch history.


## QR tags & scanning
- **QR tags (PNG):** open an item → **Download QR** to save `QR-<CODE>.png` (QR + code + name). Or go to **Inventory → QR tags** to download a PNG for every piece in the current filter (up to 60 at a time; Chrome may ask once to allow multiple downloads). Print or share the PNGs however you like. Each QR holds the item's app link, e.g. `https://omgrental.netlify.app/#/item/G021`.
- **Scan:** tap the scan icon in the top bar (opens the item) or **Scan tag to add** in New Rental (adds the piece to the booking). The first time, allow camera access.
- The camera only works on the **https** Netlify link, not when opening the HTML file directly. If the camera was blocked: tap the 🔒 in the address bar → Permissions → Camera → Allow.
- A USB/Bluetooth barcode scanner also works: it "types" the tag into the search box or the scan popup.
- Scanner libraries are bundled in `netlify-site/vendor/` (jsQR, Apache-2.0; qrcode-generator, MIT; see `vendor/LICENSES.txt`).
