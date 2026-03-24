# ShapR Tabler Icons — PowerPoint Add-in

A Microsoft Office PowerPoint task pane add-in that lets you search, browse, and insert any of the 5,000+ Tabler outline icons directly into slides — with ShapR brand colours and configurable sizes.

## How it works

1. On first open, the add-in fetches the complete Tabler Icons outline sprite (a single SVG file) from the jsDelivr CDN
2. All icon shapes are parsed client-side from `<symbol>` elements — no per-icon requests needed
3. A virtual-scroll grid renders only the icons in view, keeping performance smooth with 5,000+ icons
4. Clicking an icon calls `PowerPoint.run` with `shapes.addSvg()` to insert a coloured SVG shape onto the active slide

---

## Project structure

```
tabler-icons-addin/
├── manifest.xml       Office add-in manifest (sideloading + command surface)
├── taskpane.html      Task pane UI
├── taskpane.css       Styles
├── taskpane.js        All application logic
├── server.js          Express dev server (HTTP, localhost:3000)
├── package.json
├── README.md
└── assets/
    ├── icon-16.svg    Add-in icon 16×16
    ├── icon-32.svg    Add-in icon 32×32
    └── icon-80.svg    Add-in icon 80×80
```

---

## Local development — sideloading

### Prerequisites

- Node.js 16 or later
- Microsoft 365 subscription with PowerPoint (desktop or online)

### Steps

1. Install dependencies:

   ```bash
   cd c:\code\tabler-icons-addin
   npm install
   ```

2. Start the dev server:

   ```bash
   npm start
   ```

   The server runs at `http://localhost:3000`.

3. Sideload the manifest in PowerPoint Desktop:

   - Open PowerPoint
   - Go to **Insert** > **Get Add-ins**
   - Choose **Manage My Add-ins** > **Upload My Add-in**
   - Browse to `c:\code\tabler-icons-addin\manifest.xml`
   - Click **Upload**

   The "Tabler Icons" button appears in the **Home** ribbon under the **ShapR** group.

4. Open the task pane by clicking **Tabler Icons** in the ribbon.

### Testing the UI in a browser (no PowerPoint required)

Navigate to `http://localhost:3000/taskpane.html` in any browser.
A warning banner appears since Office.js is unavailable, but you can browse and search icons. Click events will show a toast indicating PowerPoint is required for actual insertion.

---

## Production deployment

### Option A — Microsoft 365 Admin Center (Centralized Deployment)

This is the recommended approach for deploying to your organisation.

1. **Host the files** on a web server with HTTPS. All files in this directory must be publicly accessible. Common options:
   - Azure Static Web Apps (free tier available)
   - Azure Blob Storage with a CDN endpoint
   - Any web host that supports HTTPS

2. **Update `manifest.xml`**:
   - Replace every occurrence of `https://localhost:3000` with your production URL
   - Generate a new GUID for the `<Id>` element to avoid conflicts:

     ```bash
     # PowerShell
     [System.Guid]::NewGuid().ToString()

     # Or use: https://www.guidgenerator.com/
     ```

   - Update the `<Version>` element (e.g. `1.0.1.0`) whenever you redeploy

3. **Deploy via Admin Center**:
   - Sign in to [admin.microsoft.com](https://admin.microsoft.com)
   - Go to **Settings** > **Integrated Apps** > **Upload custom apps**
   - Choose **Office Add-in** and upload the updated `manifest.xml`
   - Assign to users/groups

4. The add-in appears in PowerPoint for all assigned users within 24 hours (often much sooner).

### Option B — SharePoint App Catalog

1. Upload the manifest to your organisation's SharePoint App Catalog
2. Trust the add-in when prompted
3. The add-in becomes available to all site users

---

## HTTPS for non-localhost testing

Office add-ins require HTTPS when not running on `localhost`. To generate a self-signed certificate for local network testing:

### Using mkcert (recommended)

```bash
# Install mkcert
# Windows (Chocolatey):  choco install mkcert
# macOS:                 brew install mkcert

mkcert -install                        # Trust the local CA
mkcert localhost 127.0.0.1 ::1         # Generate cert for localhost
```

Then update `server.js` to use HTTPS:

```javascript
const https = require('https');
const fs    = require('fs');

const options = {
  key:  fs.readFileSync('localhost-key.pem'),
  cert: fs.readFileSync('localhost.pem'),
};

https.createServer(options, app).listen(3000, () => {
  console.log('HTTPS server running on https://localhost:3000');
});
```

### Using OpenSSL (alternative)

```bash
openssl req -x509 -newkey rsa:4096 -keyout key.pem -out cert.pem -days 365 -nodes \
  -subj "/CN=localhost"
```

---

## Updating ShapR brand colors

Brand colours are defined in two places. Update both when colours change.

### 1. CSS custom properties — `taskpane.css`

```css
:root {
  --color-navy:      #0B2B45;   /* Dark Navy — Primary */
  --color-gold:      #F2BE22;   /* Gold — Secondary */
  --color-magenta:   #D8256A;   /* Magenta — Tertiary */
  --color-dark-teal: #185359;   /* Dark Teal */
  --color-teal:      #30A5BF;   /* Bright Teal — Interactive accent */
}
```

### 2. Colour swatches — `taskpane.html`

Find the `.color-swatches` section and update each `value` attribute and the `style="background:..."` inline style:

```html
<label class="swatch-label" title="New Brand Color #RRGGBB">
  <input type="radio" name="icon-color" value="#RRGGBB" class="sr-only" />
  <span class="swatch" style="background:#RRGGBB;" aria-label="Color Name"></span>
</label>
```

To add or remove swatches, add or remove `<label class="swatch-label">` blocks.
To change the default selected colour, add `checked` to the desired radio input.

---

## Icon insertion details

- Icons are inserted as **SVG vector shapes** using `PowerPoint.run` + `shapes.addSvg()`
- They are placed 1 inch from the top-left corner of the slide by default
- After insertion, resize and reposition using PowerPoint's standard shape handles
- SVG shapes can be ungrouped in PowerPoint to edit individual paths
- Fallback: if `addSvg` is unavailable (older Office versions), the icon is rasterised to PNG via canvas and inserted as an image

### Insertion position

To change the default insertion position, update `insertIcon()` in `taskpane.js`:

```javascript
firstSlide.shapes.addSvg(svgFinal, {
  left:   convertPxToEmu(72),   // pixels from left edge
  top:    convertPxToEmu(72),   // pixels from top edge
  width:  convertPxToEmu(size),
  height: convertPxToEmu(size),
});
```

`convertPxToEmu(px)` converts pixels to EMU (English Metric Units, the unit Office.js uses internally). 96px = 1 inch = 914400 EMU.

---

## Troubleshooting

| Issue | Solution |
|---|---|
| "Failed to load icons" | Check internet connectivity; the sprite is fetched from jsDelivr CDN |
| Add-in not appearing in ribbon | Re-sideload the manifest; check PowerPoint version supports add-ins |
| Icon inserts as small image instead of SVG | Your Office version may not support `addSvg`; the PNG fallback is used |
| "Office.js not available" banner | You're viewing the add-in in a browser, not PowerPoint — expected |
| Port 3000 in use | `PORT=3001 npm start` |

---

## License

Icons provided by [Tabler Icons](https://tabler.io/icons) — MIT License.
Add-in code — MIT License.
