# NuORDER Client

A Windows program that turns brand linesheets and order forms into NuORDER-ready product data, handles product images, updates existing linesheets, and keeps a history of every run.

Current version: **3.5**

## Easiest way: download the compiled program (no building)

GitHub compiles the program by itself every time it changes. Download the newest one here:

**https://github.com/dioniciogarcia/tirsorder/releases/latest/download/NuORDER-Client.exe**

(or open the [Releases page](https://github.com/dioniciogarcia/tirsorder/releases)). Put it anywhere you like and double-click it.

**Updating:** open **Settings** in the program and press **Check for updates**, then **Update now**. The program downloads the newest version from GitHub, replaces itself and restarts. History and settings are kept. It also checks by itself when it starts (you can switch that off in Settings) and shows "Settings • update" in the menu when something new is available.

### How a new version is published
1. Change `nuorder_client.py` and raise `APP_VERSION` (for example 3.5 → 3.6).
2. Add a `### Version 3.7 — one less tab
- The separate "MaxMara PDF" page was removed: everything it did is now inside the first tab (Clean Data & Images, Images section) and the Image Renaming tab.

### Version 3.6` section to the Changelog below (it becomes the release notes).
3. Push to `main`. GitHub (Actions tab → *Build NuORDER Client*) compiles `NuORDER-Client.exe` on Windows and publishes it as release `v3.6`, together with `nuorder_client.py`. This takes about 5 minutes.

## How to build and run yourself (Windows)

1. Install Python 3 from [python.org](https://www.python.org/downloads/). During setup, tick **"Add python.exe to PATH"**.
2. Put these three files in the same folder:
   - `nuorder_client.py`
   - `Build_NuORDER_Client.bat`
   - `Run_NuORDER_Client_without_building.bat`
3. Double-click **`Build_NuORDER_Client.bat`** and wait until the black window says **DONE!** (3–5 minutes). It installs openpyxl, Pillow and PyInstaller, then creates **`NuORDER Client.exe`** in the same folder.
4. Open **`NuORDER Client.exe`**.

To try the program without building the .exe, double-click `Run_NuORDER_Client_without_building.bat`.

## What it does

- **Linesheet Updater**: applies a brand's update file (prices, style numbers, names, season, dates, sizes, UPCs…) to a processed file or NuORDER data export, highlighting only the changed cells.
- **Image Renaming**: matches a folder of images to the products of a file and copies them renamed (`BRAND_ID_01`, `_02`…) and resized.
- **MaxMara PDF**: pulls the product pictures out of MaxMara's PDF catalogs (one JPG per product code starting with MML, saved as CODE.jpg), ready for the Image Renaming page.
- **History**: every run, with a saved copy of each cleaned data file.

Brand files, results and the built .exe are never uploaded to this repository (see `.gitignore`).

## Changelog

### Version 3.6 — MaxMara PDFs everywhere
- The first tab (Clean Data & Images) now has a "MaxMara PDF catalogs" block in the Images section: add the PDFs, tick to extract the pictures and/or fill sizes and fabric from them.
- The Image Renaming tab can also import MaxMara PDFs.
- "Pick a brand folder" now also finds the PDFs in the folder.
- Pictures taken from PDFs follow the same resizing rules as every other picture (shortest side, JPG, Lanczos).

### Version 3.5 — updates from GitHub
- New **Settings** page with **Check for updates** and **Update now**: downloads the newest compiled program from this repository's Releases, replaces the running one and restarts it. Shows the release notes of the newest version.
- Optional check at start-up (on by default), with a "Settings • update" marker in the menu.
- GitHub compiles `NuORDER-Client.exe` automatically (Actions workflow `Build NuORDER Client`) and publishes it as a Release, so building on the laptop is no longer needed.
- Running the Python file instead of the .exe also works: Update replaces `nuorder_client.py` (a `.bak` copy is kept).

### Version 3.4 — MaxMara data format
- **New data type "MaxMara"** (detected automatically from the Excel columns Exit Collection Description / Ship Window / VPN / Model / Class Description / Variant). Every sheet (MM1, MM2 ...) becomes its own linesheet.
- **Column rules (set up like MaxMara's NuORDER portal data):** STYLE NUMBER = VPN, NAME = Model (MML...), COLOR / COLOR CODE = Color Description / Variant, DESCRIPTION = Class Description, WHOLESALE USD = Cost, RETAIL USD = Retail, SEASON as written in the file name (SS27), BRAND_ID = STYLE NUMBER_COLOR, LINESHEET NAME = season - sheet (SS27 - MM1 RP).
- **Ship Window** becomes AVAILABLE FROM / UNTIL using the season year (Dec 1 - Feb 28 of SS27 = 12/1/2026 - 2/28/2027).
- **PDF catalogs in the Clean Data list:** add the brand's PDFs next to the Excel and the program reads them for the sizes (a size is offered when its quantity is above 0, per color), the fabric composition and the pictures, which are saved as BRAND_ID_01 (resized like the other images). A composition that does not add up to 100% is flagged in the log (the PDF text itself may be cut off).
- **PDF reader** improved over Randy's script: handles rotated pages, finds the same model on several pages, and flags anything it cannot place with certainty.

### Version 3.3 — MaxMara PDF images
- New **MaxMara PDF** page (based on Randy's extractor script): add one or more PDF catalogs, choose a folder, and the program saves the picture of every product code that starts with MML as `CODE.jpg`.
- The first picture of a repeated code is kept, existing files are never overwritten, and codes where the picture can't be identified with certainty are listed in the log to check by hand.
- Optional resizing (off by default on this page; the Image Renaming page resizes afterwards).
- The build files now install PyMuPDF.

### Version 3.2 — new design
- **Left sidebar**, always open, with the logo, version and the four sections as a menu. The current section is highlighted.
- **Accent colors:** Red, Purple, Pink, Blue (the default) and two-tone Olive & Old Rose. Pick one from the swatches at the bottom of the sidebar, and the program remembers it.
- **Theme switch:** Auto (follows Windows), Light or Dark. Dark mode also darkens the window's title bar on Windows.
- **Top bar** on every page with the title, a description and a summary badge, for example "2 files · 1 brand detected" or "4 runs saved".
- **White cards** with thin borders, numbered step circles and a short description under each title.
- **Collapsible sections** (Advanced, More settings, Images) are clean bars with a **+** box, title, description and a **Show / Hide** link.
- **Buttons** are white with a thin outline and turn to the accent color when you point at them. **Run** is filled with the accent color.
- **Sizes and spacing follow the Windows zoom level** (100%, 125%, 150%), so nothing gets cut off.
- **File list:** recognized brands show in green, and column headings are aligned left.
- **History buttons** are on two rows so none get cut off.

### Version 3.1 — image tools, saved copies, UPCs
- **Image Renaming tab:**
  - Pick product file(s) and an images folder, or a whole brand folder.
  - Works with processed files, NuORDER data exports and raw brand files. A raw file is only read, never saved.
  - Images are matched the same way as in the first tab and copied with new names into a `Renamed Images` folder inside the images folder.
  - Writes `Image_Renaming_Log.csv` and `Renaming_Log.txt`.
  - Runs are recorded in History.
- **Image resizing:**
  - Always on by default in both image tabs.
  - Same settings as XnConvert: shortest side 1024 px (adjustable from 500 to 6000), proportions kept, no borders, Lanczos, always resize.
  - Saved as JPG, about 100–200 KB each.
  - Transparent backgrounds become white, and photos are turned the right way up.
- **Pick a brand folder** button (Clean Data and Image Renaming tabs): finds every data file (.xlsx, .xlsm, .xlsb, .csv) and image in the folder and its sub-folders. It skips Excel's temporary files and anything the program made before.
- **Images section** is collapsed by default, and its bar shows what's on.
- **Linesheet list in the status box** at the end of each run: every linesheet name with its product count and row color.
- **Saved copies in History:**
  - Every Clean Data run keeps a copy of its cleaned data file (no images).
  - New **Open saved copy** and **Save copy as…** buttons.
  - Deleting a run, or **Clear all**, deletes its copy too, after a warning.
  - The run details show the linesheet names, and the Excel export includes linesheet names and copy locations.
- **Update UPCs** in the Linesheet Updater:
  - Each UPC goes into the UPC column of its matching size.
  - EAN, barcode and GTIN columns count as UPCs.
  - Works with one-row-per-size update files and UPC 1 / SIZE 1 layouts.
  - A UPC that lost its leading 0 in Excel gets it back.
  - A UPC that can't be placed (size not in the file, or no color given for a multi-color style) is listed on a new **"UPCs not placed"** sheet.
  - "Everything found" also updates UPCs.
- **Updater change log:** a product matched by style + color now says so.
- **Build file** now installs Pillow, which the image resizing needs.

### Version 3.0
- **History tab:** every Clean Data run is recorded and kept forever. You can search, open a run's folder or log, edit the brand, export to Excel, delete selected runs or clear all.
- **Linesheet Updater tab:**
  - Applies a brand's update file to a processed file or NuORDER data export.
  - Matches style numbers regardless of spaces, dashes or capitals, and scans every sheet and column.
  - "What to update" choices: everything, pricing, style numbers, names, season, dates, sizes or picked columns.
  - Blank cells never erase data, and a style without a color updates all its colors.
  - Only changed cells are highlighted yellow; `_id` and column order are untouched.
  - Saves `_Updated.xlsx` and a `_Changes.xlsx` log (Changes, Not in data file, Not in update, Column pairing, Conflicts).
- **Summer 27 fixes:**
  - Every sheet with the same layout is processed.
  - SIZE 1…N columns win over a grouped sizes column.
  - Sizes: 00-14 expands in even steps, `*` is removed, sizes are capitals and stored as text.
  - Available Month goes to DELIVERY MONTH, and dates are real dates shown as 1/15/2027.
- **Armani profiles:** GA Men, EA Men, EA Women, Bodywear and Armani Exchange.
  - Suit size scales expand (46S, 46R, 46L…).
  - UPCs come from BARCODE/EAN.
  - Pictures are named MODEL_FABRIC CODE_COLOR CODE.
- **Ferragamo update:** Italian or English headers, plus extra columns (Group, Typology, Heel High, Hardware Color…).
- **Gucci update:** BRAND_ID = Style_Material_Colour, plus Division, Seasonality, Delivery Month and Fabric Description.
- **Window:**
  - Collapsible Advanced and More settings sections. Advanced is hidden for brand profiles and opens by itself when the Style Number column is missing.
  - **Start over** button, which keeps the save folder.

### Before 3.0 — the foundation
- **Cleaning:** brand linesheets and order forms (.xlsx, .xlsm, .xlsb, .csv) become NuORDER-ready data.
  - Several files combine into one master file, with each linesheet in its own light row color and a legend.
- **Data Identifier:** recognizes JOOR exports, Ferragamo, Gucci, Velvet and NuORDER Default Data, and uses general rules with column mapping for anything else.
- **NuORDER rules:**
  - " RP" is added to every linesheet name, and an optional Linesheet Name box (for a single file) replaces the file's own name.
  - Color codes are padded to 3 digits.
  - WHOLESALE/RETAIL USD get 0.01 when there's no USD price.
  - Currency comes from the prices, and the season is kept as written.
  - Size ranges expand (23-34, XXS-XL).
- **Images:**
  - Pictures pasted in sheets are extracted, including in-cell pictures.
  - Image links are downloaded from every link column.
  - A folder of loose images can be matched to the products, including fuzzy names and short color names like BLK. Unclear images are skipped and listed.
  - Everything is named NAME_01, _02…
- **Built-in .xlsb reader**, so no extra software is needed.
- **Build files:** `Build_NuORDER_Client.bat` makes a single `NuORDER Client.exe`, and `Run_NuORDER_Client_without_building.bat` runs it directly.
