# Outline Tabs for Word

A Google Docs-style document outline for Microsoft Word. A side panel lists every heading in your document; click one to jump to it, collapse sections with the arrows, and filter by typing. The section your cursor is in is highlighted and the list updates as you type.

Headings are detected from Word's built-in styles: Title and Heading 1 to Heading 6.

## Files

- `manifest.xml`: the add-in definition Word loads
- `taskpane.html`, `taskpane.css`, `taskpane.js`: the panel itself
- `assets/`: icons

## Install (about 5 minutes)

Word add-ins must be served over HTTPS, so the files need to live at a web address. GitHub Pages is free and the easiest option.

### 1. Host the files

1. Create a new public GitHub repository named `outline-tabs`.
2. Upload `taskpane.html`, `taskpane.css`, `taskpane.js` and the `assets` folder to the repository root.
3. In the repository, open Settings > Pages, set the source to the `main` branch, and save.
4. After a minute your files are live at `https://YOUR-USERNAME.github.io/outline-tabs/`.

### 2. Point the manifest at your site

Open `manifest.xml` and replace every `YOUR-USERNAME` with your GitHub username (find and replace; it appears in about a dozen places).

### 3. Load it into Word

**Word on the web (fastest to try):**
Insert > Add-ins > More Add-ins > My Add-ins > Upload My Add-in, then pick `manifest.xml`.

**Word for Windows (permanent):**
1. Put `manifest.xml` in a folder and share it (right-click > Properties > Sharing).
2. In Word: File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs. Add the shared folder's network path (for example `\\YOURPC\OutlineTabs`) and tick "Show in Menu".
3. Restart Word, then Insert > My Add-ins > Shared Folder > Outline Tabs > Add.

**Word for Mac:**
Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the folder if needed), restart Word, then Insert > My Add-ins and pick Outline Tabs.

### 4. Use it

Home tab > **Document tabs**. The panel opens on the right. If you want it to stay open, leave it docked.

## Notes

- Only paragraphs with a heading style appear, so format your headings with the Styles gallery.
- Collapsed/expanded state is remembered between sessions.
- The panel refreshes about half a second after you stop typing, plus a 5-second safety refresh. The circular arrow button forces an immediate refresh.
