# Sankey Chart Builder for Excel (v2.0)

Creates professional Sankey (flow) diagrams from a **From | To | Value** table, drawn as native Excel
shapes, with a **Sankey ribbon tab** and **automatic updates** when your data changes.

## Files
| Where | What |
|---|---|
| `Install-Sankey-Windows.cmd` / `Uninstall-Sankey-Windows.cmd` | One-step (un)installer for Excel on Windows |
| `Install-Sankey-macOS.command` / `Uninstall-Sankey-macOS.command` | One-step (un)installer for Excel on macOS (Microsoft 365) |
| `windows/` | Windows add-in: `SankeyAddin.bas`, `SankeyEvents.cls`, `customUI14.xml`, `images/`, PowerShell scripts used by the installers |
| `mac/` | macOS add-in: the same code adapted for Mac (`SankeyAddin.bas`, `SkDict.cls`, `SankeyEvents.cls`, `customUI14.xml`), the prebuilt `Sankey.xlam`, and `Build-Mac-Addin.cmd` |
| `sample-data.csv`, `sample-upload.xlsx` | Example data |
| `SankeyChart.html`, `SankeyHtml/` | HTML version |
| `SankeyFlow.pbiviz`, `sankeyFlow/` | Power BI visual |

## Quick install (automatic)
**Windows:** close Excel, then double-click **`Install-Sankey-Windows.cmd`**. It builds `Sankey.xlam` from the code files, adds the ribbon and enables the add-in; the Sankey buttons appear on the **Add-ins** tab. It briefly turns on Excel's *Trust access to the VBA project object model* and restores your previous setting; if your IT policy locks that setting, use the manual steps below. Remove with **`Uninstall-Sankey-Windows.cmd`**.

**macOS (Microsoft 365):** quit Excel, then double-click **`Install-Sankey-macOS.command`** (or run `bash Install-Sankey-macOS.command` in Terminal; if macOS blocks a downloaded file, right-click > Open). It copies `mac/Sankey.xlam` into Excel's Startup folder; the two Sankey icons (create chart, sample data) appear on an **Add-ins** tab, and the **Sankey** design tab appears when you click the chart, as on Windows. Choose *Enable Macros* if Excel asks. Remove with **`Uninstall-Sankey-macOS.command`**.
The Mac version differs slightly: node colors are entered as hex codes (Excel for Mac has no color dialog for macros), and instead of a contextual tab under Shape Format, the Sankey design tab is shown while the chart is selected (it can take up to a second to appear).

**Updating `mac/Sankey.xlam`:** it is built on a Windows PC with Excel: close Excel, run `mac\Build-Mac-Addin.cmd`, and commit the new file.

Charts already in your workbooks stay as ordinary shapes after uninstalling (they just stop auto-updating).

## Manual install (Windows, one time)
1. Open Excel > **Alt+F11**. If you installed an older version, remove the old `SankeyAddin` module
   (right-click > Remove > *No* to export) and untick/delete the old add-in.
2. **File > Import File...** and import **both** `windows/SankeyAddin.bas` **and** `windows/SankeyEvents.cls`.
3. Close the VBA editor. **File > Save As > Excel Add-in (\*.xlam)**, name it `Sankey`
   (keep the default `AddIns` folder).
4. **Close Excel completely**, then right-click `windows/Install-Ribbon.ps1` > **Run with PowerShell**.
   (If your add-in is elsewhere: `.\windows\Install-Ribbon.ps1 -AddinPath "C:\path\Sankey.xlam"`.)
5. Start Excel > **File > Options > Add-ins > Manage: Excel Add-ins > Go...** > tick **Sankey**.
   Two Sankey icons appear on the **Add-ins** tab (create chart, sample data); hover for help.

No ribbon (skipped step 4)? A small floating "Sankey" toolbar appears on the Add-ins tab instead
(Create / Title / Sample data / Help).

## Use
1. Lay out three columns: **From | To | Value** (header optional).
2. Select the table (or click one cell inside it) > **Add-ins > Sankey icon**. With no valid data selected you get an empty chart area instead.
3. Click the chart (or the empty area): a contextual **Sankey** tab appears next to the Shape Format tab (Drawing Tools). Style it there: title, subtitle, colors, flow style, values, size, unit. Changes restyle the chart immediately.
4. To fill an empty chart area, select your data and click the Sankey icon again. The second icon inserts **Sample data**.
5. To remove a chart, select it and press Delete.

## Colors
- **Colors** dropdown: 5 built-in palettes, **Workbook theme** (the colors applied under Page Layout > Colors, e.g. a company theme) and **Custom palette**.
- **Import palette > From selected cells**: cells filled with your colors, or containing hex codes (#E4002B) / RGB (228,0,43). **From a file**: .txt/.csv/.gpl/.css with hex or RGB values, or an Office theme-colors .xml.
- **Change node color**: click the chart, click a node once more so it is selected alone, press the button, pick a color. Remembered per node (survives auto-update). **Reset node colors** undoes all of them.
- Palettes with fewer colors than nodes repeat; override individual nodes as needed.

## Auto-update
With **Auto-update chart** ticked (default), the chart redraws when:
- a value or name in the source table is edited (or recalculated by a formula);
- a new row is added directly below the table.

Each chart remembers its own source range and settings (stored in the workbook), so it keeps updating
after you save/reopen as long as the add-in is loaded. Select several cells and press *Create / Update*
to point the chart at a different table.

**Trade-off:** a redraw clears Excel's Undo history. Untick Auto-update if you rely on Ctrl+Z while editing.

## Notes
- Resize the chart with the mouse and the ribbon Width/Height boxes update (checked once per second); the chart is re-laid out at the new size.
- Settings are stored inside each workbook (hidden names on the chart's sheet), never globally: a new file always starts from defaults and files don't affect each other.
- One chart per sheet. Flows must go one way (A > B > C); circular flows are rejected.
- The chart is a group of shapes (`Sankey_Chart`): you can move/resize it; ungrouping stops auto-update.
- Windows Excel 2010+, or Excel for Mac (Microsoft 365).


---
## Power BI version (SankeyFlow.pbiviz)
A Power BI custom visual with the same look: **Power BI Desktop > Visualizations pane > ... > Import a visual from a file > SankeyFlow.pbiviz**, then drop **Source**, **Target** and **Value** fields on it. Source code and build: `SankeyFlow/` (`npm run package`).


---
## HTML version (SankeyChart.html)
One self-contained file that works offline: open it in Edge or Chrome. It loads the sample data on start. Upload CSV/TSV/XLSX (or drop a file, or paste from Excel), map the columns, edit the data table, **drag nodes** to reposition them (flows follow live), click a node to recolor it, use **Select area** to crop, then **Copy PNG** (or Ctrl+C) or **Export PNG/SVG**. Try `sample-upload.xlsx`. Source: `SankeyHtml/` (rebuild with `node build.mjs`).
