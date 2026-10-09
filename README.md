# Sankey Chart Builder for Excel (v2.0)

Creates professional Sankey (flow) diagrams from a **From | To | Value** table, drawn as native Excel
shapes, with a **Sankey ribbon tab** and **automatic updates** when your data changes.

## Files
| File | Purpose |
|---|---|
| `SankeyAddin.bas` | Main code (import into VBA) |
| `SankeyEvents.cls` | Auto-update event hook (import into VBA) |
| `customUI14.xml` | Ribbon definition |
| `Install-Ribbon.ps1` | Adds the ribbon tab to the saved `.xlam` |
| `sample-data.csv` | Example data (or use *Sample data* on the ribbon) |

## Install (one time)
1. Open Excel > **Alt+F11**. If you installed an older version, remove the old `SankeyAddin` module
   (right-click > Remove > *No* to export) and untick/delete the old add-in.
2. **File > Import File...** and import **both** `SankeyAddin.bas` **and** `SankeyEvents.cls`.
3. Close the VBA editor. **File > Save As > Excel Add-in (\*.xlam)**, name it `Sankey`
   (keep the default `AddIns` folder).
4. **Close Excel completely**, then right-click `Install-Ribbon.ps1` > **Run with PowerShell**.
   (If your add-in is elsewhere: `.\Install-Ribbon.ps1 -AddinPath "C:\path\Sankey.xlam"`.)
5. Start Excel > **File > Options > Add-ins > Manage: Excel Add-ins > Go...** > tick **Sankey**.
   A **Sankey** tab appears next to Home.

No ribbon (skipped step 4)? A small floating "Sankey" toolbar appears on the Add-ins tab instead
(Create / Title / Delete / Sample data / Help).

## Where to find it
- **Sankey** tab next to Home (all options), and a **Sankey** button at the end of the **Insert** tab (Create / Sample data / Delete / Help).
- **Customize Ribbon / Quick Access Toolbar**: Excel cannot edit add-in tabs, but the commands can be added to your own tab: *Choose commands from > Macros* and pick `CreateSankey`, `InsertSampleData`, `DeleteSankeyActive`, `ChangeNodeColor`, `ResetNodeColors`, `ShowHelp`.

## Use
1. Lay out three columns: **From | To | Value** (header optional).
2. Select the table (or click one cell inside it) > **Sankey > Create / Update**.
3. Style it from the ribbon: title, subtitle, colors, flow style, values, size, unit.
   Changes restyle the chart on the active sheet immediately.
4. Try **Sample data** for a ready-made example.

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
- Windows Excel 2010+ only.


---
## Power BI version (SankeyFlow.pbiviz)
A Power BI custom visual with the same look: **Power BI Desktop > Visualizations pane > ... > Import a visual from a file > SankeyFlow.pbiviz**, then drop **Source**, **Target** and **Value** fields on it. Source code and build: `SankeyFlow/` (`npm run package`).


---
## HTML version (SankeyChart.html)
One self-contained file that works offline: open it in Edge or Chrome. It loads the sample data on start. Upload CSV/TSV/XLSX (or drop a file, or paste from Excel), map the columns, edit the data table, **drag nodes** to reposition them (flows follow live), click a node to recolor it, use **Select area** to crop, then **Copy PNG** (or Ctrl+C) or **Export PNG/SVG**. Try `sample-upload.xlsx`. Source: `SankeyHtml/` (rebuild with `node build.mjs`).
