# OCV Report Builder

https://sangnew.github.io/ocv-report-builder-web/

Turns the **Mass Production E&L Grade OCV Tracking Sheet** workbook into the Frozen IR / Spot report
(the `test1.xlsx` layout: LOT, Cell ID, Grade, dOCV, Frozen IR, voltage drop, spot details, sizes).

The workbook is read in the browser (never uploaded). After signing in with the team password, the built report and every value typed into it are shared with the whole team in real time (Firebase, same team account as the other tracker sites).

## How it works

1. Load the tracking workbook. Only the `Master E & L` sheet is read at first, which is fast even for 30MB+ files.
2. Pick LOTs. By default only cells that have their own OCV tracking sheet are included, which matches the sample report.
3. Review the table:
   - **Yellow**: needs manual input. Shape, Location and the sizes are never in the workbook. Other columns are yellow when the Master row is empty.
   - **Light blue**: the tracking-sheet analysis disagrees with Master E & L. Please verify these.
   - **Green**: typed by the team, or copied from a previous report.
4. Download the Excel file. Highlights are kept in the file, and a `Legend` sheet explains them.

Optionally, load a previous report in the same format to carry over hand-entered values by Cell ID.
Typed values are saved per Cell ID for the whole team, so they survive rebuilding the report. **Reset** can unload the workbook from your page only, or (after typing RESET) delete the shared report and all typed values for everyone.

## Column rules

| Report column | Source |
|---|---|
| LOT, Cell ID, Grade, dOCV | Master: Lot ID, Cell ID, Grade, dOCV (mV) |
| Frozen IR Result (35MOhm) | Master: Frozen IR (Mohm) |
| Frozen IR Pass/Fail | Master: Frozen IR Result; if empty, NG when Frozen IR < 35 MΩ, otherwise OK |
| voltage drop / no drop | From the cell's OCV tracking sheet: per layer (column B), dOCV = max(C−D, C−E) over the tracking dates in row 5. Drop when one inner layer is > 2.6σ above the others (same as the sheet's R6 formula) **and** ≥ 1.5 mV (adjustable); otherwise NTF. Light blue when this disagrees with Master E & L. Without a tracking sheet: Master E & L. |
| Voltage Dropped Layer | Layer number (column B) of that layer |
| dOCV (V) | Master: Voltage Drop (V) when it is for the same layer, else the computed dOCV (4 decimals) |
| Spot Found | "Spot Found" when Master: Burn mark/Pinhole/None is filled |
| Top/ Back, x, y | Master: Anode Top/Back, X (mm), Y (mm) |
| SEM/EDS Analysis | Master: EDS Impurity Results, else the element in the NTF column (Cu, Ni, Fe, …) |
| Shape, Location, Long side, Short side, Height | Manual |

## Files

- `convert.js`: conversion rules. Pure functions, also runnable in Node.
- `xlsx-subset.js`: fast partial reader. It unzips only the needed parts of the workbook with fflate.
- `worker.js`: parses workbooks in a Web Worker.
- `app.js`, `index.html`, `styles.css`: the UI.
