# OCV Report Builder

https://sangnew.github.io/ocv-report-builder-web/

Turns the **Mass Production E&L Grade OCV Tracking Sheet** workbook into the Frozen IR / Spot report
(the `test1.xlsx` layout: LOT, Cell ID, Grade, dOCV, Frozen IR, voltage drop, spot details, sizes, AZS, DNC).

Anyone with the link can use it. Files are read in the browser and are never uploaded.

## How it works

1. Load the tracking workbook. Only the `Master E & L` sheet is read at first, which is fast even for 30MB+ files.
2. Pick LOTs. By default only cells that have their own OCV tracking sheet are included, which matches the sample report.
3. Review the table:
   - **Yellow**: needs manual input. Shape, Location and the sizes are never in the workbook. Other columns are yellow when the Master row is empty.
   - **Light blue**: taken from the cell's own tracking sheet (dropped layer `R6`, max dOCV `S6`) because the Master row was empty. Please verify these.
   - **Green**: typed by you, or copied from a previous report.
4. Download the Excel file. Highlights are kept in the file, and a `Legend` sheet explains them.

Optionally, load a previous report in the same format to carry over hand-entered values by Cell ID.
Values you type are remembered in the browser (per Cell ID) until you use **Reset → Also forget the values I typed**.

## Column rules

| Report column | Source |
|---|---|
| LOT, Cell ID, Grade, dOCV | Master: Lot ID, Cell ID, Grade, dOCV (mV) |
| Frozen IR Result (35MOhm) | Master: Frozen IR (Mohm) |
| Frozen IR Pass/Fail | Master: Frozen IR Result; if empty, NG when Frozen IR < 35 MΩ, otherwise OK |
| voltage drop / no drop | "NTF" if Master NTF column is NTF; "Drop" if Anode Sheet / Voltage Drop is filled; otherwise the tracking sheet's R6 |
| Voltage Dropped Layer | Master: Anode Sheet (else tracking sheet R6) |
| dOCV (V) | Master: Voltage Drop (V) (else tracking sheet S6) |
| Spot Found | "Spot Found" when Master: Burn mark/Pinhole/None is filled |
| Top/ Back, x, y | Master: Anode Top/Back, X (mm), Y (mm) |
| SEM/EDS Analysis | Master: EDS Impurity Results, else the element in the NTF column (Cu, Ni, Fe, …) |
| Shape, Location, Long side, Short side, Height, AZS, DNC | Manual |

## Files

- `convert.js`: conversion rules. Pure functions, also runnable in Node.
- `xlsx-subset.js`: fast partial reader. It unzips only the needed parts of the workbook with fflate.
- `worker.js`: parses workbooks in a Web Worker.
- `app.js`, `index.html`, `styles.css`: the UI.
