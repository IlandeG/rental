# rental

## Weekly highlights auto-updater

`apps-script/Code.gs` is a Google Apps Script that adds a new weekly block
(merged "Week" cell + 6 empty rows) to the `highlights` tab of the spreadsheet
every Sunday.

### Install

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Replace the default `Code.gs` with `apps-script/Code.gs` from this repo.
3. Open `appsscript.json` (gear icon → "Show appsscript.json") and replace it
   with `apps-script/appsscript.json` so the timezone is `Asia/Jerusalem`.
4. In the editor, select the function `installWeeklyTrigger` and click **Run**.
   Grant the permissions Google asks for. This schedules `addWeeklyRows` to
   run every Sunday at ~06:00 Israel time.
5. (Optional) Select `addWeeklyRows` and click **Run** once to see a new block
   appear immediately.

### What the trigger does

- Scans column A of the `highlights` tab to find the latest known week
  (labels like `3-9.5` or `26.4-2.5`).
- Computes the following Sun–Sat range.
- Inserts 6 fresh rows, merges column A, writes the new week label, and
  draws borders on the block.
- If the latest block is already the upcoming week, it does nothing.

### Configuration

Edit the `CONFIG` object at the top of `Code.gs`:

| Key              | Meaning                                       |
| ---------------- | --------------------------------------------- |
| `sheetName`      | Tab name (default `highlights`)               |
| `rowsPerWeek`    | Number of blank rows per week block (default 6) |
| `triggerWeekday` | Day the trigger fires (default Sunday)        |
| `triggerHour`    | Hour the trigger fires (default 06:00)        |
| `weekStartsOn`   | 0 = Sunday, 1 = Monday, …                     |
