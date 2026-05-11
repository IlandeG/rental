/**
 * Google Apps Script: מוסיף 5 שורות חדשות בראש הטבלה בכל שבוע,
 * ממזג את עמודת ה-Week ומציב טווח תאריכים של השבוע הקרוב (ראשון-שבת).
 *
 * הגדרה ראשונית: הריצו פעם אחת את הפונקציה `installWeeklyTrigger`
 * כדי לקבוע הפעלה אוטומטית כל יום ראשון בבוקר.
 */

const SHEET_NAME = 'Sheet1';      // שם הגיליון - לעדכון במידת הצורך
const HEADER_ROWS = 2;            // מספר שורות הכותרת בראש הטבלה
const ROWS_PER_WEEK = 5;          // מספר השורות שמוסיפים בכל שבוע
const WEEK_COLUMN = 1;            // עמודת "Week"
const TOTAL_COLUMNS = 6;          // Week, Operation, Brief, Source Link, Views, Engagement

/**
 * הפונקציה הראשית - מופעלת ע"י הטריגר השבועי.
 */
function addWeeklyRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();

  const insertAtRow = HEADER_ROWS + 1;

  sheet.insertRowsBefore(insertAtRow, ROWS_PER_WEEK);

  const weekRange = sheet.getRange(insertAtRow, WEEK_COLUMN, ROWS_PER_WEEK, 1);
  weekRange.merge();
  weekRange.setHorizontalAlignment('center');
  weekRange.setVerticalAlignment('middle');
  weekRange.setFontWeight('bold');
  weekRange.setValue(getUpcomingWeekLabel());

  const fullRange = sheet.getRange(insertAtRow, 1, ROWS_PER_WEEK, TOTAL_COLUMNS);
  fullRange.setBorder(true, true, true, true, true, true);
}

/**
 * מחזיר טווח תאריכים של השבוע הקרוב בפורמט "D-D.M" (לדוגמה: "10-16.5").
 * אם השבוע חוצה חודשים, הפורמט יהיה "D.M-D.M" (לדוגמה: "26.4-2.5").
 */
function getUpcomingWeekLabel(referenceDate) {
  const today = referenceDate || new Date();
  const dayOfWeek = today.getDay(); // 0 = ראשון
  const daysUntilSunday = (7 - dayOfWeek) % 7 || 7;

  const sunday = new Date(today);
  sunday.setDate(today.getDate() + daysUntilSunday);
  sunday.setHours(0, 0, 0, 0);

  const saturday = new Date(sunday);
  saturday.setDate(sunday.getDate() + 6);

  const startDay = sunday.getDate();
  const startMonth = sunday.getMonth() + 1;
  const endDay = saturday.getDate();
  const endMonth = saturday.getMonth() + 1;

  if (startMonth === endMonth) {
    return `${startDay}-${endDay}.${endMonth}`;
  }
  return `${startDay}.${startMonth}-${endDay}.${endMonth}`;
}

/**
 * מתקין טריגר שבועי - הריצו פעם אחת בלבד.
 * הטריגר ירוץ כל יום ראשון בין 06:00 ל-07:00.
 */
function installWeeklyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'addWeeklyRows') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('addWeeklyRows')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(6)
    .create();
}
