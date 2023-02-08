/**
 * Special type to transmit data name of the sheet, and the index of the specific cells
 */
type MonthData = {
    readonly sheetName : string;
    readonly totalTime : string;
    readonly targetTime : string;
    readonly overtime : string;
}

/**
 * API-esque function that starts the process of setting up a new sheet. Locks other
 * Scripts from running, while active.
 */
function newMonth() : void {
    LockService.getScriptLock().tryLock(1000000);
    datePicker();
    LockService.getScriptLock().releaseLock();
}

/**
 * Loads the HtmlService for DatePicker.html into a prompt, so the user can pick for which
 * month and year, the sheet is intended. 
 */
function datePicker() : void {
    const html = HtmlService.createHtmlOutputFromFile('DatePicker')
                            .setWidth(500)
                            .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'Please pick the month you want to record!');
}

/**
 * Gets called by the DatePicker.html through a Script Runner. Inserts a new sheet and then
 * populates said sheet.
 */
function updateSheets(dateStr : string) : void {
    const date : Date = new Date(dateStr);
    const dateName : string = getSheetName(date, "en-EN");

    SpreadsheetApp.getActiveSpreadsheet().insertSheet(dateName);

    updateOverview(addNewMonth(date, dateName));
}

/**
 * Updates the overview by appending a new row for the newly added sheet, with references to
 * all the relevant data on the other sheet.
 */
function updateOverview(monthData : MonthData) : void {
    const mainSheet : GoogleAppsScript.Spreadsheet.Sheet | null = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview");
    const prefix : string = '=\'' + monthData.sheetName + '\'!';

    if (mainSheet) {
        mainSheet.appendRow([monthData.sheetName, prefix + monthData.totalTime, prefix + monthData.targetTime, prefix + monthData.overtime]);
    } else {
        throw new Error('Please make sure you have a main-sheet named "Overview"')
    }
}

/**
 * Handles everything there is to do for populating the new sheet. This includes the header,
 * dates, weekdays, formulas, data validation rules, formatting, and more.
 */
function addNewMonth(date : Date, sheetName : string) : MonthData {
    const sheet : GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSheet();
    const month : number = date.getMonth();
    let row : number = 2;

    // Independently from the picked date, the day of the month is set to the first.
    date.setDate(1);

    // Sheet header
    sheet.appendRow(['Date', 'Weekday', 'Start time', 'End time', 'Work time']);

    // Append new rows, for every day of the month
    while (date.getMonth() === month) {
        let dayName : string = getDayName(date, "en-EN");
        sheet.appendRow([date, 
                         dayName, 
                         '', 
                         '', 
                         `=IF(MINUS(D${row};C${row}) > TIMEVALUE("06:00:00"); IF(MINUS(D${row};C${row}) > TIMEVALUE("09:00:00"); MINUS(D${row};C${row}) - TIMEVALUE("00:45:00"); MINUS(D${row};C${row}) - TIMEVALUE("00:30:00")); MINUS(D${row};C${row}))`]);
        
        // Weekend days have a grey backdrop
        if (dayName === 'Saturday' || dayName === 'Sunday') {
            sheet.getRange(row,1,1,5).setBackgroundRGB(120,120,120);
        }
        
        row++;
        date.setDate(date.getDate() + 1);
    }

    // Add a data validation rule that requires the use of proper date formats for the
    // "Start time" and "End time" columns. 
    const requireDateEule : GoogleAppsScript.Spreadsheet.DataValidation = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
    sheet.getRange(`C2:D${row-1}`).setDataValidation(requireDateEule);

    sheet.appendRow([' ']);
    sheet.appendRow(['', 'Total working time', `=SUM(E2:E${row - 1})`]);
    sheet.appendRow(['', 'Target time', `=MULTIPLY(TIMEVALUE("08:00:00"); COUNTIFS(B2:B${row - 1}; "<>Sunday"; B2:B${row - 1}; "<>Saturday"))`]);

    row = sheet.getLastRow();
    sheet.appendRow(['', 'Overtime', `=C${row - 1}-C${row}`]);

    // Enforce date and time specific formats, for the "Date" coulmn and the lower block
    sheet.getRange(`C${row - 1}:C${row + 1}`).setNumberFormat('[hhh]:mm');
    sheet.getRange(`A2:A${row}`).setNumberFormat('dd-mm-yyyy');

    // Set font of the sheet header and lower block to bold.
    sheet.getRange(1,1,1,5).setFontWeight("bold");
    sheet.getRange(`B${row - 1}:B${row + 1}`).setFontWeight("bold");

    const retVal : MonthData = { sheetName: sheetName, 
                                 totalTime: `C${row - 1}`, 
                                 targetTime: `C${row}`,
                                 overtime: `C${row + 1}` };

    return retVal;
}

/**
 * Returns the name of a weekday for any given date.
 */
function getDayName(date : Date, locale : string) : string {
    return date.toLocaleDateString(locale, { weekday: 'long' });        
}

/**
 * Returns the name of a sheet, composed of the month (in text form), and the year
 * (in numeric form).
 */
function getSheetName(date : Date, locale: string) : string {
    return date.toLocaleDateString(locale, {year: 'numeric', month: 'long'});
}

/**
 * Gets called when the document is opened.
 */
function onOpen() : void {
    const ui : GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();
    ui.createMenu('Timeplan')
      .addItem('Add New Month', 'newMonth')
      .addToUi();
}
