/**
 * Special type to transmit data name of the sheet, and the index of the specific cells
 */
type MonthData = {
    readonly sheetName : string;
    readonly totalTime : string;
    readonly targetTime : string;
    readonly overtime : string;
    readonly vacationDays : string;
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
 * @param dateStr the name of date
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
 * @param monthData the specific MonthData that shall be appended to the overview sheet
 */
function updateOverview(monthData : MonthData) : void {
    const mainSheet : GoogleAppsScript.Spreadsheet.Sheet | null = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview");
    const prefix : string = '=\'' + monthData.sheetName + '\'!';

    if (mainSheet) {
        mainSheet.appendRow([monthData.sheetName, prefix + monthData.totalTime, prefix + monthData.targetTime, prefix + monthData.overtime, prefix + monthData.vacationDays]);
    } else {
        throw new Error('Please make sure you have a main-sheet named "Overview"')
    }
}

/**
 * Handles everything there is to do for populating the new sheet. This includes the header,
 * dates, weekdays, formulas, data validation rules, formatting, and more.
 * @param date the date of the given month/sheet
 * @param sheetName the name of the given month/sheet
 * @returns The MonthData for the month/sheet
 */
function addNewMonth(date : Date, sheetName : string) : MonthData {
    const sheet : GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSheet();
    const month : number = date.getMonth();
    let row : number = 2;

    // Independently from the picked date, the day of the month is set to the first.
    date.setDate(1);

    // Sheet header
    sheet.appendRow(['Date', 'Weekday', 'Start time', 'End time', 'Additional break /\nInterruption', 'Work time', 'Vacation', 'Sick day', 'Holidays', 'Comments']);

    // Append new rows, for every day of the month
    while (date.getMonth() === month) {
        let dayName : string = getDayName(date, "en-EN");
        sheet.appendRow([date, 
                         dayName, 
                         '', 
                         '', 
                         '',
                         `=IF(D${row}-C${row}-E${row} > TIMEVALUE("06:00:00"); IF(D${row}-C${row}-E${row} > TIMEVALUE("09:00:00"); D${row}-C${row}-E${row} - TIMEVALUE("00:45:00"); D${row}-C${row}-E${row} - TIMEVALUE("00:30:00")); D${row}-C${row}-E${row})`,
                         '',
                         '']);
        
        // Weekend days have a grey backdrop
        if (dayName === 'Saturday' || dayName === 'Sunday') {
            sheet.getRange(row, 1, 1, 10).setBackgroundRGB(200, 200, 200);
        }
        
        // Holidays have a blue backdrop and the name of the holiday in the respective column
        let holidayName = isGermanHoliday(date);
        if (holidayName !== null) {
            sheet.getRange(row, 1, 1, 10).setBackgroundRGB(200, 200, 240);
            sheet.getRange(row, 9).setValue(holidayName);
        }
        
        row++;
        date.setDate(date.getDate() + 1);
    }

    // Add a data validation rule that requires the use of proper date formats for the
    // "Start time" and "End time" columns. 
    const requireDateRule : GoogleAppsScript.Spreadsheet.DataValidation = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
    sheet.getRange(`C2:E${row-1}`).setDataValidation(requireDateRule);

    const checkboxDataRule : GoogleAppsScript.Spreadsheet.DataValidation = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
    sheet.getRange(`H2:H${row-1}`).setDataValidation(checkboxDataRule);

    const dropDownDataRule : GoogleAppsScript.Spreadsheet.DataValidation = SpreadsheetApp.newDataValidation().requireValueInList(['Half', 'Full']).setAllowInvalid(false).build();
    sheet.getRange(`G2:G${row-1}`).setDataValidation(dropDownDataRule);

    sheet.appendRow([' ']);
    sheet.appendRow(['', 'Total working time', `=SUMIF(G2:G${row - 1}; "<>Full"; F2:F${row - 1})`]);
    sheet.appendRow(['', 'Target time', `=MULTIPLY(TIMEVALUE("08:00:00"); COUNTIFS(B2:B${row - 1}; "<>Sunday"; B2:B${row - 1}; "<>Saturday"; G2:G${row - 1}; "="; H2:H${row - 1}; "=FALSE"; I2:I${row - 1}; "=") + (COUNTIFS(G2:G${row - 1}; "=Half") * 0,5))`]);

    row = sheet.getLastRow();
    sheet.appendRow(['', 'Overtime', `=C${row - 1}-C${row}`]);
    sheet.appendRow(['', 'Vacation Days', `=COUNTIFS(G2:G${row - 1}; "=Full"; H2:H${row - 1}; "=FALSE") + COUNTIFS(G2:G${row - 1}; "=Half") * 0,5`]);

    // Enforce date and time specific formats, for the "Date" coulmn and the lower block
    sheet.getRange(`C${row - 1}:C${row + 1}`).setNumberFormat('[hhh]:mm');
    sheet.getRange(`A2:A${row}`).setNumberFormat('dd-mm-yyyy');

    // Set font of the sheet header and lower block to bold.
    sheet.getRange(1, 1, 1, 10).setFontWeight("bold");
    sheet.getRange(`B${row - 1}:B${row + 2}`).setFontWeight("bold");

    // Resize the columns to fit all the data without breaking out of boundaries
    sheet.autoResizeColumn(2);
    sheet.autoResizeColumn(5);
    sheet.autoResizeColumn(9);

    const retVal : MonthData = { sheetName: sheetName, 
                                 totalTime: `C${row - 1}`, 
                                 targetTime: `C${row}`,
                                 overtime: `C${row + 1}`,
                                 vacationDays: `C${row + 2}` };

    return retVal;
}

/**
 * Returns the name of any given holidays, for a specific date, if there are any.
 * @param date the date for which to check if there are holidays
 * @returns The name of any given holidays, for a specific date, if there are any.
 */
function isGermanHoliday(date : Date) : string | null {
    let cal = CalendarApp.getCalendarById('de.german.official#holiday@group.v.calendar.google.com');
    let holidays : GoogleAppsScript.Calendar.CalendarEvent[] = cal.getEventsForDay(date);

    if (holidays.length == 0) return null;

    if (holidays[0].getDescription().includes("Baden-Württemberg") || holidays[0].getDescription() === "Gesetzlicher Feiertag") return holidays[0].getTitle();
    return null;
}

/**
 * Returns the name of a weekday for any given date.
 * @param date the date of the given day
 * @param locale localisation region
 * @returns Name of the day
 */
function getDayName(date : Date, locale : string) : string {
    return date.toLocaleDateString(locale, { weekday: 'long' });
}

/**
 * Returns the name of a sheet, composed of the month (in text form), and the year
 * (in numeric form).
 * @param date the date of the given sheet
 * @param locale localisation region
 * @returns Name of the sheet composed of the month (in text form), and the year (in numeric form).
 */
function getSheetName(date : Date, locale: string) : string {
    return date.toLocaleDateString(locale, {year: 'numeric', month: 'long'});
}

/**
 * Gets called when the document is opened.
 * @param e a related event
 */
function onOpen(e : any) : void {
    const ui : GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();
    ui.createMenu('Timeplan')
      .addItem('Add New Month', 'newMonth')
      .addToUi();
}