// Add custom calendar menu when spreadsheet is opened
const onOpen = () => {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate calendar events', functionName: 'generateCalendarEvents'}
  ];
  spreadsheet.addMenu('Calendar', menuItems);
};

// Generate iCalendar calendar header
const _getCalendarHeader = (calendarName) => {
  let calendarHeader = "";
  calendarHeader += `BEGIN:VCALENDAR\n`;
  calendarHeader += `PRODID: -//M.P. Nitowski//Calendar Generator//EN\n`;
  calendarHeader += `VERSION:2.0\n`;
  calendarHeader += `CALSCALE:GREGORIAN\n`;
  calendarHeader += `METHOD:PUBLISH\n`;
  calendarHeader += `X-WR-CALNAME:${calendarName}\n`;
  calendarHeader += `X-WR-TIMEZONE:UTC\n`;
  return calendarHeader;
};

const _zeroPad = (n) => n < 10 ? '0' + n : '' + n;
const _formatDate = (d) => `${d.getUTCFullYear()}${_zeroPad(d.getUTCMonth() + 1)}${_zeroPad(d.getUTCDate())}`;
const _formatTime = (d) => `${_zeroPad(d.getUTCHours())}${_zeroPad(d.getUTCMinutes())}${_zeroPad(d.getUTCSeconds())}`;
const _formatUTCDatetime = (d) => _formatDate(d) + 'T' + _formatTime(d) + 'Z';

// Transform row into multiple iCalendar calendar events
const _getCalendarEventsFromRow = (row) => {
  const now = new Date();
  const currentYear = now.getUTCFullYear();

  const startMonth = parseInt(row[0].split('-')[0].split('/')[0], 10);
  const startDay = parseInt(row[0].split('-')[0].split('/')[1], 10);
  let date = new Date(currentYear, startMonth - 1, startDay);

  let calendarEvents = '';
  for (const column of row.slice(1)) {
    const now = new Date();
  
    const nextDate = new Date(date);
    nextDate.setDate(date.getDate() + 1);
  
    let calendarEvent = "";
    calendarEvent += `BEGIN:VEVENT\n`;
    calendarEvent += `DTSTART;VALUE=DATE:${_formatDate(date)}\n`;
    calendarEvent += `DTEND;VALUE=DATE:${_formatDate(nextDate)}\n`;
    calendarEvent += `DTSTAMP:${_formatUTCDatetime(now)}\n`;
    calendarEvent += `UID:${Utilities.getUuid()}\n`;
    calendarEvent += `CLASS:PRIVATE\n`;
    calendarEvent += `CREATED:${_formatUTCDatetime(now)}\n`;
    calendarEvent += `DESCRIPTION:\n`;
    calendarEvent += `LAST-MODIFIED:${_formatUTCDatetime(now)}\n`;
    calendarEvent += `SEQUENCE:0\n`;
    calendarEvent += `SUMMARY:${column}\n`;
    calendarEvent += `TRANSP:TRANSPARENT\n`;
    calendarEvent += `END:VEVENT\n`;

    calendarEvents += calendarEvent;

    date = nextDate;
  }
 
  return calendarEvents;
};

// Get iCalendar calendar footer
const _getCalendarFooter = () => 'END:VCALENDAR';

// Generate iCalendar calendar based on active spreadsheet, and send it to recipient
// Creates new iCalendar file in (possibly newly created) generated_calendars folder in your Google Drive
// Shares file with recipient and emails direct download link with instructions from your current Gmail account
const generateCalendarEvents = () => {
  const calendarName = Browser.inputBox("What would you like to call the calendar?", Browser.Buttons.OK_CANCEL);
  if (calendarName == 'cancel') {
    return;
  }

  const recipient = Browser.inputBox(
    "Who would you like to share the calendar with? (Enter their email address)",
    Browser.Buttons.OK_CANCEL
  );
  if (recipient == 'cancel') {
    return;
  }

  let calendar = "";
  calendar += _getCalendarHeader(calendarName);

  const sheet = SpreadsheetApp.getActive();
  const sheetData = sheet.getDataRange().getValues();

  for (const row of sheetData.slice(1)) {
    calendar += _getCalendarEventsFromRow(row);
  }

  calendar += _getCalendarFooter();

  const filename = `${calendarName}_${_formatUTCDatetime(new Date())}.ical`;
  let folder = Drive.Files.list({q: `title='generated_calendars'`}).items[0];
  if (!folder) {
    folder = Drive.Files.insert({title: 'generated_calendars', mimeType: 'application/vnd.google-apps.folder'});
  }
  const file = Drive.Files.insert({title: filename, mimeType: 'text/calendar', parents: [{id: folder.id}]}, Utilities.newBlob(calendar));
  Drive.Permissions.insert({type: 'user', role: 'reader', value: recipient}, file.id);

  const calendarDownloadLink = `https://docs.google.com/uc?export=download&id=${file.getId()}`;

  GmailApp.sendEmail(
    recipient,
    `New Calendar Download Link for ${calendarName}`,
    `Import calendar events for ${calendarName} with this download link: ${calendarDownloadLink}\n`
    + "You may need to copy and paste the link into Safari to import on iOS."
  );

  Browser.msgBox(
    `Download link to new calendar file created in Google Drive: ${calendarDownloadLink}`
  );
};
