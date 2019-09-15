// Sheets
var availabilitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Poll');
var musiciansSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Musicians');
var helperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Helpers');

// Layout of the Doodle Poll sheet
var monthRowIndex = 3;
var dateRowIndex = 4;
var gigNameRowIndex = 5;
var firstMusicianAvailabilityRowIndex = 6;

// Layout of the Musicians sheet
var nameColIndex = 0;
var melody1ColIndex = 2;
var melody2ColIndex = 3;
var callingColIndex = 4;
var chordColIndex = 5;
var chordAndMelodyColIndex = 6;
var percussionColIndex = 7;
var gradYearColIndex = 8;
var paidColIndex = 9;
var charityColIndex = 10;
var totalColIndex = 11;

// Layout of the Helpers sheet
var newPaidGiggerStartingValueCell = "B2";