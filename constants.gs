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
var harpColIndex = 7;
var percussionColIndex = 8;
var gradYearColIndex = 9;
var paidColIndex = 10;
var charityColIndex = 11;
var totalColIndex = 12;

// Layout of the Helpers sheet
var newPaidGiggerStartingValueCell = "B2";