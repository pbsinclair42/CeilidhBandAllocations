function onEdit() {
  var helperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Helpers');
  var musiciansData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Musicians').getDataRange().getValues();
  var musicianTotals=[]
  for (var i=0; i<musiciansData.length; i++){
    for (var j=2; j<8; j++){
      if (musiciansData[i][j]=='Y'){
        musicianTotals.push(musiciansData[i][10]);
        break;
      }
    }
  }
  musicianTotals.sort();
  var median = musicianTotals[Math.floor(musicianTotals.length / 2)];
  helperSheet.getRange(2,2).setValue(median);
}
