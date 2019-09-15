// Keep the New Paid Gigger Starting Value in the Helpers sheet up to date
function onEdit() {
  try{
    var musiciansData = musiciansSheet.getDataRange().getValues();
    var musicianTotals=[]
    // for every musician,
    for (var i=0; i<musiciansData.length; i++){
      // if they're currently eligible for non-charity gigs,
      for (var j=2; j<8; j++){
        if (musiciansData[i][j]=='Y'){
          // include their total in the calculations
          musicianTotals.push(musiciansData[i][10]);
          break;
        }
      }
    }
    musicianTotals.sort();
    // set new paid gigger starting value to the median of the existing musicans' values
    var newPaidGiggerStartingValue = musicianTotals[Math.floor(musicianTotals.length / 2)];
    helperSheet.getRange(newPaidGiggerStartingValueCell).setValue(newPaidGiggerStartingValue);
  } catch(e){
    Logger.log(e);
  }
}
