var availabilitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Poll');
var musiciansSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Musicians');

function onOpen() {
  refreshMenu()
}

function refreshMenu(){
  var ui = SpreadsheetApp.getUi();
  var allocationsMenu = ui.createMenu('Allocate...');
  var doodleData = availabilitySheet.getDataRange().getValues();
  var isUnallocatedCeilidhs = false;
  for (var i=1; i<doodleData[4].length; i++){
    if (/^\d+:\d+:\d+$/.test(doodleData[doodleData.length-1][i])){// only include non-allocated ceilidhs
      allocationsMenu.addItem(getCeilidhName(i), "allocateCeilidh"+i)
      isUnallocatedCeilidhs = true;
    }
  }
  if (isUnallocatedCeilidhs){
    allocationsMenu.addSeparator()
                   .addItem("All", "allocateAll")
  } else {
    allocationsMenu.addItem("All ceilidhs allocated!", "refreshMenu")
  }
  ui.createMenu('Allocations')
      .addSubMenu(allocationsMenu)
      .addSeparator()
      .addItem('Refresh', 'refreshMenu')
      .addToUi();
}

for (var i=1; i<availabilitySheet.getDataRange().getValues()[4].length; i++){
  eval("function allocateCeilidh"+i+"(){return allocateCeilidh("+i+")}");
}

function getMusicians(isCharity, silent){
  var doodleData = availabilitySheet.getDataRange().getValues();
  var musiciansData = musiciansSheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();
  var musicianIndices = {};
  var musicians = {};
  for (var i=6; i<doodleData.length-1; i++){
    var name = doodleData[i][0];
    if (name.indexOf('(')>0){
      name = name.slice(0, name.indexOf('(')-1);
    }
    musicianIndices[name] = i;
  }
  for (var i=1; i<musiciansData.length; i++){
    var musician = {"name": musiciansData[i][0],
                    "melody1": musiciansData[i][2]=='Y' || (isCharity && musiciansData[i][2]=='C'),
                    "melody2": musiciansData[i][3]=='Y' || (isCharity && musiciansData[i][3]=='C'),
                    "calling": musiciansData[i][4]=='Y' || (isCharity && musiciansData[i][4]=='C'),
                    "chord": musiciansData[i][5]=='Y' || (isCharity && musiciansData[i][5]=='C'),
                    "chordAndMelody": musiciansData[i][6]=='Y' || (isCharity && musiciansData[i][6]=='C'),
                    "percussion": musiciansData[i][7]=='Y' || (isCharity && musiciansData[i][7]=='C'),
                    "paid": musiciansData[i][8],
                    "charity": musiciansData[i][9],
                    "total": musiciansData[i][10],
                    "doodleIndex": musicianIndices[musiciansData[i][0]],
                    "musicianDataIndex": i
                   };
    
    if (!silent && musician.doodleIndex===undefined){
      if (ui.alert(musician.name+' not on Doodle poll.  Continue? ', ui.ButtonSet.YES_NO)==ui.Button.NO){
        return undefined;
      }
    } else {
      musicians[musicianIndices[musician.name]] = musician;
      delete musicianIndices[musician.name];
    }
  }
  for (var name in musicianIndices){
    if (!silent && ui.alert(name+' on Doodle poll but not musician info sheet.  Continue? ', ui.ButtonSet.YES_NO)==ui.Button.NO){
      return undefined;
    } else {
      var musician = {"name": name,
                      "melody1": false,
                      "melody2": false,
                      "calling": false,
                      "chord": false,
                      "chordAndMelody": false,
                      "percussion": false,
                      "paid": 100000,
                      "charity": 10000,
                      "total": 90000,
                      "doodleIndex": musicianIndices[name],
                      "musicianDataIndex": -1
                     };
      musicians[musicianIndices[musician.name]] = musician;
    }
  }
  return musicians;
}

function allocateAll(){
  var doodleData = availabilitySheet.getDataRange().getValues();
  var musicians = getMusicians(); //check that the musicians database is correct
  if (musicians===undefined){
    return;
  }
  for (var i=1; i<doodleData[4].length; i++){
    if (/^\d+:\d+:\d+$/.test(doodleData[doodleData.length-1][i])){// only allocate non-allocated ceilidhs
      var cancelled = ! allocateCeilidh(i, true);
      if (cancelled){
        refreshMenu();
        return;
      }
    }
  }
  refreshMenu();
}

function allocateCeilidh(col, noUI) {
  var doodleData = availabilitySheet.getDataRange().getValues();
  var numMusicians = doodleData[5][col].indexOf('(4)')>=0 ? 4 : 3;
  var isCharity = doodleData[5][col].indexOf('(c)')>=0;
  
  var musicians = getMusicians(isCharity, noUI);
  if (musicians===undefined){
    return; // if there's an error in the musician database, don't allocate
  }
  var available = [];
  var ifNeedBe = [];
  for (var i=6; i<doodleData.length-1; i++){
    var musician = musicians[i];
    if (doodleData[i][col]=="OK"){
      musician.isIfNeedBe=false;
      available.push(musician);
    } else if (doodleData[i][col]=="(OK)"){
      musician.isIfNeedBe=true;
      available.push(musician);
    }
  }
  
  var bands = getBestBands(available, numMusicians, isCharity)
  var ui = SpreadsheetApp.getUi();
  var tryAgain=ui.Button.YES;
  while (tryAgain==ui.Button.YES){
    for (var i=0; i<bands.length; i++){
      
      var result = ui.alert(makeAllocationString(col, bands[i]), 'Accept allocation?', ui.ButtonSet.YES_NO_CANCEL);
      if (result == ui.Button.YES) {
        assignAllocation(col, bands[i], isCharity);
        if (noUI!==true){
          refreshMenu();
        }
        return true;
      } else if (result == ui.Button.CANCEL){
        return false;
      }
    }
    tryAgain = ui.alert("No more possible bands for "+getCeilidhName(col)+".  Try again?", ui.ButtonSet.YES_NO_CANCEL);
    if (tryAgain==ui.Button.CANCEL){
      return false;
    }
  }
  return true;
}

function assignAllocation(col, band, isCharity){
  for (var i=7; i<availabilitySheet.getDataRange().getValues().length; i++){
    availabilitySheet.getRange(i,col+1).setValue("");
  }
  for (var part in band){
    for (var i=0; i<band[part].length; i++){
      var player = band[part][i];
      availabilitySheet.getRange(player.doodleIndex+1,col+1).setValue(part);
      if (isCharity){
        musiciansSheet.getRange(player.musicianDataIndex+1, 10).setValue(player.charity+1);
        player.charity+=1;
        player.total-=1;
      }else{
        musiciansSheet.getRange(player.musicianDataIndex+1, 9).setValue(player.paid+1);
        player.paid+=1;
        player.total+=1;
      }
    }
  }
  
  var allocationString = makeAllocationString(col, band);
  availabilitySheet.getRange(availabilitySheet.getDataRange().getValues().length, col+1).setValue(allocationString);
}

function makeAllocationString(col,band){
  var ceilidhName = getCeilidhName(col);
  var players = "";
  for (var part in band){
    players+=part + " - " + band[part].map(function(x){return x.name}).join(', ') + '; ';
  }
  return ceilidhName + ": " + players;
}

function getCeilidhName(col){
  var monthCell = availabilitySheet.getRange(4,col+1);
  var month = (monthCell.isPartOfMerge() ? monthCell.getMergedRanges()[0].getCell(1, 1) : monthCell).getValue();
  month = month.slice(0, month.indexOf(' '));
  var availabilityData = availabilitySheet.getDataRange().getValues();
  var name = availabilityData[5][col];
  if (name.indexOf('(4)')>0){
    name = name.slice(0, name.indexOf('(')-1);
  } else if (name.indexOf('(c)')>0){
    name = name.slice(0, name.indexOf('(')-1);
    name+=" (Charity)"
  }
  return availabilityData[4][col] + " " + month + " (" + name + ")";
}

function getBestBands(available, numMusicians, isCharity, maxNumberOfPossibileBandsReturned){
  var availableGroups = shuffle(getSubsets(available, numMusicians));
  if (isCharity){
    availableGroups.sort(function(groupA, groupB){
      // prioritise first groups with less ifNeedBe members, then groups who between them have played the least charity gigs, then groups who between them have played the least overall gigs
      return groupA.map(function(x){return x.charity + (x.isIfNeedBe ? 1000 : 0) + x.total/1024}).reduce(function(a,b){return a+b}) - groupB.map(function(x){return x.charity + (x.isIfNeedBe ? 1000 : 0) + x.total/1024}).reduce(function(a,b){return a+b})
    });
  }else{
    availableGroups.sort(function(groupA, groupB){
      // prioritise first groups with less ifNeedBe members, then groups who between them have played the least gigs
      return groupA.map(function(x){return x.total + (x.isIfNeedBe ? 1000 : 0)}).reduce(function(a,b){return a+b}) - groupB.map(function(x){return x.total + (x.isIfNeedBe ? 1000 : 0)}).reduce(function(a,b){return a+b})
    });
  }
  var possibleBands = [];
  for (var i=0; i<availableGroups.length; i++){
    var band = getValidBand(availableGroups[i])
    if (band!==undefined){
      possibleBands.push(band);
      if (maxNumberOfPossibileBandsReturned && possibleBands.length >= maxNumberOfPossibleBandsReturned){
        return possibleBands;
      }
    }
  }
  return possibleBands;
}

function getValidBand(band){
  // every band needs a caller
  if (band.find(function(x){return x.calling})===undefined){
    return undefined;
  }
  if (band.length==3){
    //must be one melody I, one melody II or chord+melody, and one chord; or one melody I, one chords+melody, one percussion
    var possibleBands = shuffle(permutator(band));
    for (var i=0; i<6; i++){
      var thisBand = possibleBands[i];
      if (thisBand[0].melody1 && (thisBand[1].melody2 || thisBand[1].chordAndMelody) && thisBand[2].chord){
        return {'Melody': [thisBand[0], thisBand[1]], 'Chords': [thisBand[2]]}
      }
      if (thisBand[0].melody1 && thisBand[1].chordAndMelody && thisBand[2].percussion){
        return {'Melody': [thisBand[0]], 'Melody + Chords': [thisBand[1]], 'Percussion': [thisBand[2]]}
      }
    }
  } else if (band.length==4){
    //must be one melody I, two melody II or chord+melody, and one chord; or one melody I, one melody II or chords+melody, one chord, and one percussion
    var possibleBands = permutator(band);
    for (var i=0; i<24; i++){
      var thisBand = possibleBands[i];
      if (thisBand[0].melody1 && (thisBand[1].melody2 || thisBand[1].chordAndMelody) && (thisBand[2].melody2 || thisBand[2].chordAndMelody) && thisBand[3].chord){
        return {'Melody': [thisBand[0], thisBand[1], thisBand[2]], 'Chords': [thisBand[3]]}
      }
      if (thisBand[0].melody1 && (thisBand[1].melody2 || thisBand[1].chordAndMelody) && thisBand[2].chord && thisBand[3].percussion){
        return {'Melody': [thisBand[0], thisBand[1]], 'Chords': [thisBand[2]], 'Percussion': [thisBand[3]]}
      }
    }
  }
  return undefined;
}

function getSubsets(superset, n){
  result = []
  _getSubsets(superset, n, 0, [], 0, result);
  return result;
}

function _getSubsets(superset, n, previous, indices, nestingLevel, result){
  if (nestingLevel<n){
    for (var i=previous; i<superset.length; i++){
      _getSubsets(superset, n, i+1, indices.concat([i]), nestingLevel+1, result);
    }
  } else {
    var toReturn = []
    for (var i=0; i<n; i++){
      toReturn.push(superset[indices[i]])
    }
    result.push(toReturn);
  }
}

function permutator(inputArr) {
  var results = [];

  function permute(arr, memo) {
    var cur, memo = memo || [];

    for (var i = 0; i < arr.length; i++) {
      cur = arr.splice(i, 1);
      if (arr.length === 0) {
        results.push(memo.concat(cur));
      }
      permute(arr.slice(), memo.concat(cur));
      arr.splice(i, 0, cur[0]);
    }

    return results;
  }

  return permute(inputArr);
}

function shuffle(a) {
    var j, x, i;
    for (i = a.length - 1; i > 0; i--) {
        j = Math.floor(Math.random() * (i + 1));
        x = a[i];
        a[i] = a[j];
        a[j] = x;
    }
    return a;
}