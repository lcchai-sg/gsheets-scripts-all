function config(){
   /////// CONFIG ////////
  // Form
  const sfn = "Arkusz2"; // sheet name with form to fill
  const fr = "b2:b10";   // range in that sheet
  const fcol = [2,1,5,7];  // colummns to fill 
  // (id, name, cash, bag)

  // Data
  const sdn = "Arkusz1";  // sheet name with source data
  const dr = "A2:A18";    // searched area with data
  const dcol = [2, 3, 4]; // colummns with data
  // ( name, cash, bag)
 
  return [sfn, fr, fcol, sdn, dr, dcol];
};

function sort(r, col) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataRange = spreadsheet.getRange(r); // posible range of an id to sort

  dataRange.sort({column: col, ascending: true});
};

function prepareList(sn, r){
  const sourceSheet = SpreadsheetApp.getActive().getSheetByName(sn);  
  const sourceRange = sourceSheet.getRange(r).getValues(); // todays employees list
  const idList = [];

  for (let i=0; i<sourceRange.length;i++){
    if (sourceRange[i][0] == ""){break;}; // empty (last) cell
    idList.push(sourceRange[i][0]);
  };
  return idList;
};

function getData(sn, r, col, ids){
  const sourceSheet = SpreadsheetApp.getActive().getSheetByName(sn);  
  const sourceRange = sourceSheet.getRange(r); // searched area with data
  let results =[];

  for(let n in ids){
    const searchFor = ids[n]; 
    results.push([ids[n]]); // push new array to results list, statring with id

    // searching id
    let found = sourceRange.createTextFinder(searchFor).findNext(); //find id line
    if( found == null){
      let val = "";
      for(let i in col){
        results[n].push(val); // add empty values for not used id
        i++;
        };
      n++;
      continue;  // skip if meets unused id
    };
    
    let k = found.getRow(); // row with searched id
    for(let i in col){
      let val = sourceSheet.getRange(k,col[i]).getValue();
      results[n].push(val);
      i++;
    };
    n++;
  }  
  return results;
};

function fillCells(sn, col, data){
  const targetSheet = SpreadsheetApp.getActive().getSheetByName(sn);
  let i = 2; // starting line //////////config exception
  
  for(let n in data){ // number of copied lines
    for(let x in data[n]){
      targetSheet.getRange(i,col[x]).setValue(data[n][x]);
      x++;
    }
    n++;  
    i++;
  };
};

// Alert pop up
function notEmpty() {
  const title = "Lista pracowników";
  const msg = "Zakres zwróconych pogotowi kasjerskich nie jest pusty. Przygotować listę mimo to?";  

  var result = SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  if(result == SpreadsheetApp.getUi().Button.OK) {
    return true;
  } else {
    return false;
  };
};

// Test new day list
function isEmpty(sn){
  var cont = true;

  //name sheet and range
  const cr = "E2:E10" //exception from config
  
  const sourceSheet = SpreadsheetApp.getActive().getSheetByName(sn);  
  const sourceRange = sourceSheet.getRange(cr).getValues(); 

  //check if range is empty
  for (let x in sourceRange){if(sourceRange[x] != ''){
      // call pop up
      cont = notEmpty()
      break;
      };
  };
  return cont
};

// main function
function filler(){
  cfg = config();

  if(!isEmpty(cfg[3])){return};
  
  sort(cfg[1], cfg[2][0]);

  list = prepareList(cfg[0], cfg[1]);

  data = getData(cfg[3], cfg[4], cfg[5], list);

  fillCells(cfg[0], cfg[2], data);
};
