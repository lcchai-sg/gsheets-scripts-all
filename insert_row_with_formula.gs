function insert_row() {
  var spreadsheet = SpreadsheetApp
			.getActive();
  
  var sheet = spreadsheet
		.getActiveSheet();
  
  var number = sheet
		.getActiveCell()
  		.getValue();
  
  sheet
  	.getRange(
  		spreadsheet
		.getCurrentCell()
		.getRow(), 1, 1, sheet
		.getMaxColumns()
	).activate();
  
  spreadsheet
  	.getActiveSheet()
    	.insertRowsAfter(
    		spreadsheet
		    .getActiveRange()
		    .getLastRow(), number-1
	 );
  
  spreadsheet
  	.getActiveRange()
    .offset(
    	spreadsheet
	    .getActiveRange()
	    .getNumRows(), 0, number-1,
    		spreadsheet
		.getActiveRange()
		.getNumColumns()
	).activate();
  
  spreadsheet
    .getActiveRangeList()
    .clear({
      	contentsOnly: true,
        skipFilteredRows: true
    	});
  
  spreadsheet
    .getCurrentCell()
    .offset(0, 25)
    .activate();
};

function rowstoColumnsformula(){
	var spreadsheet = SpreadsheetApp
				.getActive();
	var row = spreadsheet
			.getActiveCell()
			.getRow()-1;
	var structure =
		'AF'+row+":"+'AK'+row+";"+
		'AL'+row+":"+'AQ'+row+";"+
		'AR'+row+":"+'AW'+row+";"+
		'AX'+row+":"+'BC'+row+";"+
		'BD'+row+":"+'BI'+row+";"+
		'BJ'+row+":"+'BO'+row+';'+
		'BP'+row+":"+'BU'+row;
	return spreadsheet
		.getActiveCell()
		.setFormula(
		'='+'{'+structure+'}'
	);};
