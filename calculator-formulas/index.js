var headingColor = '#e69138';
var headingColorLighter = '#f9cb9c';
var dates,
    disps,
    dest,
    sourceName;



function testGeneration() {
  // list of dispatchers
  var disps = [['Max', 'Mike', 'Daniel', 'Jake', 'Matt', 'James', 'Ryan', 'Dominic', 'David', 'Pete'], ['Ross', 'Ray', 'Todd', 'Oscar', 'Stan'], ['Mark'], ['Tony', 'Howard', 'Jack', 'Tom', 'Travis', 'Nate', 'George', 'Dave', 'Sean']];
  
  // duration of pay period
  var start = new Date( 2019, 05, 24 );
  var end = new Date( 2019, 06, 21 );
  var dates = { start: start, end: end };
  

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dest = ss.getSheetByName( 'WEEK' );
  var sourceName = 'July_2019';
  

  initializeVariables( dates, disps, dest, sourceName );
  generation();
}


/**
 * 
 * @param {Object} dt Dates of pp.
 * @param {Object} dp List of teams and dispatchers.
 * @param {Object} dst Destination Sheet.
 * @param {String} sn The name of the sheet to link formulas to.
 */
function initializeVariables( dt, dp, dst, sn ) {
    dates = dt;
    disps = dp;
    dest = dst;
    sourceName = sn;
}


function generation() {
  var start = dates.start;
  var end = dates.end;
  var diffWeeks = Math.ceil( ( end - start ) / 604800000 ); // 604800000 - ms in one week
  var weeksList = createWeeksBreakpoints( start, diffWeeks );

  // start adding rows, starting from first row
  var row = 1;
  
  
  for ( var i = 0; i < diffWeeks; i++ ) {

    var table = createTable( weeksList[i], row );

    // set values, borders and adjust row
    dest.getRange( row, 1, table.length, table[0].length ).setValues( table );
    dest.getRange( row, 1, table.length, table[0].length ).setBorder( true, true, true, true, true, true );
    row += table.length + 1;
  }
}


/**
 * Generates a table for the chosen week
 * @param {Object} startDay The Monday of the week.
 * @param {Number} row The row to insert the next table to.
 * @returns {Array} Two dimentional array with data.
 */
function createTable(startDay, row) {
  // create array with days of the week
  var days = createDaysArray( startDay );
  
  // create header for the table and style it
  var head = createHead( days );
  styleHead( row );
  
  var body = createBody( row );
  
  Logger.log( JSON.stringify(body) );
  return head.concat( body );
}



function createBody(startRow) {
  // get the top of the table
  var row = startRow + 2;
  
  
  var body = [];
  //each team
  for ( var i = 0; i < disps.length; i++ ) {
      var team = disps[i];
      
      body = body.concat( generateTeam( startRow, row, team ) );
      
      row += team.length;
  }
  
  
  
  return body;
}


/**
 * Generates data (formulas) on team.
 * @param {Number} startRow The row where the table starts.
 * @param {Number} currentRow The row where the team starts.
 * @returns {Array} Formulas for profit and count accounting.
 */
function generateTeam(startRow, currentRow, names) {
  var currentRow = currentRow;
  var team = [];
  
  for ( var i = 0; i < names.length; i++) {
    team.push( [ 
      names[i], 
      qtyFormula( 'B', startRow, currentRow ),
      sumFormula( 'B', startRow, currentRow ),
      qtyFormula( 'D', startRow, currentRow ),
      sumFormula( 'D', startRow, currentRow ),
      qtyFormula( 'F', startRow, currentRow ),
      sumFormula( 'F', startRow, currentRow ),
      qtyFormula( 'H', startRow, currentRow ),
      sumFormula( 'H', startRow, currentRow ),
      qtyFormula( 'J', startRow, currentRow ),
      sumFormula( 'J', startRow, currentRow ),
      qtyFormula( 'L', startRow, currentRow ),
      sumFormula( 'L', startRow, currentRow ),
      qtyFormula( 'N', startRow, currentRow ),
      sumFormula( 'N', startRow, currentRow ),
      null,
      null
    ] );
    
    ++currentRow;
  }
  
  return team;
}


/**
 * The function generates the formula for quantity.
 * @param {String} letter Column.
 * @param {Number} startRow Row.
 * @param {Number} currentRow Current Row.
 */
function qtyFormula(letter, startRow, currentRow) {
  return '=COUNTIFS( ' + sourceName + '!$K:$K, $A' + currentRow + ', \
           ' + sourceName + '!$C:$C, ' + letter + '$' + startRow + ' \
         )';
}


/**
 * The function generates the formula for sum.
 * @param {String} letter Column.
 * @param {Number} startRow Row.
 * @param {Number} currentRow Current Row.
 */
function sumFormula(letter, startRow, currentRow) {
  return '=SUMIFS(' + sourceName + '!$J:$J, \
          ' + sourceName + '!$K:$K, $A' + currentRow + ', \
          ' + sourceName + '!$C:$C, ' + letter + '$' + startRow + ' \
         )';
}


/**
 * Gives some styling to the head (font weight, merging, font color and bg color)
 * @param {Object} dest The table to which to apply styles.
 * @param {Number} row The row of table start.
 */
function styleHead(row) { 

  // Dispatcher in first col
  dest.getRange( row, 1, 2, 1 ).mergeVertically(); 
  
  // Merge dates two cells into one
  var col = 2;
  for ( var i = 0; i < 8; i++) {
    dest.getRange( row, col, 1, 2 ).mergeAcross();
    col += 2;
  }
  
  // Date format
  dest.getRange( row, 2, 1, 14 )
  
  // Add colors
  dest.getRange( row, 2, 2, 16 ).setBackground( headingColor );
  dest.getRange( row, 1, 2, 1 ).setBackground( headingColor );
  dest.getRange( row, 2, 1, 14 ).setBackground( headingColorLighter );
  
  
  // Bold
  dest.getRange( row, 2, 2, 16 ).setFontWeight( 'bold' );
  dest.getRange( row, 1, 2, 1 ).setFontWeight( 'bold' );
  
  
  // White font colot
  dest.getRange( row, 1, 1, 1 ).setFontColor( '#fff' );
  dest.getRange( row, 16, 1, 1 ).setFontColor( '#fff' );
  
}

/**
 * Creates the head of the week report.
 * @param {Array} weekdays Array with the days.
 * @returns {Array} Array with head.
 */
function createHead(weekdays) {
  var head = [['Dispatcher'],[null]]; // first cell
  
  weekdays.forEach( function(day) {
    head[0].push( day.toDateString(), null );
    head[1].push( 'Q-ty', '$' );
  } );
  
  head[0].push( 'Summary', null );
  head[1].push( 'Q-ty', '$' );
  
  return head;
}


/**
 * Creates an arrays with dates (from startDay to one week ahead)
 * @param {Object} startDay Date to start from.
 * @returns {Array} Array with dates of the week.
 */
function createDaysArray(startDay) {
   var days = [];
    
  // no side-effects
  var currentDay = new Date( startDay.valueOf() );
  days.push( new Date( startDay.valueOf() ) );
  
  
  while ( days.length != 7 ) {
    var nextDay = currentDay.setDate( currentDay.getDate() + 1 );
    days.push( new Date( currentDay.valueOf() ) );
  }
  
  return days;
}


/**
 * Creates a list of Mondays in a pay period.
 * @param {Date} start A day to start the pay period with.
 * @param {Number} amount Amount of weeks in pay period.
 * @returns {Array} List of Mondays.
 */
function createWeeksBreakpoints( start, amount ) {
  var breakpoints = [];
  var date = new Date( start.valueOf() );
  
  while ( breakpoints.length != amount ) {
    breakpoints.push( new Date( date.valueOf() ) );
    date.setDate( date.getDate() + 7 );
  }
  
  return breakpoints;
}