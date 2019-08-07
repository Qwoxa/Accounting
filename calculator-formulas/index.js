var headingColor = '#e69138';
var headingColorLighter = '#f9cb9c';
var grey = '#d9d9d9';
var darkGrey = '#666666';
var dates,
    disps,
    dest,
    sourceName;



function testGeneration() {
  // list of dispatchers
  var disps = [['Max', 'Mike', 'Daniel', 'Jake', 'Matt', 'James', 'Ryan', 'Dominic', 'David', 'Pete'], 
                ['Ross', 'Ray', 'Todd', 'Oscar', 'Stan'], 
                ['Mark'],
                ['Tony', 'Howard', 'Jack', 'Tom', 'Travis', 'Nate', 'George', 'Dave', 'Sean']];
  
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
  

  return head.concat( body );
}



function createBody(startRow) {
  // get the top of the table
  var row = startRow + 2;
  var totals = [];
  
  var body = [];
  //each team
  for ( var i = 0; i < disps.length; i++ ) {
      var team = disps[i];
      
      var generatedTeam = generateTeam( startRow, row, team );
      body = body.concat( generatedTeam );
      

      styleTeam( row, team.length );

      row += generatedTeam.length;
      totals.push( row - 1 );
      
  }
  
  
  var tableTotal = createTableTotal( totals );
  body = body.concat( tableTotal );
  dest.getRange( row, 1, 1, 17 ).setBackground( darkGrey );
  dest.getRange( row, 1, 1, 17 ).setFontColor( '#fff' );
  
  return body;
}


function createTableTotal(indexes) {
    var A = [], B = [], C = [], D = [], E = [], F = [], G = [],
      H = [], I = [], J = [], K = [], L = [], M = [], N = [], O = [],
      P = [], Q = [];
      
    for ( var i = 0; i < indexes.length; i++ ) {
      A.push( 'A' + indexes[i] );
      B.push( 'B' + indexes[i] );
      C.push( 'C' + indexes[i] );
      D.push( 'D' + indexes[i] );
      E.push( 'E' + indexes[i] );
      F.push( 'F' + indexes[i] );
      G.push( 'G' + indexes[i] );
      H.push( 'H' + indexes[i] );
      I.push( 'I' + indexes[i] );
      J.push( 'J' + indexes[i] );
      K.push( 'K' + indexes[i] );
      L.push( 'L' + indexes[i] );
      M.push( 'M' + indexes[i] );
      N.push( 'N' + indexes[i] );
      O.push( 'O' + indexes[i] );
      P.push( 'P' + indexes[i] );
      Q.push( 'Q' + indexes[i] );
    }
  
    return [[
      'TOTAL',
      '=' + B.join('+'),
      '=' + C.join('+'),
      '=' + D.join('+'),
      '=' + E.join('+'),
      '=' + F.join('+'),
      '=' + G.join('+'),
      '=' + H.join('+'),
      '=' + I.join('+'),
      '=' + J.join('+'),
      '=' + K.join('+'),
      '=' + L.join('+'),
      '=' + M.join('+'),
      '=' + N.join('+'),
      '=' + O.join('+'),
      '=' + P.join('+'),
      '=' + Q.join('+'),
    ]];
}



function styleTeam( currentRow, teamLen) {

  var col = 2;
  for ( var i = 0; i < 7; i++ ) {
    dest.getRange( currentRow, col, teamLen, 1 ).setBackground( grey );
    col += 2;
  }
  
  
  // vertical total
  dest.getRange( currentRow + teamLen, 1, 1, 17 ).setBackground( headingColor );
  dest.getRange( currentRow + teamLen, 1, 1, 17 ).setFontColor( '#fff' );
  
  // summary qty
  dest.getRange( currentRow, col, teamLen, 1 ).setBackground( darkGrey );
  dest.getRange( currentRow, col, teamLen, 1 ).setFontColor( '#fff' );
}

/**
 * Generates data (formulas) on team.
 * @param {Number} startRow The row where the table starts.
 * @param {Number} currentRow The row where the team starts.
 * @returns {Array} Formulas for profit and count accounting.
 */
function generateTeam(startRow, currentRow, names) {
  var currentRow = currentRow;

  var totalSum = totalTeamSumFormula.bind( this, currentRow, currentRow + names.length -1 );
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
      qtySummary( currentRow ),
      sumSummary( currentRow )
    ] );
    
    ++currentRow;
  }
  
  var summaryCount = "=B".concat(currentRow, "+D").concat(currentRow, "+F").concat(currentRow, "+H").concat(currentRow, "+J").concat(currentRow, "+L").concat(currentRow, "+N").concat(currentRow);
  var summarySum = "=C".concat(currentRow, "+E").concat(currentRow, "+G").concat(currentRow, "+I").concat(currentRow, "+K").concat(currentRow, "+M").concat(currentRow, "+O").concat(currentRow);
  
  
  team.push( ['Total', totalSum( 'B' ), totalSum( 'C' ), totalSum( 'D' ), totalSum( 'E' ), totalSum( 'F' ),
  totalSum( 'G' ), totalSum( 'H' ), totalSum( 'I' ), totalSum( 'J' ), totalSum( 'K' ), totalSum( 'L' ),
  totalSum( 'M' ), totalSum( 'N' ), totalSum( 'O' ), summaryCount, summarySum] );
  
  
  return team;
}


function qtySummary(row) {
  return "=B".concat(row, "+D").concat(row, "+F").concat(row, "+H").concat(row, "+J").concat(row, "+L").concat(row, "+N").concat(row);
}


function sumSummary(row) {
  return "=C".concat(row, "+E").concat(row, "+G").concat(row, "+I").concat(row, "+K").concat(row, "+M").concat(row, "+O").concat(row);
}


function totalTeamSumFormula(startRow, endRow, letter) {
  return '=sum(' + letter + startRow + ':' + letter + endRow + ')';
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