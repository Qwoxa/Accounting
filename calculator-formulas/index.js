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
function initializeVariables(dt, dp, dst, sn) {
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


function Table(startDay, row) {
  this._startDay = startDay;
  this._row = row;
  this._data = [];

  /**
   * Initializes the table.
   * @param {Object} startDay The Monday of the week.
   * @param {Number} row The row to insert the table to.
   * @returns {Array} Two dimentional array with data.
   */
  this.init = function() {
    // create array with days of the week
    this._createDaysArray();
    
    // create header and body
    this._data = new Head( this._weekdays, this._row );
    this._data = this._data.concat( new Body(row) );

    return this._data;
  };

  /**
   * Creates an arrays with dates for the headings
   */
  this._createDaysArray = function() {
    this._weekdays = [];
    
    // no side-effects
    var currentDay = new Date( this._startDay.valueOf() );
    this._weekdays.push( new Date( this._startDay.valueOf() ) );
    
    // add to the this._weekdays the whole week
    while ( this._weekdays.length != 7 ) {
      var nextDay = currentDay.setDate( currentDay.getDate() + 1 );
      this._weekdays.push( new Date( nextDay.valueOf() ) );
    }
  };
}

/**
 * Creates the head of the week report.
 * @param {Array} weekdays Array with the days.
 * @returns {Array} Array with head.
 */
function Head(titles, row) {
  this._titles = titles;
  this._row = row;
  this._lastCol = 16;
  this._data = [['Dispatcher'], [null]];

  /**
   * Initialized the header.
   * @returns {Array} The array with header.
   */
  this.init = function() {
    var datesRow = this._data[0];
    var marksRow = this._data[1];

    this._titles.forEach( function(title) {
      datesRow.push( title.toDateString(), null );
      marksRow.push( 'Q-ty', '$' );
    } );

    datesRow.push( 'Summary', null );
    marksRow.push( 'Q-ty', '$' );

    this._setHeadStyles();
    return this._data;
  };


  /**
   * Sets styles for the header: merging, backgrounds, colors.
   */
  this._setHeadStyles = function() {
    // MERGING
    dest.getRange( this._row, 1, 2, 1 ).mergeVertically(); 
  
    for ( var col = 2, i = 0; i < 8; i++ ) {
      dest.getRange( this._row, col, 1, 2 ).mergeAcross();
      col += 2;
    }

    // BACKGROUNDS AND COLORS
    dest.getRange( row, 2, 2, this._lastCol ).setBackground( headingColor );
    dest.getRange( row, 1, 2, 1 ).setBackground( headingColor );
    dest.getRange( row, 2, 1, this._lastCol - 2 ).setBackground( headingColorLighter );
    dest.getRange( row, 1, 1, 1 ).setFontColor( '#fff' );
    dest.getRange( row, this._lastCol, 1, 1 ).setFontColor( '#fff' );
  };
}



function Body(row) {
  this._tableStartRow = row;
  this._bodyStartRow = row + 2;
  this._totalTeamResults = [];
  this._data = [];

  this.init = function() {
    for ( var i = 0; i < disps.length; i++ ) {
      var team = disps[i];
      
      var teamRecords = this._generateTeamRecords( team );
      var teamRecordsAmount = teamRecords.length;
      this._data = this._data.concat( teamRecords );
      

      styleTeam( this._bodyStartRow, teamRecordsAmount );

      this._bodyStartRow += generatedTeam.length;
      this._totalTeamResults.push( this._bodyStartRow - 1 );
    }
  };

  this._generateTeamRecords = function(team) {

    function qtyFormula(startRow, currentRow, letter) {
      return '=COUNTIFS( ' + sourceName + '!$K:$K, $A' + currentRow + ', \
              ' + sourceName + '!$C:$C, ' + letter + '$' + startRow + ' \
            )';
    }


    function sumFormula(startRow, currentRow, letter) {
      return '=SUMIFS(' + sourceName + '!$J:$J, \
              ' + sourceName + '!$K:$K, $A' + currentRow + ', \
              ' + sourceName + '!$C:$C, ' + letter + '$' + startRow + ' \
            )';
    }

    function qtySummary(row) {
      return "=B".concat(row, "+D").concat(row, "+F").concat(row, "+H").concat(row, "+J").concat(row, "+L").concat(row, "+N").concat(row);
    }
    
    
    function sumSummary(row) {
      return "=C".concat(row, "+E").concat(row, "+G").concat(row, "+I").concat(row, "+K").concat(row, "+M").concat(row, "+O").concat(row);
    }


    var totalSumFormula = function(firstRow, lastRow, letter) {
      return '=sum(' + letter + startRow + ':' + letter + endRow + ')';
    }.bind( null, this._bodyStartRow, this._bodyStartRow + team.length - 1 );

    var teamRecords = [];


    
  for ( var i = 0; i < names.length; i++) {
    var qty = qtyFormula.bind( null, this._tableStartRow, this._bodyStartRow );
    var sum = sumFormula.bind( null, this._tableStartRow, this._bodyStartRow );
    var letters = ['B', 'D', 'F', 'H', 'J', 'L', 'N'];

    letters.forEach( function(letter) {
      teamRecords.push( qty(letter), sum(letter) );
    } );

    teamRecords.push (
      qtySummary( this._bodyStartRow ),
      sumSummary( this._bodyStartRow )
    );

    
    ++this._bodyStartRow;
  }
  
  var summaryCount = "=B".concat(currentRow, "+D").concat(currentRow, "+F").concat(currentRow, "+H").concat(currentRow, "+J").concat(currentRow, "+L").concat(currentRow, "+N").concat(currentRow);
  var summarySum = "=C".concat(currentRow, "+E").concat(currentRow, "+G").concat(currentRow, "+I").concat(currentRow, "+K").concat(currentRow, "+M").concat(currentRow, "+O").concat(currentRow);
  
  
  team.push( ['Total', totalSum( 'B' ), totalSum( 'C' ), totalSum( 'D' ), totalSum( 'E' ), totalSum( 'F' ),
  totalSum( 'G' ), totalSum( 'H' ), totalSum( 'I' ), totalSum( 'J' ), totalSum( 'K' ), totalSum( 'L' ),
  totalSum( 'M' ), totalSum( 'N' ), totalSum( 'O' ), summaryCount, summarySum] );
  
  
  return team;
  };
}



function createBodyff(startRow) {
 
  
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