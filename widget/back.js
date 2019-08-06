// Current spreadsheet
var ss = SpreadsheetApp.getActiveSpreadsheet();

function testing() {
  getAllData( '085', '095', 'Sheet7' );
  


}
/**
 * The function parses data from B085 and B095 form and inserts them into dest table.
 * @param {String} B085 Name of the sheet with dispatcher names.
 * @param {String} B095 Name of the sheet with unit numbers.
 * @param {String} dest Name of the sheet to which to import the parsed data.
 * @returns {JSON} Errors and unmatched records.
 */
function getAllData(B085, B095, dest) {
    // parse data from the B085 form
    var B085Output = parseB085( B085 );
    var B085Data = B085Output.data;
    var B085Excluded = B085Output.excluded;

    // parse data from the B095 form
    var trucks = parseB095( B095 );
    Logger.log( trucks );
    // Array with loads, which have not been matched with unit number 
    var withoutUnitNumber = [];
    

    // add to B085 form data unit numbers
    B085Data.forEach( function(record) {
      var loadNumber = record[0];
      
      for ( var i = 0; i < trucks.length; i++ ) {
        if ( trucks[i][0] == loadNumber ) {
          // splice array trucks, so at the end it will be either empty,
          // either some records from B095 will not be matched
          record[2] = trucks[i][1];
          trucks.splice( i, 1 );
          break;
        }
      }
      
      if ( record[2] === null ) withoutUnitNumber.push(record);
    });
   
   
   
   var destinarionTable = ss.getSheetByName( dest );
   
   // if output is not empty, and the first element either
   if ( B085Data.length && B085Data[0].length ) {
   
     // sort by load number
     B085Data.sort( function (a, b) {
       return a[0] - b[0];
     } );
     
     // insert to the destination table
     destinarionTable.getRange( 1, 1, B085Data.length, B085Data[0].length ).setValues( B085Data );
   }
   
  // JSON file with potential problems
   var json = JSON.stringify([
     {
       data: B085Excluded,
       description: 'Those were not taken in account when the B085 form was parsed. Make sure there\'s no valid records'
     },
     {
       data: withoutUnitNumber,
       description: 'Those records were found in B085, but the match was not found in B095'
     },
     {
       data: trucks,
       description: 'Those were found in B095, but not found a match in B085'
     }
   ]);
   
   return json;
}


/**
 * Extracts the data from the B085 form into structured array.
 * @param {String} B085 Name of the sheet with B085 form.
 * @returns {Array} Structured data.
 */
function parseB085(B085) {
    var B085 = ss.getSheetByName( B085 );
    // all data from B085 form
    var formData = B085.getSheetValues( 1, 1, B085.getLastRow() || 1, B085.getLastColumn() || 1 );
    
    // records which do not match the requirements in if statements
    var excludedRecords = [];

    // Array with objects. Each object is added for the particular
    // dispatcher. In this object records are added, and his name
    // is added as a 'disp' property
    var dispatcherRecords = [{
      records: []
    }];


    // aggregate records by dispatcher name
    formData.forEach( function processData(record) {
        if ( String( record[1] ).trim() == 'Invoices:' && record[0] != 'Grand Totals' ) {
            // if a summary record, add the 'disp' property to the current object, 
            // and create a new one with records array inside
            dispatcherRecords[ dispatcherRecords.length-1 ].disp = record[0];

            // create new obj
            var nextDispatcher = {};
            nextDispatcher.records = [];
            dispatcherRecords.push( nextDispatcher );
        
        } else if (record[1] !== ''  && record[2] !== 'Date Range:' && record[0] != 'Grand Totals' ) {
            // if this is a typical record - add them to the 'records' array of the current obj
            dispatcherRecords[dispatcherRecords.length-1].records.push( record );
        } else if ( record[0] != 'Grand Totals' ) {
            // if record is not typical - add the one to the excludedRecords array
            excludedRecords.push( record );
        } 
    } );
    
    // all the records with disp names
    var records = [];
    
    // exract all the records from 'dispatcherRecords' to the
    // records array. Add name of the dispatcher to each record
    dispatcherRecords.forEach( function(dispatcher){
      dispatcher.records.forEach( function(record) {
        records.push( record.concat( dispatcher.disp ) );
      } );
    } );
    
    
    // form the structure of the output, set the unitNumber to null (will be taken)
    // from B095 form. Below the structure of the output:
    // [ loadNumber,  brokerName, null, revenue, loadPayment, miles, profit, dispatcherName ]
    var output = records.map( function(record) {
      var structuredRecord = [];
      
      // RegExp test added not to allow the code to throw an error, if it does not match
      // Record's value should be string, to perform RegExps on it
      if ( /\d{5}/.test( String(record[0]) ) ) {
        structuredRecord.push( String( record[0] ).match(/\d{5}/)[0], record[1], null, record[7], record[9], record[5], record[7] - record[9], record[record.length-1] );
      } else {
        structuredRecord.push( null, null, null, null, null, null, null, null );
        excludedRecords.push( record );
      }
      
      return structuredRecord;
    });
    

    return {
      data: output,
      excluded: excludedRecords
    };
}

  
/**
 * The function extracts load and unit number from B095 form.
 * @param {String} B095 Name of the sheet with the data from B095 the form.
 * @returns {Array} Array with load and unit numbers.
 */
function parseB095(B095) {
    Logger.log( 'here ');
    var B095 = ss.getSheetByName( B095 );
    // all data from B095 form
    var formData = B095.getSheetValues( 1, 1, B095.getLastRow() || 1, B095.getLastColumn() || 1 );

    // extract load number and unit number
    // if unit number is '' or 'Average per Invoice:'
    // then return null
    var loadTruckArray = formData.map( function(record){
      if ( record[13] !== '' && String( record[13] ).trim() !== 'Average per Invoice:    ' && /\d{5}/.test( String(record[0]) ) ) {
        return [ String(record[0]).match(/\d{5}/)[0], record[13] ];
      } else {
        return null;
      }
    } );
    

    // remove records with null (no unit)
    var trucks = loadTruckArray.filter( function(record) {
      if ( record === null ) {
        return false;
      }
      
      return true;
    });


    return trucks;
}



/**
 * The function searches for duplicates/changes between two tables, marks them and sorts
 * @param {String} monthly Name of the sheet with the static records.
 * @param {String} dest Name of the sheet to which the records have just been inserted.
 */
function handleDuplicatesAndDifferences(monthly, dest) {
    var monthly = ss.getSheetByName( monthly );
    var dest = ss.getSheetByName( dest );

    // get data and occurrences
    var monthlyData = monthly.getSheetValues( 2, 4, monthly.getLastRow() || 2, 8 );
    var destData = dest.getSheetValues( 1, 1, dest.getLastRow() || 1, dest.getLastColumn() || 1 );
    var occurrences = searchForDuplicates( monthly, dest );

    // iterate occurrences; if the are a match - status is exists, otherwise - changed
    for ( var i = 0; i < occurrences.length; i++ ) {
        var monthlyIndex = occurrences[i].monthly;
        var destIndex = occurrences[i].dest;

        // Indexes of [Broker, Truck, Load$, Driver$, Miles$ and Dispatcher]
        var indexToCheck = [1, 2, 3, 4, 5, 7];

        for ( var j = 0; j < indexToCheck.length; j++ ) {
          var index = indexToCheck[j];
          var monthlyCell = String( monthlyData[monthlyIndex][index] ).trim(); // trimmed cell in monthly report
          var destCell = String( destData[destIndex][index] ).trim(); // trimmed cell in destination report

          if ( monthlyCell !== destCell ) {
            // if in the record the first field was changed - add indexes array (indexes of changes)
            // and set the status prop of occurrences ot be 'changed'
            if ( !occurrences[i].status ) {
              occurrences[i].status = 'changed';
              occurrences[i].indexes = [];
            }
            
            // add the index of a changed cell
            occurrences[i].indexes.push( index );
          }
        }
    }

    // add the status to the dest sheet
    for ( var i = 0; i < occurrences.length; i++ ) {
        var destIndex = occurrences[i].dest;
        if ( occurrences[i].status === 'changed' ) {

            //iterate all changes
            var indexes = occurrences[i].indexes;
            for ( var j = 0; j < indexes.length; j++ ) {
              dest.getRange( (destIndex + 1), (indexes[j] + 1), 1, 1 ).setBackground( 'yellow' );
            }
        }
    }
}


/**
 * The function searches for duplicates between two sheets.
 * @param {Stirng} monthly Table with records.
 * @param {String} dest Table to which the records were inserted.
 * @returns {Array} Array with indexes of duplicates in sheets.
 */
function searchForDuplicates(monthly, dest) {
    var monthlyData = monthly.getSheetValues( 2, 4, monthly.getLastRow() || 2, 8 );
    var destData = dest.getSheetValues( 1, 1, dest.getLastRow() || 1, dest.getLastColumn() || 1 );
    
    var occurrences = [];
    for ( var i = 0; i < destData.length; i++ ) {
      if ( destData[i][0] === '' ) continue;
      
      for ( var j = 0; j < monthlyData.length; j++ ) {
        if ( destData[i][0] === monthlyData[j][0] ) {
            occurrences.push( {
                monthly: j,
                dest: i
            } );
          break;
        }
      }
    }

    return occurrences;
}