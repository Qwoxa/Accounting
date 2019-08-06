function RenderCalculator(options) {


    // Деструктуризация параметров
    var calculatorPage = options.page;
    var dispatchers = options.dispatchers;
    var data = options.data;
    var changedDate = options.changedDate || '';
    var labels = ['Day', 'Week', 'Pay Period'];                                                    // Из-за того, что свойства периодов написаны в CamelCase, labels используются 
                                                                                                   // для проставления адекватных названий периодов
    
    
    (function clearPage() {
      calculatorPage.clearContents();
      calculatorPage.clearFormats();
    })();
    
    
    /**
     * Вставляет текущую дату, колонку Check profit for date
     */
    (function displayInfoPanel() {
      calculatorPage.getRange( 2, 1, 2, 2 )
        .setValues([['Today:', null],
                    ['Check profit for date:', changedDate]])
        .setFontWeight('bold');
      
      calculatorPage.getRange( 2, 2, 1, 1 )
        .setFormula( '=today()' );
                    
      calculatorPage.getRange( 2, 1, 2, 1 )
        .setHorizontalAlignment('right');
    })();
    
    
    /**
     * Размещает объект data на странице calculatorPage
     */
    (function showStatistics() {
      var col = 2;                                                                                 // начиная со второй колонки (первая под имена)
      var dispatchersArr = dispatchers.map(function(disp) {return [disp]});                        // сформировать двумерный массив с именами диспетчеров
      dispatchersArr.unshift(['Total']);                                                           // Добавить Тотал к списку диспетчеров
      dispatchersArr.reverse();                                                                    // Total - в конец (это делается потому, что такая структура у data)
      
      
      calculatorPage.getRange( 6, 1, dispatchersArr.length, 1 ).setValues( dispatchersArr );       // выставить в первой колонке имена диспетчеров
      var periods = Object.keys( data );                                                           // periods - это пэй период, неделя, день.
      periods.reverse();                                                                           // Из-за неправильного порядка в этой строке применяем метод reverse
      
      for (var i = 0; i < periods.length; i++) {                                                   // Для каждого периода
        var interval = data[periods[i]];                                                            
        var row = 6;                                                                               // Данные о профите/кол-ве идут с 6 строки
        calculatorPage.getRange( 4, col, 1, 1 ).setValue( labels[i] );                             // Из замыкания получим название периода
        calculatorPage.getRange( 5, col, 1, 2 ).setValues( [['Profit', 'Amount']] );  
        
        
        var intervalKeys = Object.keys( interval );                                                // Достать ключи, которые фактически равняются именам диспетчеров
        intervalKeys.reverse();                                                                    // сделать их в алфавитном порядке
        for (var j = 0; j < intervalKeys.length; j++) {
        
           calculatorPage.getRange( row, col, 1, 2 ).setValues( [[interval[intervalKeys[j]].profit, interval[intervalKeys[j]].amount]] ); // добавить данные о диспетчере
           row++;                                                                                  // Следующий диспетчер будет на следующей строке
        }
        
        col += 2;                                                                                  // Следующий период будет на 2 колонки правее
      }
    })();
    
    
    (function addStyling() {
      calculatorPage.getRange( 4, 1, 1, 9 ).
        setHorizontalAlignment('center');
  
      calculatorPage.getRange('B4:C4').mergeAcross();
      calculatorPage.getRange('D4:E4').mergeAcross();
      calculatorPage.getRange('F4:G4').mergeAcross();
      calculatorPage.getRange('H4:I4').mergeAcross();
      
      var lastRow = calculatorPage.getLastRow();
      
      calculatorPage.getRange('A4:I' + lastRow).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN);
      calculatorPage.getRange('A4:G5').setFontWeight('bold');
      calculatorPage.getRange( 6, 2, lastRow - 5, 1 ).setNumberFormat("[$$]#,##0.00");
      calculatorPage.getRange( 6, 4, lastRow - 5, 1 ).setNumberFormat("[$$]#,##0.00");
      calculatorPage.getRange( 6, 6, lastRow - 5, 1 ).setNumberFormat("[$$]#,##0.00");
      calculatorPage.getRange( 6, 8, lastRow - 5, 2 ).setNumberFormat("[$$]#,##0.00");
      
      calculatorPage.getRange('A:Q').setHorizontalAlignment('center');
    })();
  
  }
  
  
  
  