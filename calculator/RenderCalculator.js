function RenderCalculator(options) {


  // Деструктуризация параметров
  var calculatorPage = options.page;
  var dispatchers = options.dispatchers;
  var data = options.data;
  
  var teams = dispatchers.reduce( function(acc, disp) {
  Logger.log( acc );
    if ( !acc.hasOwnPropertyName( disp.team ) ) {
      acc[disp.team] = [];
    }
    
    acc[disp.team].push( acc.name );
  }, {} );
  
  Logger.log( teams );
  
  /**
   * Размещает объект data на странице calculatorPage
   */
  (function showStatistics() {

    
  })();
  

}


