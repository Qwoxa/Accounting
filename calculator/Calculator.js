function testingCalc() {
  main( '', 'August', 'out' );
}

function main(date, src, calc) {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); // Текущая таблица
    var srcPage = ss.getSheetByName( src ); // Страница, из которой нужно вытянуть данные по грузам
    var calcPage = ss.getSheetByName( calc ); // Страница, куда нужно выводить данные

    // Список диспетчеров (те, которых нужно учитывать при рассчётах)
    var dispatchers = getDispatchers( ss );
    
    // Дефолтная дата - сегодняшняя
    if ( typeof date !== 'object' ) {
        date = new Date();
        date.setHours( 0, 0, 0, 0 );
    }
    

    // Вытягивает данные из заданной таблицы
    var extractedData = srcPage.getSheetValues( 2, 3, srcPage.getLastRow() || 1, 9 );
    
    
    // Фильтрует полученные данные (указан ли диспетчер, корректна ли сумма, есть ли дата)
    // Если условия удовлетворены - данные о дате, сумме профита и имени диспетчера
    var data = extractedData.reduce(function(accum, record) {
      var dateBooked = record[0];
      var profit = record[7];
      var dispatcher = record[8];
      var dateBookedIsValid = typeof dateBooked === 'object';
      var isDispatcherDefined = dispatcher !== '';

      if ( dateBookedIsValid && isDispatcherDefined && !isNaN(profit) ) {
        accum.push( [dateBooked, profit, dispatcher] );
      }
      
      return accum;
    }, []);
  
  
    
    // Module Dates.gs
    // Получаем даты начала и конца платёжного периода
    var exs = ss.getSheetByName('Pay Period');
    var excetionsPayPeriod = exs.getRange( 1, 2, exs.getLastRow() || 1, 2 ).getValues();
    var dates = CalculatorDate( date, excetionsPayPeriod );
    
    
    // Get statistics
    var statistics = {};
    

    var next = dates.start;
    while ( next.valueOf() !== dates.end.valueOf() ) {
      // next iteration = + one day
      nextIt = new Date( next.valueOf() );
      nextIt.setDate( nextIt.getDate() + 1 );

        statistics[next] = getStatistics( data, {
          start: next,
          end: nextIt
        }, dispatchers );
      

      next = nextIt;
    }
    
    Logger.log( statistics );
//    // JUST FOR TESTING PURPOSES
//    
//    RenderCalculator({
//      page: calcPage,
//      data: statistics,
//      dispatchers: dispatchers
//    });
    
  }
  
  
  
 /**
 * Возвращает объекты диспетчеров с их тимами.
 * @param {Object} ss Таблица со списком диспетчеров
 * @return {Array} Список диспетчеров (объект - имя, тим)
 */ 
function getDispatchers(ss) {
  // получаем таблицу со списком
  var dispatchersTable = ss.getSheetByName('Dispatchers');

  // вытягиваем цвета ячеек и имена
  var dispatchers = dispatchersTable.getRange( 1, 1, dispatchersTable.getLastRow() || 1, 2 ).getValues();
 
  // фильтруем имена
  var validNames = dispatchers.filter(function(dispatcher, index) {
    return dispatcher[0] != '' && dispatcher[0] != 'Name';
  });
  
  var out = [];
  // вытянуть данные в объекты
  validNames.map(function(disp){
    var obj = {};
    obj.name = disp[0];
    obj.team = disp[1];
    
    out.push( obj );
  });
  
  return out;
}
  
  
  /**
   * Получить статистику за определённый промежуток времени по указанным диспетчерам.
   * @param {Object} data Отфильтрованные данные из профит борда.
   * @param {Object} bounds Временные рамки, за которые нужно подсчитать статистику.
   * @param {Object} dispatchers Список диспетчеров, которых нужно учитывать.
   * @return {Object} Статистика.
   */
  function getStatistics( data, bounds, dispatchers ) {
   // Объект, куда будут записываться данные по профиту
   var output = {
      total: {
        amount: 0,
        profit: 0
      }
    };
    
    // Инициализируем объекты для каждого диспетчера в объект для статистики
    dispatchers.forEach(function(dispatcher) {
       output[dispatcher] = {};
       output[dispatcher].amount = 0;
       output[dispatcher].profit = 0;
       output[dispatcher.team] = {};
       output[dispatcher.team].amount = 0;
       output[dispatcher.team].profit = 0;
    });
    
    
    // Перебор записей из профит борда
    data.forEach(function(record) {
      var dateBooked = record[0];
      var profit = record[1];
      var dispatcher = record[2];
      
      // если груз забукал нужный диспетчер и временные рамки соответствуют - учитываем груз
      if ( dispatcher in output && dateBooked >= bounds.start && dateBooked < bounds.end ) {
        output[dispatcher].profit += Number( profit );
        output[dispatcher].amount++;
        output.total.profit += Number( profit );
        output.total.amount++;
        output[dispatcher.team].profit += Number( profit );
        output[dispatcher.team].amount++;
      }
    });
    
    return output;
  }
  