function main(date, src, calc) {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); // Текущая таблица
    var srcPage = ss.getSheetByName( src ); // Страница, из которой нужно вытянуть данные по грузам
    var calcPage = ss.getSheetByName( calc ); // Страница, куда нужно выводить данные

    // Список диспетчеров (те, которых нужно учитывать при рассчётах)
    var dispatchersList = getDispatchersList( ss );
    
    // Дефолтная дата - сегодняшняя
    if ( typeof date !== 'object' ) {
        date = new Date();
        date.setHours( 0, 0, 0, 0 );
    }
    
    // Извлечь данные из профит борда
    var extractedData = extractData( profitBoardSS );
    
    
    // Фильтрует полученные данные: проверяет не Cancel ли груз, указан ли диспетчер и указана ли дата, когда груз забукали
    // Если условия удовлетворены - данные о дате, сумме профита и имени диспетчера
    // добавлятся в основной массив
    var data = extractedData.reduce(function(accum, record) {
      var dateBooked = record[0];
      var profit = record[14];
      var dispatcher = record[15];
      
      var dateBookedIsValid = typeof dateBooked === 'object';
      var isMarkedCanceled = String( record[8] ).toLowerCase() === 'cancel'; // Если cancel написали в Notes
      var isProfitCanceled = String( profit ).toLowerCase() === 'no'; // Если в профит написали No - это тоже cancel
      var isDispatcherDefined = dispatcher !== '';
      
      if ( dateBookedIsValid && !isMarkedCanceled && !isProfitCanceled && isDispatcherDefined ) {
        accum.push( [dateBooked, profit, dispatcher] );
      }
      
      return accum;
    }, []);
  
    
    // Module Dates.gs
    // Получаем даты начала и конца платёжного периода, недели и дня
    var excetionsPayPeriod = payPeriodS.getRange( 1, 2, payPeriodS.getLastRow() || 1, 2 ).getValues();
    var dates = CalculatorDate( chosenDate, excetionsPayPeriod );
    
    
    
    
    // Get statistics
    var statistics = {};
    
    for (var key in dates) {
      statistics[key] = getStatistics( data, dates[key], dispatchersList );
    }
  
    // JUST FOR TESTING PURPOSES
    
    RenderCalculator({
      page: calculatorSS.getSheetByName('test'),
      changedDate: null,
      data: statistics,
      dispatchers: dispatchersList
    });
    
  }
  
  
  /**
   * Вытягивает все данные из таблицы, формируя массив на выходе
   * param {Object} sheet Страница, из которой нужно вытянуть данные.
   * return {Object} Массив данных из таблицы.
   */
  function extractData(sheet) {

    var lastRow = sheet.getLastRow() || 1;
    var lastColumn = sheet.getLastColumn() || 1;
    // Проходит по каждой странице таблицы, получая из неё данные в виде массивов, формируя один массив
    var sheetsData = sheet.getSheetValues( 2, 4, lastRow, lastColumn );
  
    // Объединить все элементы массива в массив data
    var data = sheetsData.reduce( function(accum, sheetData) {
      return accum.concat( sheetData );
    }, [] );
    
    
    return data;
  }
  
  
  /**
   * Возвращает список диспетчеров 360.
   * @param {Object} ss Таблица со списком диспетчеров
   * @return {Object} Список диспетчеров.
   */ 
  function getDispatchersList(ss) {
    // получаем таблицу со списком
    var dispatchersTable = ss.getSheetByName('Dispatchers');
  
    // вытягиваем цвета ячеек и имена
    var dispatchers = dispatchersTable.getRange( 1, 1, dispatchersTable.getLastRow(), 1 ).getValues();
    var styles = dispatchersTable.getRange( 1, 2, dispatchersTable.getLastRow(), 1 ).getBackgrounds();
   
    // фильтруем имена
    var validNames = dispatchers.filter(function(dispatcher, index) {
      return dispatcher[0] != '' && dispatcher[0] != 'Name' && styles[index][0] != '#ffffff';
    });
    
    // двухмерный массив -> одномерный
    validNames.map(function(name){
      return name[0];
    });
    
    return validNames;
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
      }
    });
    
    return output;
  }
  