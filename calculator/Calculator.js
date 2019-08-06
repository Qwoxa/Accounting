function main() {

    var controlPanelSS = SpreadsheetApp.openById('1siWzmBYQpiLabUWA1ZIVs6UHtrhJCPumkmdR9yitPHM'); // С этой таблицы достаём имена диспетчеров
    var profitBoardSS = SpreadsheetApp.openById('1_f2hK3uf4tQoZ3XZfAL-SsEihrk9MwBTS5mw3fs2YD0'); // С этой таблицы достаём данные о грузах
    var calculatorSS = SpreadsheetApp.openById('1CYrmPmWvh7C-qiBRbFSrG6-RFm8WNGOgcyMeM5So-DA'); // Таблица калькулятора
    var calculatorS = calculatorSS.getSheetByName('Calculator Delta'); // На эту страницу нужно выводить информацию / брать из неё дату
    var payPeriodS = calculatorSS.getSheetByName('Pay Period'); // На этой странице находятся исключения в платёжных периодах
    
    // Список диспетчеров (те, которых нужно учитывать при рассчётах)
    var dispatchersList = getDispatchersList( controlPanelSS );
    
    // Здесь логика, как получить дату. isDateChanged показывает, менялась ли дата вручную.
    var chosenDate =  calculatorS.getRange('B3').getValue() || calculatorS.getRange('B2').getValue();
    var isDateChanged = typeof calculatorS.getRange('B2').getValue() === 'object';
    
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
   * param {Object} ss Таблица, из которой нужно вытянуть данные.
   * return {Object} Массив данных из таблицы.
   */
  function extractData(ss) {
    var sheets = ss.getSheets();
    
    // Проходит по каждой странице таблицы, получая из неё данные в виде массивов, формируя один массив
    var sheetsData = sheets.map(function(sheet) {
      // Если страница страница, то достаём только первую строку во избежание ошибок
      var lastRow = sheet.getLastRow() || 1;
      var lastColumn = sheet.getLastColumn() || 1;
      return sheet.getSheetValues( 1, 1, lastRow, lastColumn );
    });
    
  
    // Объединить все элементы массива в массив data
    var data = sheetsData.reduce( function(accum, sheetData) {
      return accum.concat( sheetData );
    }, [] );
    
    
    return data;
  }
  
  
  /**
   * Возвращает список диспетчеров 360.
   * @param {Object} controlPanel Таблица со списком диспетчеров
   * @return {Object} Список диспетчеров.
   */ 
  function getDispatchersList(controlPanel) {
    // получаем таблицу со списком
    var dispatchersTable = controlPanel.getSheetByName('Dispatchers');
  
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
  