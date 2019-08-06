/**
 * Функция возвращает объект с датами начала и конца платёжного периода, недели и дня.
 * @params {Object} date Дата, для которой нужно найти начало и конце платёжного периода и недели.
 * @params {Object} exceptions Массив, содержащий все исключения по Pay Period (когда изменяются рамки платёжного периода)
 */
function CalculatorDate(date, exceptions) {
    // Сделаем копию date, чтобы не было сайд эффектов, обнулим часы и минуты
    var date = new Date( date.valueOf() );
    date.setHours(0, 0, 0, 0);
    
    var payPeriod = {};
    var week = {};
    var day = {};
    
    
    /**
     * Инициализировать payPeriod
     */
    (function setPayPeriod() {
      // Если параметр exceptions был указан
      if ( typeof exceptions !== 'undefined' ) {
        // Обрабатываем исключения в обратном порядке
        exceptions.reverse();
        
        
        // Если указанная дата date является исключением -
        // берём укзаанный диапазан как начало и конец Pay Period.
        exceptions.forEach(function(exception) {
          var from = new Date( exception[0] );
          var to = new Date( exception[1] );
          
          if ( date >= from && date < to ) {
            payPeriod.start = from;
            payPeriod.end = to;
          }
        });
      }
      
      // Если было исключение - выходим из функции.
      if ( Object.keys( payPeriod ).length ) return;
      
      
      // Если не было исключений, то выбираем даты Pay Period по стандартному алгоритму (getPayPeriod)
      // Для конца Pay Period как параметр функции нужно передать дату + 1 месяц
      var nextMonth = new Date( date.valueOf() );
      nextMonth.setMonth( nextMonth.getMonth() + 1 );
      
      payPeriod.start = getPayPeriod( date );
      payPeriod.end = getPayPeriod( nextMonth );
    })();
   
   
    /**
     * Получает последний день Pay Period
     * param {Object} date Дата, для которой нужно найти конец Pay Period.
     * return {Object}  Последний день Pay Period.
     */
    function getPayPeriod(date) {
      // Сделаем копию date, чтобы не было сайд эффектов
      var date = new Date(date.valueOf());
      var year = date.getFullYear();
      var month; 
      
      if ( payPeriodStartsThisMonth(date) ) {
        month = date.getMonth();
      } else {
        month = date.getMonth() - 1;
        
        // Если Pay Period начался в предыдущем году
        if (month == -1) {
          month = 11;
          year -= 1;
        }
      }
      
      date.setFullYear(year, month, 1);
      
      // Узнать последнюю среду, после - последний понедельник
      var lastMonday;
      while (true) {
        if (date.getMonth() != month) break;
        if (date.getDay() == 3) {
          lastMonday = date.getDate() - 2;
        }
        date.setDate(date.getDate() + 1);
      }
    
      date.setFullYear(year, month, lastMonday);
      return date;
    }
    
    
    /**
     * Определяет, в этом ли месяце начался Pay Period для заданной даты
     * @param {Object} date Дата, для которой нужно узнать начинается ли этот Pay Period в этом месяце.
     * @return {Boolean} Если для заданной даты Pay Period начинается в этом месяце - возвращает true; иначе - false
     */
    function payPeriodStartsThisMonth(date) {
      // Сделаем копию date, чтобы не было сайд эффектов
      var date = new Date(date.valueOf());
      
      var dayCount = date.getDay();
      var month = date.getMonth();
      
      // Получить дату понедельника
      if (dayCount == 0) {
        date.setDate(date.getDate() - 6);
      } else if (dayCount != 0) {
        date.setDate(date.getDate() - dayCount + 1);
      }
    
      if (month != date.getMonth())  return false;
      var count = 0;
      var month  = date.getMonth();
      
      while (true) {
        if (date.getMonth() != month) break;
        if (date.getDay() == 3) ++count;
        
        date.setDate(date.getDate() + 1);
        if (count == 2) return false;
      }
      
      return true;
    }
    
  
    /**
     * Инициализировать week
     */
    (function setWeek() {
      // Для конца nextWeek нужно создать отдельный объект
      var nextWeek = new Date( date.valueOf() );
      nextWeek.setDate( nextWeek.getDate() + 7 );
      
      
      week.start = getWeek( date );
      week.end = getWeek( nextWeek );
    })();
    
    
    /**
     * Получает понедельник для заданной даты
     * @param {Object} date День, для которого нужно найти предшествующий понедельник
     * @return {Object} Дата понедельника.
     */
    function getWeek(date) {
      // Сделаем копию date, чтобы не было сайд эффектов
      var date = new Date(date.valueOf());
      
      // Получить день недели
      var dayCount = date.getDay();
      var thisWeek = date;
      
      // Модифицировать дату
      if (dayCount == 0) {
        thisWeek.setDate(date.getDate() - 6);
      } else if (dayCount != 0) {
        thisWeek.setDate(date.getDate() - dayCount + 1);
      }
      
      return thisWeek;
    }
    
    
    /**
     * Инициализировать day
     */
    (function setDay() {
      // Для конца Pay Period как параметр функции нужно передать дату + 1 месяц
      var tomorrow = new Date( date.valueOf() + 1000 * 3600 * 24 );
      
      day.start = date;
      day.end = tomorrow;
    })();
    
    
     var Exports = {
       payPeriod: payPeriod,
       week: week,
       day: day
     };
     
     
     return Exports;
  }