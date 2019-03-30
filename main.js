var botName   = '@KPab_bot'
var API_TOKEN = 'TELEGRAM_BOT_TOKEN'
var CONFIG_ID = 'GOOGLE_ID__CONFIG'
var scriptURL = 'https://script.google.com/macros/s/GOGGLE_SCRIPT_SHARE_URL/exec'

function doPost(e) {
  var error = false
  var message = ''
  
  // получаем сигнал от бота
  var update = JSON.parse(e.postData.contents)
  
  //проверяем тип полученного, нам нужен только тип "сообщение"
  if (update.hasOwnProperty('message')) {
    var msg = update.message
  
    //Разбор полей от кого
    var chatId = msg.chat.id
    var chatType = msg.chat.type
    if (chatType == 'private') {
      var chatDesc = msg.chat.first_name +' '+ msg.chat.last_name + '(@'+msg.chat.username+')'
    }else{
      var chatDesc = msg.chat.title
    }
    
    // проверяем, является ли сообщение командой к боту
    if (msg.hasOwnProperty('entities') && msg.entities[0].type == 'bot_command') {
      
      //Загрузка конфигурации
      var botConfigSheet = SpreadsheetApp.openById(CONFIG_ID).getSheetByName('CONFIG')
      var botConfig = botConfigSheet.getRange('A2:C')
      var botConfigVal = botConfig.getValues()
      var grafikToken = 'none'
      for (var row = 0; row < botConfigVal.length; row++) {
        if (chatId == botConfigVal[row][1]) {
          grafikToken = botConfigVal[row][2]
        }
      }
      //разбираем команду
      var command = msg.text.split(" ");
      
      //проверяем GOOGLE_ID
      if (((grafikToken == 'none' )||(grafikToken == 'deleted'))&&(command[0]!='/setup')&&(command[0]!='/setup'+botName)&&(command[0]!='/delete')&&(command[0]!='/delete'+botName)) {
        error = true
        message = message+'\nУстановите ID вашего графика в Google таблицах командой: <strong>/setup [GOOGLE_ID]</strong>'
      }
      
      //===TODAY===
      //==YESTERDAY==
      //==TOMORROW==
      out_today: if ( ( (command[0] == '/today')     || (command[0] == '/today'    +botName) ||
                        (command[0] == '/yesterday') || (command[0] == '/yesterday'+botName) ||
                        (command[0] == '/tomorrow')  || (command[0] == '/tomorrow' +botName)     
      ) && (!error) ) {
               
        var offset = 0;
        if ((command[0] == '/yesterday') || (command[0] == '/yesterday'+botName)) {offset=-1}
        if ((command[0] == '/tomorrow')  || (command[0] == '/tomorrow' +botName)) {offset=+1}
               
        // загружаем настройки отображения графика
        try {
          var grafikConfig =  SpreadsheetApp.openById(grafikToken).getSheetByName('KPab_bot').getRange('A:D').getValues()
        }catch(e) {
          error = true;
          message = message+'\nНеверный Google_ID или ошибка доступа';
          break out_today;
        }
        var isMainCaptionSet = false;
        
        // открываем график работы 
        var sheets =  SpreadsheetApp.openById(grafikToken).getSheets()
        
        // обходим все листы кроме конфига
        for (var s=0; s < sheets.length; s++) {
          var sheet = sheets[s];
          var sheetName = sheet.getName();
          
          if ( (sheetName != 'KPab_bot') && ((command.length==1)||(command[1]==sheetName)) ) {
            var isCaptionSet = false
           
            //получаем первую строку
            var firstRow = sheet.getRange("A1:1").getValues()[0]
            var grafik = sheet.getRange("2:367").getValues()
            
            var d=0;
            while (''+firstRow[0] != grafik[d][0] ) { d++ }
            d=d+offset;
            
            //в цикле проверяем кто работает и добавляем в сообщение
            for (var i = 1; i < firstRow.length; i++) {
              for (var j = 3; j < grafikConfig.length; j++) {
                if ((grafik[d][i] == grafikConfig[j][1])&&((grafikConfig[j][0]=='*')||(grafikConfig[j][0]==sheetName)) ){ 
                  
                  if (!isMainCaptionSet) {
                    message = message+'\n'+grafikConfig[1+offset][0]
                    isMainCaptionSet = true;
                  }
                  if (!isCaptionSet) {
                    message = message + grafikConfig[1+offset][2] + sheetName + grafikConfig[1+offset][3]
                    isCaptionSet = true;
                  }
                  message = message + grafikConfig[j][2]+firstRow[i]+grafikConfig[j][3] 
                }
              }
            }
            
            if (!isMainCaptionSet) {
              message = message+'\n'+grafikConfig[1+offset][1]
              isMainCaptionSet = true;
            }
          }
        }
        
        if (!isMainCaptionSet) {
          message = message+'\nНи чего не найдено! Проверьте настройки.'
          error = true;
        }
      
      }
      
      //===CHATID===
      if (command[0] == '/chatid') {
        message = ''+chatId
      }
      
      //===PING===
      if (command[0] == '/ping') {

        message = 'pong';
        if (command[1]) {message = 'pong '+command[1]}
      }
      
      //===TEST===
      if (command[0] == '/test') {

        message = e.postData.contents;
      }
      
      
      //===SETUP===
      out_setup: if ( ((command[0] == '/setup')||(command[0] == '/setup'+botName)) ) {
        
        if (command.length==2) {
          
          // установка в none и deleted эквивалентны команде удалить
          if ((command[1] == 'deleted')) {
            command[0] = '/delete'+botName
            break out_setup
          }
          // проверяем доступность таблицы
          try {
            var test =  SpreadsheetApp.openById(command[1]).getSheetByName('KPab_bot').getRange('A:D').getValues()
          }catch(e) {
            error = true;
            message = message+'\nНе удалось открыть лист KPab_bot по указанному GOOGLE_ID';
            break out_setup;
          }
          
          //Новые значения
          var newValues = [chatDesc, chatId, command[1]];
          
          //Поиск уже установленного токена
          var row = 0;
          while ( (chatId != botConfigVal[row][1])&&(botConfigVal[row][1] != '') ) { row++ }
          row = row+2;  
          botConfigSheet.getRange('A'+row+':C'+row).setValues([newValues])
          message = message +'\nНовый GOOGLE_ID установлен!'
        }else{
          error = true
          message = messge+'\nНеверное число параметров! Используйте команду:\n<strong>/setup [GOOGLE_ID]</strong>'
        }
      }
      
      //===DELETE===
       if ((command[0] == '/delete')||(command[0] == '/delete'+botName)) {
         var newValues = [chatDesc, chatId, 'deleted'];
         
         //Поиск уже установленного токена
         var row = 0;
         while ( (chatId != botConfigVal[row][1])&&(botConfigVal[row][1] != '') ) { row++ }
         row = row+2;  
         botConfigSheet.getRange('A'+row+':C'+row).setValues([newValues])
         message = message+'\nGOOGLE_ID удален!'
         
         //очистить cron
         command[0]='/cron'
         command[1]='clear'
       }
      
      
      //===CRON===
      if (((command[0] == '/cron')||(command[0] == '/cron' +botName))&&(!error)) {
       
        //Загрузка конфигурации
        var botCronSheet = SpreadsheetApp.openById(CONFIG_ID).getSheetByName('CRON')
        var botCron = botCronSheet.getRange('A2:G')
        var botCronVal = botCron.getValues()

        if (command[1]=='list') {
          message = message+'\n<b>Список задач:</b>\n<i>(Num) </i><b>"COMMAND"</b> <i>[minute hour day-of-month month day-of-week]</i>'
          for (var row = 0; row < botCronVal.length; row++) {
            if (chatId == botCronVal[row][0]) {
              message = message+'\n('+(row+1)+') <b>"'+ botCronVal[row][1]+'"</b> ['+botCronVal[row][2]+' '+botCronVal[row][3]+' '+botCronVal[row][4]+' '+botCronVal[row][5]+' '+botCronVal[row][6]+']'
            }
          }
        }
        
        if (command[1]=='del') {
          var row = Number(command[2])
          if ( (chatId == botCronVal[row-1][0]) ) {
            botCronSheet.getRange(row+1, 1, 1, 7).deleteCells(SpreadsheetApp.Dimension.ROWS)
            message = message+'\nЗадание cron удалено'
          }else{
            message = message+'\nОшибка удаления задания '
            error = true
          }
        }
        
        if (command[1]=='clear') {
          for (var row = botCronVal.length-1; row>=0; row--) {
            if (chatId == botCronVal[row][0]) {
              botCronSheet.getRange(row+2, 1, 1, 7).deleteCells(SpreadsheetApp.Dimension.ROWS)
            }
          }
          message=message+'\nВсе задания cron очищены';
        }
        
        if (command[1]=='add') {
          
          var cmd = msg.text.split('"')
          var param = cmd[2].split(' ')
          
          if (!param[1]) {param[1] = '*'}
          if (!param[2]) {param[2] = '*'}
          if (!param[3]) {param[3] = '*'}
          if (!param[4]) {param[4] = '*'}
          if (!param[5]) {param[5] = '*'}

          
          botCronSheet.getRange('A2:G2').insertCells(SpreadsheetApp.Dimension.ROWS).setValues([[chatId,cmd[1],param[1],param[2],param[3],param[4],param[5]]])
          message = message+'\nЗадание cron добавлено'
        }
      }
      
      
      //===HELP===
      //===START===
      if ((command[0] == '/help')||(command[0] == '/help'+botName)||(command[0] == '/start')||(command[0] == '/start'+botName)) {
        message = '\
Привет! Я бот "Кто работает", скоращенно "КРаб".\n\
Моя задача - отображать список сотрудников на основании графика работы.\n\
\n\
График работы - это простая Google таблица, в которой в левой верхней ячейке написана формула <b>=СЕГОДНЯ()</b>, в столбцах указываются фамилии сотрудников, а строках даты и график работы.\n\
Пример: https://docs.google.com/spreadsheets/d/1cbxaPQIMtQ1INnc8PH6CLaOL4h3xcXZzwDwvTJMut58/edit?usp=sharing\n\
\n\
Обрати внимание на служебный лист <b>KPab_bot</b>.\n\
Здесь в первых 3х строках указаны выводимые сообщения для команд yesterday/today/tomorrow(слева направо):\n\
- если кто-то работает в этот день.\n\
- если все отдыхают.\n\
- перед названием отдела(листа) <i>(хочешь стобы писалось с новой строки - введи здесь Ctrl+Enter).</i>\n\
- после названия отдела(листа).\n\
Далее идут шаблоны, которые сравниваются с графиком за текущий день и, если есть совпадение, выводится фамилия сотрудника.\n\
Первый столбец это название листа, к которому будет применяться шаблон. Знак <b>*</b> обозначает любой лист.\n\
Второй столбец это строка поиска. Срабатывает только точное совпадение, шаблоны и подстановочные символы не применимы.\n\
Третий столбец используется перед выводом фамилии сотрудника. Это хорошее место для вставки перевода строки или запятой перед фамилией.\n\
Четвертый столбец выводится после фамилии сотрудника. Здесь удобно вставить закрывающий тег.\n\
\n\
Чтобы подключиться к существующей таблице выполни команду <b>/setup [GOOGLE_ID]</b>.\n\
Например, для подключения демонстрационной таблицы из примера:\n\
<b>/setup 1cbxaPQIMtQ1INnc8PH6CLaOL4h3xcXZzwDwvTJMut58</b>\n\
\n\
После установки можешь использовать команду <b>/today</b> для получения списка всех работающих сотрудников на текущий день.\n\
Либо команду <b>/today Техподдержка</b> чтобы узнать только про один отдел.\n\
\n\
Аналогичные команды <b>/yesterday</b> и <b>/tomorrow</b> отвечают за вчерашний и завтрашний дни.\n\
\n\
Команда <b>/delete</b> удаляетт настройки бота и все задачи cron.\n\
\n\
\n\
Команда <b>/cron</b> позволит настроить автоматическое выполнение команд в назначенное время.\n\
<b>/cron add "/command" m h d m w</b> - добавит задание на выполнение комманды в указанные минуту час день месяц день_недели. Время указывается в часовом поясе UTC. Используйте подстановочный символ <b>*</b> или пропустите параметр.\n\
Пример: <b>/cron add "/today" 0 6</b> создаст задачу "/today" на 6:00 UTC (9:00 MSK) ежедневно.\n\
<b>/cron list</b> - выведет список уже существующих заданий.\n\
<b>/cron del [N]</b> - удалит задачу номер N. Порядковый номер можно узнать командой <b>/cron list</b>. Учтите, что после каждого удаления порядковые номера изменяются.\n\
<b>/cron clear</b> - удалит все задачи cron.\n\
\n\
(c)2019 @akokarev\n\
https://github.com/akokarev/KPab_bot\n\
2533296@gmail.com'
      }
      
      //=======
      
      //формируем ответ
      var payload = {
        'method': 'sendMessage',
        'chat_id': String(chatId),
        'text': message,
        'parse_mode': 'HTML'
      }     
      var data = {
        "method": "post",
        "payload": payload
      }
      
      // и отправляем его боту
      UrlFetchApp.fetch('https://api.telegram.org/bot' + API_TOKEN + '/', data);
      
    }
  }

}

function doCron(e) 
{
  //Загрузка конфигурации
  var botCronSheet = SpreadsheetApp.openById(CONFIG_ID).getSheetByName('CRON')
  var botCron = botCronSheet.getRange('A2:G')
  var botCronVal = botCron.getValues()
  
  for (var row = 0; row < botCronVal.length; row++) {
        if (
          ( (e.minute          == botCronVal[row][2])||('*'==botCronVal[row][2]) )&& //минута
          ( (e.hour            == botCronVal[row][3])||('*'==botCronVal[row][3]) )&& //час
          ( (e['day-of-month'] == botCronVal[row][4])||('*'==botCronVal[row][4]) )&& //день
          ( (e.month           == botCronVal[row][5])||('*'==botCronVal[row][5]) )&& //месяц
          ( (e['day-of-week']  == botCronVal[row][6])||('*'==botCronVal[row][6]) )   //день недели 
        ) {
          var v_postData = {contents:'{"message":{"chat":{"id":'+botCronVal[row][0]+',"title":"GSTriger","type":"group"},"text":"'+botCronVal[row][1]+'","entities":[{"type":"bot_command"}]}}'}
          var e={postData:v_postData}
          doPost(e)
        }
  }
}

function setWebhook() {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + API_TOKEN + "/setWebhook?url=" + scriptURL);
}

function TestPost() {
  var chatid = 'YOUR_CHAT_ID'
  var msg = '/ping'
  var v_postData = {contents:'{"message":{"chat":{"id":'+chatid+',"title":"TestPost","type":"group"},"text":"'+msg+'","entities":[{"type":"bot_command"}]}}'}
  var e={postData:v_postData}
  doPost(e)
}
