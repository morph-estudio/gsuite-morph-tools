function createTimeTriggerSpecifcDate() {
 ScriptApp.newTrigger("postLessons")
   .timeBased()
   .at(new Date('September 20, 2022 18:30'))
   .create();
}

// Morph Chat Scheduler Configuration
const CONFIG_SHEET = 'Mensajes'; // 🧰 Name of Morph Chat Scheduler 'configuration' sheet
const LESSON_SHEET = 'Mensajes'; // ✏️ Name of Morph Chat Scheduler 'lessons' sheet 
//const EMAIL_COLUMN = 5; // 📥 CONFIG_SHEET column which contains the 'Notify By Email When Complete' checkboxes 

/**
 * This function posts from the Chat Scheduler LESSON_SHEET 
 * to the Google Chat room(s) specified on the CONFIG_SHEET 
 * using a timed trigger manually configured by the Chat Scheduler Sheet owner
 * See https://support.google.com/chat/answer/7653861 for more on Google Chat rooms
 */
function postLessons() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const DATE_FORMAT = "M/d/yyyy k:mm:ss"; // 📆 Timestamp format for posted lessons on MAIN_SHEET, https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html

  let rows = sh.getDataRange().getValues();
  rows.forEach((row, index) => {
    if (index === 0) return; // Check if this row is the headers, if so we skip it

  try {

        let lessons = getLessons(ss);
        if (lessons && lessons.length > 0) {
          const lesson = lessons.shift();  // Get first unposted lesson
          Logger.log(`Roomnamebase = ${lesson.roomName}`)
          Logger.log(`RoomNOTIFY = ${lesson.notify}`)
          const config = getDatabaseShit(ss, lesson.roomName, lesson.notify);
          Logger.log(`Roomurl = ${config.roomurls}`)
          Logger.log(`configEmails = ${config.emails}`)
          let payload = getContent(lesson);  // Get lesson content in Google Chat message format
          let options = {
                'method' : 'post',
                'contentType': 'application/json; charset=UTF-8',
                'payload' : JSON.stringify(payload)
              };
          let date = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), DATE_FORMAT);
          UrlFetchApp.fetch(config.roomurls, options); // Post lesson to chat room(s)
          ss.getSheetByName(LESSON_SHEET).getRange(lesson.row,1,1,2).setValues([[true,date]]); // Update lesson in LESSON_SHEET


        if (lessons.length === 0) {
          stopTrigger('postLessons'); // Delete timed trigger if all lessons have been posted
        }
        if (lesson.notify) {
          sendEmail(ss,config.emails); // If enabled, send email notifications when lesson posting complete
        }


        }
        


  } catch(err) {
    Logger.log(err);
  }

  });
}

/**
 * This function gets the MCScheduler configuration from the CONFIG_SHEET
 * @param {Spreadsheet} ss - current Google Sheet
 * @return {Object} configuration object representing the MCScheduler CONFIG_SHEET
 */

function getConfigs(ss, index) {

  const rOOM_URL_COLUMN = fieldIndex(sh(), 'ChatRoomURL') + 1; // 🔗 CONFIG_SHEET column which contains the Chat Room URLs (column numbering starts at 1)
  const config = ss.getSheetByName(CONFIG_SHEET);
  Logger.log(`Room URL Column: ${rOOM_URL_COLUMN}`)
  const rooms = config.getRange(index + 1, rOOM_URL_COLUMN, config.getLastRow()-1).getValues()
                      .filter(row => row[0] !== '' && row[0].match(/^https:\/\/chat\.googleapis\.com/));
  return {
    roomurls: rooms.length > 0 ? Array.from(new Set(rooms.reduce((p,c) => p.concat(c)))) : null,
  }
}


function getConfig(ss, index) {

  const data = ss.getSheetByName(LESSON_SHEET);
  let config = data.getDataRange().getValues().map((row,index) => {
    return {row:index+1,roomName:row[10],sendEmail:row[11]};
  });
  Logger.log(`config = ${config.roomName}`)
  //config.shift(); // Shift off column titles row
  //config = config.filter(row => !row.posted);

  Logger.log(`Roomname = ${config.roomName}`)

  //let roomurl = getDatabaseShit(config.roomName)
  //Logger.log(roomurl)

  return {
    roomurls: roomurl,
  }
}



















function getDatabaseShit(ss, roomName, notify) {
  let searchValue = roomName; Logger.log(`SearchValue = ${searchValue}`)
  let headerName = 'ChatRoomName';
  const data = JSON.parse(UrlFetchApp.fetch('https://opensheet.elk.sh/1CeRqF8LP_3G-Gdo28TRkqD7VFEMYrjQSTIMJZW-bz6Y/ROOMS').getContentText()); Logger.log(data)
  //const rets = parsedDB.map(i => i[headerName])[option];
  //const ret = parsedDB[0]["ChatRoomURL"];

  const result = getFieldValue_({ data: data, searchKey: headerName, searchValue: searchValue, resultKey: 'ChatRoomURL'}); Logger.log(`RoomName = ${searchValue}`)

  const emails = [];
  if (notify) {
    emails.push(Session.getActiveUser().getEmail());
    ss.getEditors().forEach(editor => emails.push(editor.getEmail()));
  }

  return {
    roomurls: result,
    emails: Array.from(new Set(emails))
  }

  Logger.log(result)
}

function getFieldValue_({ data, searchKey, searchValue, resultKey } = {}) {
  const matches = data.filter(obj => obj[searchKey] === searchValue);
  if (!matches.length) {
    return null;
  }
  const firstMatch = matches[0];
  return firstMatch[resultKey];
}

function getConfigbase(ss) {
  const config = ss.getSheetByName(CONFIG_SHEET);
  const rooms = config.getRange(2, ROOM_URL_COLUMN, config.getLastRow()-1).getValues()
                      .filter(row => row[0] !== '' && row[0].match(/^https:\/\/chat\.googleapis\.com/));
  const notifications = config.getRange(1,EMAIL_COLUMN,3).getValues().reduce((p,c) => p.concat(c));
  const emails = [];
  // If email notifications enabled and notify Sheet editors enabled then collect Sheet editor email addresses
  if (notifications[0] && notifications[1]) {
    ss.getEditors().forEach(editor => emails.push(editor.getEmail()));
  } 
  // If email notifications enabled and notify Sheet viewers enabled then collect Sheet viewer email addresses
  if (notifications[0] && notifications[2]) {
    ss.getViewers().forEach(viewer => emails.push(viewer.getEmail()));
  }  
  // Flatten Chat rooms array and remove duplicate entries in rooms and emails
  return {
    roomurls: rooms.length > 0 ? Array.from(new Set(rooms.reduce((p,c) => p.concat(c)))) : null,
    emails: emails.length > 0 ? Array.from(new Set(emails)) : null
  }
}

/**
 * This function creates a Google Chat room post for a given lesson from the TSChatWise LESSON_SHEET 
 * @param {Object} lesson - a lesson from the TSChatWise LESSON_SHEET 
 * @return {Object} lesson content to be posted to a Google Chat room(s)
 */
function getContent(lesson) {
  const LESSON_BUTTON_TEXT = lesson.buttontext || 'HAZ CLIC PARA SABER MÁS'; // 🔳 Lesson chat message button text
  const widgets = [];
  const image = {"image":{"imageUrl":lesson.image}};
  const buttons = {"buttons":[{"textButton": {"text":LESSON_BUTTON_TEXT,
                                "onClick":{"openLink":{"url":lesson.link}}}}]};
  let text;
  // If lesson is card based, create a card lesson message
  // else create a text lesson message
  if (lesson.type) {
    if (lesson.name != '') {
      text = {"textParagraph":{"text": Utilities.formatString('<b>%s</b><br><br>%s<br>', lesson.name, lesson.description)}};
    } else {
      text = {"textParagraph":{"text": lesson.description}};
    }
    
    widgets.push(text);
    if (lesson.image && lesson.image !== '') {
      widgets.push(image);
    }
    if (lesson.link && lesson.link !== '') {
      widgets.push(buttons);
    }
    return {"cards":[{"sections":[{"widgets": widgets}]}]};
  } else {
    text = "*" + lesson.name + "*\n\n" + lesson.description;
    if (lesson.link) {
      text += " \n" + lesson.link;
    }
    return {"text":text};
  }
}

/**
 * This function retrieves the lesson information from the MCScheduler LESSON_SHEET 
 * @param {Spreadsheet} ss - current Google Sheet
 * @return {<Object>} array of lesson objects from the MCScheduler LESSON_SHEET  (with column titles row removed)
 */
function getLessons(ss) {
  const data = ss.getSheetByName(LESSON_SHEET);
  const lessons = data.getDataRange().getValues().map((row,index) => {
    return {row:index+1,posted:row[0],type:row[2],name:row[3],
            description:row[4],link:row[5],image:row[6],buttontext:row[7],roomName:row[10],notify:row[11]};
  });
  lessons.shift(); // Shift off column titles row
  return lessons.filter(row => !row.posted);
}

/**
 * This function sends email notifications when MCSchedulere has completed posting lessons on LESSON_SHEET 
 * @param {Spreadsheet} ss - current Google Sheet
 * @param {<string>} emails - email addresses
 */
function sendEmail(ss,emails) {
  let email = `<a href="${ss.getUrl()}">${ss.getName()}</a> has completed!`,
      subject = `${ss.getName()} has completed!`;

      let mailOptions = {
        htmlBody: email,
      };

      GmailApp.createDraft(emails.join(','), subject, '', mailOptions).send();


}

/**
 * This function deletes all time based triggers when MCScheduler has completed posting lessons on LESSON_SHEET 
 * @param {string} triggerFunction - MCScheduler Apps Script function to run for each trigger execution
 */
function stopTrigger(triggerFunction='postLessons') {
  ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getEventType() === ScriptApp.EventType.CLOCK && trigger.getHandlerFunction() === triggerFunction)
      .forEach(trigger => ScriptApp.deleteTrigger(trigger));
} 





/**
 * This function tests the posting to all configured chat rooms
 */
function testBot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let payload = {"text":"*Testing Morph Chat Scheduler*"};  // Test message
    let config = getConfig(ss); // Get Morph Chat Scheduler configuration from CONFIG_SHEET
    let options = {
      'method' : 'post',
      'contentType': 'application/json; charset=UTF-8',
      'payload' : JSON.stringify(payload)
    };
    if (config.roomurls) {
        // Post test to chat room(s)
        config.roomurls.forEach(room => UrlFetchApp.fetch(room, options));  
    } else {
      throw new Error('Morph Chat Scheduler: No se ha configurado ninguna sala de chat.');
    }
  } catch(err) {
    console.log(`MCScheduler: ocurrió el siguiente error - "${err.message}"`);    
  }
}