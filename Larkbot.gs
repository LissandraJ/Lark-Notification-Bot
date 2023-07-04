let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dataSet");

/**
 * Problem with googlesheet : sometimes the function getLastRow returns a blank row,
 * so it's necessary to get the last populated row.
 */
let plastRow = getLastPopulatedRow(sheet);
let p_task = sheet.getRange(plastRow, 1).getValue(); //第1列，col A
let p_due = Utilities.formatDate(sheet.getRange(plastRow, 2).getValue(), "GMT+8", "yyyy-MM-dd HH:mm"); //第2列，col B，北京时间
let p_platform = sheet.getRange(plastRow, 9).getValue(); //第9列，col I
let p_link = sheet.getRange(plastRow, 5).getValue(); //第5列，col E
let p_keys = sheet.getRange(plastRow, 3).getValue(); //第3列，col C
let p_words = sheet.getRange(plastRow, 4).getValue(); //第4列，col D
let p_assignee = sheet.getRange(plastRow, 11).getValue();

let translators_id = { //Translators LarkID
  'Translator1': 'ou_UserLarkID',
  'Translator2': 'ou_UserLarkID',
  'Translator3': 'ou_UserLarkID',
  'Translator4': 'ou_UserLarkID',
  'Translator5': 'ou_UserLarkID',
  // etc
}

let translators_name = Object.keys(translators_id) //内部译员的名字
let tracker = 'https://docs.google.com/spreadsheets/d/SheetID'  //GoogleSpreadsheet that tracks the progress of translation tasks
let invoice = 'https://j9tq1bsdoo.feishu.cn/sheets/SheetID' //Invoice in lark to fill out

function getLastPopulatedRow(sheet) {
  let arrVals = sheet.getDataRange().getValues();

  for (let i = arrVals.length - 1; i > 0; i--) {
    if (arrVals[i].join('')) {
      return i + 1
    }
  }

  return 0;
}

/**
 * Formats p_assignee string to Object : {languageID: translatorName}.
 *  p_assignee model :
 * "ar asd, zh_TW dfegy, en Karen, fr 惠欣, de CL, zh_TW 帅帅, ur mobi,te mobi, it CL, ms Sa Sing, pt_BR Clara Alice, ru CL, es Shuda, th CL, tr Cus, vi 阮"
 *  CL =  third party translation company for europeen languages, no access to corporation chat, to be informed in their chat
 *  mobi = third party translation company for indian languages, no access to corporation chat, to be informed in their chat
 *  zh_TW & zh_HK =  Taiwan & HongKong Team, no access to corporation chat, to be informed in their chat
 * @returns Object
 */
function languages_translators() {
  let formattedAssignees = {};
  let assigneesArr = p_assignee.split(', ');
  // Now: ['ar asd', 'ma Sa Sing', ...]

  assigneesArr.forEach((assigneeStr) => {
    let splitAssignee = assigneeStr.trim().split(' ');
    // Now: ['ms', 'Sa', 'Sing']

    if (splitAssignee.length <= 1) {
      console.error('get_translators: Error while formatting assignees', assigneesArr, assigneeStr);
    }

    let langID = splitAssignee.shift();
    // Now: ['Sa', 'Sing']

    formattedAssignees[langID] = splitAssignee.join(' ');
  });

  return formattedAssignees;
}



/**
 * Sends a message on Lark to inform task-related translators
 * @returns Object
 */

function whom_to_at() {
  let language_translator = languages_translators()
  languages = Object.keys(language_translator)
  let at_content = [] //corporation chat
  let CL_content = [] //CL third-party chat
  let mobi_content = [] // mobi third-party chat
  let HKTW_content = [] // HK,TW

  l
  for (let i = 0; i < languages.length; i++) {
    lang = languages[i]
    translator = obj[lang]
    if (translators_name.indexOf(translator) != -1) { //corporation chat
      at_content.push([
        {
          "tag": "text",
          "text": lang
        },
        {
          "tag": "at",
          "user_id": translators_id[translator]
        }
      ])
    }


    else if (translator == 'CL') { //CL third-party chat

      CL_content.push(
        {
          "tag": "text",
          "text": lang + ", "
        }
      )
    }


    else if (translator == 'mobiActivation') { //mobi third-party chat

      mobi_content.push(
        {
          "tag": "text",
          "text": lang + ", "
        }
      )
    }

    else if (lang == 'zh_TW' || lang == 'zh_HK') { //HK & TW chat

      HKTW_content.push(
        [{
          "tag": "text",
          "text": lang + ": "
        },
        {
          "tag": "text",
          "text": translator
        }
        ]
      )
    }

  }


  let out_content = [CL_content,// For CL to upload invoice
    [
      {
        'tag': 'text',
        'text': 'Invoice : '

      },
      {
        "tag": "a",
        "text": 'Upload invoice',
        "href": invoice
      }]
  ]

  let in_and_out_content = {
    'in_content': at_content, // corporqtion chat
    'CL_content': CL_content, // CL chat
    'out_content': out_content, //CL chat + invoice
    'mobi_content': [mobi_content], //mobi chat
    'HKTW_content': HKTW_content //HKTW Team  chat
  }
  Logger.log(CL_content)
  return in_and_out_content
}




function larkBot() {
  let in_and_out = Who_to_At()

  let in_url = "https://open.feishu.cn/open-apis/bot/v2/hook/larkbotID" //corporation message bot url
  let out_url = 'https://open.feishu.cn/open-apis/bot/v2/hook/larkbotID' //CL message bot url
  let mobi_url = 'https://open.feishu.cn/open-apis/bot/v2/hook/larkbotID' //mobi message bot url
  let HKTW_url = 'https://open.feishu.cn/open-apis/bot/v2/hook/larkbotID' //HKTW Team message bot url
  let in_content = in_and_out.in_content
  let out_content = in_and_out.out_content
  let CL_content = in_and_out.CL_content
  let mobi_content = in_and_out.mobi_content
  let HKTW_content = in_and_out.HKTW_content
  let headers = {
    "content-type": "application/json"
  }
  let text_content = [//任务详情
    [
      {
        "tag": "text",
        "text": "Due date: " + p_due
      }
    ],
    [
      {
        "tag": "text",
        "text": "Keys & Words : " + p_keys + " & " + p_words
      }
    ],
    [
      {
        "tag": "text",
        "text": "Platform : "
      },
      {
        "tag": "a",
        "text": p_platform,
        "href": p_link
      },
      {
        "tag": "text",
        "text": "  Tracker : "
      },
      {
        "tag": "a",
        "text": 'Update',
        "href": tracker
      }
    ],
    [
      {
        "tag": "text",
        "text": "Translator(s): "
      }
    ]
  ]

  let post_content = {
    "msg_type": "post",
    "content": {
      "post": {
        "zh_cn": {
          "title": p_task,
          "content": ''
        }
      }
    }
  }

  var options = {
    "headers": headers,
    "method": 'POST',
    "payload": JSON.stringify(post_content)
  }


  //tasks for corporation translators
  if (in_content.length != 0) {
    post_content['content']['post']['zh_cn']['content'] = text_content.concat(in_content) //内部
    var in_feishu = JSON.parse(UrlFetchApp.fetch(in_url,
      {
        "headers": headers,
        "method": 'POST',
        "payload": JSON.stringify(post_content)
      })
    )
  }

  //tasks for CL
  if (CL_content.length != 0) {
    post_content['content']['post']['zh_cn']['content'] = text_content.concat(out_content) //cl
    var out_feishu = JSON.parse(UrlFetchApp.fetch(out_url,
      {
        "headers": headers,
        "method": 'POST',
        "payload": JSON.stringify(post_content)
      })
    )
  }
  //tasks for mobi
  if (mobi_content[0].length != 0) {
    post_content['content']['post']['zh_cn']['content'] = text_content.concat(mobi_content) //mobi
    var mobi_feishu = JSON.parse(UrlFetchApp.fetch(mobi_url,
      {
        "headers": headers,
        "method": 'POST',
        "payload": JSON.stringify(post_content)
      })
    )
  }

  //tasks for HKTW team
  if (HKTW_content.length != 0) { 
    post_content['content']['post']['zh_cn']['content'] = text_content.concat(HKTW_content) //internal
    var internal_feishu = JSON.parse(UrlFetchApp.fetch(HKTW_url,
      {
        "headers": headers,
        "method": 'POST',
        "payload": JSON.stringify(post_content)
      })
    )

  }


  Logger.log(in_and_out)


}
