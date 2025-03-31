/**
 * ボタン押下した日付から前の日付にあるスプレッドシート情報をLINEで通知するプログラム
 * */ 

// LINEでグループにpush通知するプログラム
const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty("LINE_ACCESS_TOKEN");
const LINE_ENDPOINT = 'https://api.line.me/v2/bot/message/push';

function pushToLine(groupId: string, messages: any) {
    const today = Utilities.formatDate(new Date(), "JST", "yyyy年MM月dd日")

    const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_TOKEN}`
    };

    const postData = {
        'to': groupId,
        "messages": [
            {
                "type": "text",
                "text": `先生各位\n${today}までの連絡事項となります。\nご確認の上、修正報告をお願いいたします。\nよろしくお願いいたします。\n\n${messages}\n\n` //最後にGAS側でスプレッドシートリンクを挿入する
            },
        ]
    };

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: "post" as GoogleAppsScript.URL_Fetch.HttpMethod,
        headers: headers,
        payload: JSON.stringify(postData)
    };
    UrlFetchApp.fetch(LINE_ENDPOINT, options);
}

function onButtonClick() {
    const sheetName = "シート1";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(sheetName);
    const lastColumn: number = targetSheet?.getLastColumn()!;

    const groupId = PropertiesService.getScriptProperties().getProperty("TEST_GROUP_LINE")!;
    // const groupId = PropertiesService.getScriptProperties().getProperty("GROUP_LINE_ID");  // 本番用LINEグループのID
    const filteredData = getFilterdData();
    
    if (filteredData.length === 0) {
        return;
    }

    // データを整形してメッセージを作成
    const messages = filteredData.map((item, index) => {
        return `【${index+1}件目】\n` + 
        `店舗名: ${item.店舗名}\n` +
        `記入日: ${item.記入日}\n` +
        `施術業務日: ${item.施術業務日}\n` +
        `連絡内容: ${item.連絡内容}`;
    }).join("\n\n");

    // LINE送信
    pushToLine(groupId, messages);

    // 送信後、チェックボックスをTrueに更新
    filteredData.forEach(item => {
        targetSheet?.getRange(item.rowIndex, lastColumn).setValue(true);
    });
}
// LINEでグループにpush通知するプログラムここまで


// スプレッドシートからデータを取得
function getDatum(sheetName: string): { dateUnixTime: number, 記入日: string; 施術業務日: string; 店舗名: any; 連絡内容: any; 送信チェック: boolean; rowIndex: number}[] {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(sheetName);

    const lastRow: number | undefined = targetSheet?.getRange(2,2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    const lastColumn: number = targetSheet?.getLastColumn()!;

    if(!lastRow || lastRow < 4) return [];
    const allData = targetSheet?.getRange(4, 2, lastRow-3, 4).getValues();

    return allData?.map((row, index) => ({
        "dateUnixTime": Date.parse(row[0])/1000,
        "記入日": Utilities.formatDate(row[0],"JST", "yyyy/MM/dd"),
        "施術業務日": Utilities.formatDate(row[1],"JST", "yyyy/MM/dd"),
        "店舗名": row[2],
        "連絡内容": row[3],
        "送信チェック": targetSheet?.getRange(index + 4, lastColumn).getValue(),
        "rowIndex": index + 4  // 実際の行番号を保持
    })) ?? [];
}

function getFilterdData(){
    const sheetName = "シート1";
    // 今日のUnixタイムスタンプを取得
    const today = Math.floor(Date.now() / 1000);
    // 1週間前のUnixタイムスタンプを取得 (7日 * 24時間 * 60分 * 60秒)
    const oneWeekAgo = today - (7 * 24 * 60 * 60);
    
    // getDatumで取得したデータを一週間以内に絞り込む
    const data = getDatum(sheetName).filter(item => {
        return item.dateUnixTime >= oneWeekAgo && item.dateUnixTime <= today && item["送信チェック"] === false;
    });

    console.log(data);
    return data;
}


function doPost(e: any){
    const contents = JSON.parse(e.postData.contents);
    if(contents.events && contents.events.length > 0) {
        const event = contents.events[0];

        if(event.type === "join"){
            const groupId = event.source.groupId;
            console.log(groupId);
        }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "200" })).setMimeType(ContentService.MimeType.JSON);

}

/**
 * メッセージを送信する
 * @param {string} groupId グループID
 * @param {string} message 送信するメッセージ
 */
function sendReplyMessage(groupId: string, message: string) {
    const url = 'https://api.line.me/v2/bot/message/push';
    const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_TOKEN}`,
    };
    const payload = {
        to: groupId,
        messages: [
            {
            type: 'text',
            text: message,
            },
        ],
    };

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: 'post',
        headers: headers,
        payload: JSON.stringify(payload),
    };

    UrlFetchApp.fetch(url, options);
}