/**
 * ボタン押下した日付から前の日付にあるスプレッドシート情報をLINEで通知するプログラム
 * */ 

// LINEでグループにpush通知するプログラム
const LINE_TOKEN = 'YOUR_CHANNEL_ACCESS_TOKEN';
const LINE_ENDPOINT = 'https://api.line.me/v2/bot/message/push';

function pushToLine(groupId: string, messages: string[]) {
    const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${LINE_TOKEN}`
    };

    const postData = {
        'to': groupId,
        'messages': messages.map(text => ({
            'type': 'text',
            'text': text
        }))
    };

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: "post" as GoogleAppsScript.URL_Fetch.HttpMethod,
        headers: headers,
        payload: JSON.stringify(postData)
    };
    UrlFetchApp.fetch(LINE_ENDPOINT, options);
}

function onButtonClick() {
    const groupId = 'YOUR_GROUP_ID';  // LINEグループのID
    const filteredData = getFilterdData();
    
    if (filteredData.length === 0) {
        return;
    }

    // データを整形してメッセージを作成
    const messages = filteredData.map(item => 
        `店舗名: ${item.店舗名}\n` +
        `記入日: ${item.記入日}\n` +
        `施術業務日: ${item.施術業務日}\n` +
        `連絡内容: ${item.連絡内容}`
    );

    pushToLine(groupId, messages);
}
// LINEでグループにpush通知するプログラムここまで


// スプレッドシートからデータを取得
function getDatum(sheetName: string): { dateUnixTime: number, 記入日: string; 施術業務日: string; 店舗名: any; 連絡内容: any; }[] {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(sheetName);

    const lastRow: number | undefined = targetSheet?.getRange(2,2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    // console.log(lastRow);

    if(!lastRow || lastRow < 4) return [];
    const allData = targetSheet?.getRange(4, 2, lastRow-3, 4).getValues();

    return allData?.map((row) => ({
        "dateUnixTime": Date.parse(row[0])/1000,
        "記入日": Utilities.formatDate(row[0],"JST", "yyyy/MM/dd"),
        "施術業務日": Utilities.formatDate(row[1],"JST", "yyyy/MM/dd"),
        "店舗名": row[2],
        "連絡内容": row[3],
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
        return item.dateUnixTime >= oneWeekAgo && item.dateUnixTime <= today;
    });

    console.log(data);
    return data;
}


