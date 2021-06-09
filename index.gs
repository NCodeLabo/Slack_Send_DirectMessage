/*
スプレッドシートに記載されているメールアドレスに対して、DMを送信する
 Send DM to email address listed in the Google Spreadsheet.
送信時の文章はGoogleドキュメントから取得する
 Text to be sent is taken from Google Docs.
*/

const TOKEN = "ここにあなたのUserOAuthTokenを貼り付けてください。"

function main() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadSheet.getActiveSheet();
    const sendMailList = sheet.getDataRange().getValues(); //選択しているシートのセルのデータをリストで全取得
    let successCount = 0, missCount = 0;

    /*
    main処理
     main process
    */
    for (i in sendMailList) {
        if (i > 2) {
            let userIds = "", usersName = "";
            let userId, name;
            let mail = sendMailList[i][0];

            /*
            1:1のDMと複数でのGroupChatで処理を分ける
             Split the process between DirectMessage and GroupChat
            */

            /*
            【True】
            GroupChatの場合
            for GroupChat
            【Else】
            1:1DMの場合
            for Direct Message
            */
            if (sendMailList[i][1]) {
                /*
                ユーザーID[id,id,...]
                ユーザー名[@hogehoge さん @hogehogeさん ....]の形に変更する
                change the userID [id,id,...]
                change the userName [@hogehogeさん @hogehogeさん ....]
                */
                for (n in sendMailList[0]) {
                    if (sendMailList[i][n]) {
                        userId = lookupByEmail(sendMailList[i][n]);
                        userIds = userIds + userId + ",";
                        name = getUserName(userId);
                        usersName = usersName + name + "さん ";
                    }
                }
                userId = conversationsOpen(userIds.slice(0, -1));
            } else {
                userId = lookupByEmail(mail);
                usersName = getUserName(userId) + " さん";
            }

            /*
            送信処理
            send process
            */
            if (userId) {
                const doc = DocumentApp.openByUrl(sendMailList[1][0]);
                let text = doc.getBody().getText();
                text = text.replace("${name}", usersName);
                console.log(usersName + "にDM送信処理を実行します！\n\idは 【" + userId + "】 です");
                console.log(text);
                try {
                    const send = postDM(userId, text);
                    successCount = successCount + 1;
                    console.log("DM送信成功しました。")
                    sendMailList[i][8] = "送信済";//行の一番右のセルに送信済みをつける
                } catch (e) {
                    missCount = missCount + 1;
                    const errorMsg = Browser.msgBox(usersName + " へのDM送信失敗しました。エラー文を確認して対応をお願いします。\n【エラー文】" + e);
                    console.log(usersName + " へのDM送信失敗しました。\n【エラー文】" + e);
                    sendMailList[i][8] = "送信失敗";
                }
            }
        } else {
            sendMailList[i][8] = "";
        }
    }
    sheet.getRange(1, 1, sendMailList.length, 9).setValues(sendMailList);
    const MsgBox = Browser.msgBox("送信処理が完了しました。\n\成功" + successCount + "件\n\失敗" + missCount + "件です");
    console.log("successCount = " + successCount + "\n\missCount = " + missCount);
}

/*
DM送信
 post message
@type none
*/
function postDM(userId, text) {
    const options = {
        "method": "post",
        "contentType": "application/x-www-form-urlencoded",
        "payload": {
            "token": TOKEN,
            "channel": userId,
            "text": text
        }
    };
    const url = 'https://slack.com/api/chat.postMessage';
    UrlFetchApp.fetch(url, options);
    return;
}

/*
Slackのidからメンション用のユーザー名を返す
 get user name by SlackID.
ユーザー名の先頭@をつけてメンションにする
 add "@" front of userName for mention
メンションは<@userName>の形式
 Mentions should be entered in the format <@userName>
@type String
*/
function getUserName(id) {
    let name = "";
    const options = {
        "method": "GET",
        "contentType": "application/x-www-form-urlencoded",
        "payload": {
            "token": TOKEN,
            "user": id
        }
    };
    const url = 'https://slack.com/api/users.info';
    const response = UrlFetchApp.fetch(url, options);
    const res = JSON.parse(response);
    /*
    正しくjsonを取得できない場合のtry-catch
     try-catch in case you can't get the json correctly
    */
    try {
        name = "<@" + res.user.name + ">";
    } catch (e) {
        console.log("getUserNameでエラーが発生しました\n\【id】" + id + "\n【name】" + name + "\n【エラー文】" + e);
        const errorMsg = Browser.msgBox("getUserNameでエラーが発生しました\n\【id】" + id + "\n【name】" + name + "\n【エラー文】" + e);
    }
    return name;
}

/*
メールアドレスを参照してSlackのユーザーIDを返す
 get SlackUserID by email address.
@type String
*/
function lookupByEmail(email) {
    let id = "";
    const options = {
        "method": "GET",
        "contentType": "application/x-www-form-urlencoded",
        "payload": {
            "token": TOKEN,
            "email": email
        }
    };
    const url = 'https://slack.com/api/users.lookupByEmail';
    const response = UrlFetchApp.fetch(url, options);
    const res = JSON.parse(response);
    /*
    正しくjsonを取得できない場合のtry-catch
     try-catch in case you can't get the json correctly
    */
    try {
        id = res.user.id
    } catch (e) {
        console.log("lookupByEmailでエラーが発生しました\n\【id】" + id + "\n【email】" + email + "\n【エラー文】" + e);
        const errorMsg = Browser.msgBox("lookupByEmailでエラーが発生しました\n\【id】" + id + "\n【email】" + email + "\n【エラー文】" + e);
    }
    return id;
}

/*
グループDMの参加者を参照してグループDMのidを返す get group chat id by member ids. 
@type String
*/
function conversationsOpen(users) {
    let id = "";
    const options = {
        "method": "GET",
        "contentType": "application/x-www-form-urlencoded",
        "payload": {
            "token": TOKEN,
            "users": users
        }
    };
    const url = 'https://slack.com/api/conversations.open';
    const response = UrlFetchApp.fetch(url, options);
    const res = JSON.parse(response);
    /*
    正しくjsonを取得できない場合のtry-catch
     try-catch in case you can't get the json correctly
    */
    try {
        id = res.channel.id;
    } catch (e) {
        console.log("conversationsOpenでエラーが発生しました\n\【id】" + id + "\n【users】" + users + "\n【エラー文】" + e);
        const errorMsg = Browser.msgBox("conversationsOpenでエラーが発生しました\n\【id】" + id + "\n【users】" + users + "\n【エラー文】" + e);
    }
    return id;
}
