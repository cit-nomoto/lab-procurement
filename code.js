/*
!! デプロイ時操作 !!

gas　エディタ→サービス→Admin SDK APIのインポート
→メンション時のUID取得用、一度特定したら関数を使わず直接IDを入れても良い

gas　トリガー→トリガーの編集→イベントの種類を選択→「変更時」に指定→保存
→activeCellの変更検出用

sendGoogleChatMessage関数の調達希望表リンクを適切なものに差し替え

searchUserId関数のメンション相手のアドレスを適切なものに差し替え

2024年 4月11日
11期生　野本
*/

function chatNotice() {
    //シート情報取得
    let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let activeSheet = activeSpreadsheet.getActiveSheet(); 
    if(activeSheet.getName() != "シート1"){
      return;
    }
    

    //行・列の取得
    let activeCell = activeSheet.getActiveCell();
    let column = activeCell.getColumn();

    //9列目以外の操作は弾く、エラーハンドリング
    if(column !== 9) {
        return;
    }

    // データ取得、カラム宣言
    var value = activeCell.getValue();
    var row = activeCell.getRow();
    var rowData  = activeSheet.getRange(row, 2, 1, 8).getValues()[0];

    var No = rowData[0];
    var Name = rowData[1];
    var Product = rowData[2];
    var use = rowData[3];
    var quantity = rowData[4];
    var Status = rowData[7];

    //メッセージ作成
    let message = "";
    switch (value) {
        case "申請":
            message = 
            `No. ${No}
            申請者： ${Name}
            品　名： ${Product}
            用　途： ${use}
            数　量： ${quantity}` ;
            break;
        case "承認(発注)":
        case "却下":
            message = 
            `No. ${No}
            ${Name}さん：${Status}しました。`;
            break;
        case "着荷":
            message = 
            `No. ${No}
            ${Name}さん：${Status}しました。
            検収後、領収書を返却し、
            調達希望表のステータスを検収済に変更してください。`;
            break;
        default:
            return;
    }
    //メッセージの送信
    sendGoogleChatMessage(message);
}
    
function sendGoogleChatMessage(message) {
    // webhookをスペースから取得
    var webhookUrl = "*****";  
    
    var cardMessage = {
        "text": `<users/${searchUserId()}>`, //メンションはカード内で宣言できないため、別途テキストとして送信
        "cards": [
            {
                "header": {
                    "title": "調達希望表を更新しました。"
                },
                "sections": [
                    {
                        "widgets": [
                            {
                                "textParagraph": {
                                    "text": message
                                }
                            },
                             {
                                "buttons": [
                                    {
                                        "textButton": {
                                            "text": "調達表希望表リンク",
                                            "onClick": {
                                                "openLink": {
                                                    //適切なリンクに差し替える
                                                    "url": "******"
                                                }
                                            }
                                        }
                                    }
                                ]
                            },
                        ]
                    }
                ]
            }
        ]
    };

    var options = {
        method: "POST",
        contentType: "application/json",
        payload: JSON.stringify(cardMessage)
    };

    UrlFetchApp.fetch(webhookUrl, options);
}

function searchUserId() {
    //UID出力
    const email = '****'; //メンション相手のアドレスに置き換え
    const user = AdminDirectory.Users.get(email, { viewType: 'domain_public' });
    const userId = (user.id); // Google Chat の USER_IDを拾う
    return userId;
}