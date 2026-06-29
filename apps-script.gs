/**
 * 粟島 民宿 予約相談フォーム → スプレッドシート記録＆メール通知
 * 使い方:
 *  1. Googleスプレッドシートを新規作成
 *  2. 拡張機能 → Apps Script を開く
 *  3. このコードを全部貼り付けて保存
 *  4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *     - 実行ユーザー: 自分
 *     - アクセスできるユーザー: 全員
 *  5. 発行された「ウェブアプリのURL」を輝紀に渡す（フォームに設定する）
 */
var NOTIFY_EMAIL = "honbo@reterras.co.jp";   // 通知の宛先
var SHEET_NAME   = "予約相談";

function doPost(e){
  try{
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if(sh.getLastRow() === 0){
      sh.appendRow(["受付日時","代表者","電話","チェックイン","泊数","部屋数","大人","子供","食事","来る船","送迎","希望の宿","やりたいこと","その他希望","備考"]);
    }
    sh.appendRow([new Date(), d.name, d.tel, d.checkin, d.nights, d.rooms, d.adults, d.kids, d.meal, d.boat, d.pickup, d.yado, d.experiences, d.expOther, d.note]);

    var body =
      "新しい宿泊相談が届きました。\n\n" +
      "代表者　： " + d.name + "\n" +
      "電話　　： " + d.tel + "\n" +
      "チェックイン： " + d.checkin + "　／　" + d.nights + "泊\n" +
      "部屋数　： " + d.rooms + "室\n" +
      "人数　　： 大人" + d.adults + "名・子供" + d.kids + "名\n" +
      "食事　　： " + d.meal + "\n" +
      "来る船　： " + d.boat + "\n" +
      "送迎　　： " + d.pickup + "\n" +
      "希望の宿： " + d.yado + "\n" +
      "やりたいこと： " + ((d.experiences || "") + (d.expOther ? "（その他:" + d.expOther + "）" : "") || "特になし") + "\n" +
      "備考　　： " + (d.note || "なし") + "\n\n" +
      "（このメールは予約相談フォームから自動送信されています）";
    MailApp.sendEmail(NOTIFY_EMAIL, "【粟島 宿泊相談】" + d.name + "様", body);

    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

// 動作確認用（任意）: デプロイ後にブラウザでURLを開くとこの文字が出ればOK
function doGet(){
  return ContentService.createTextOutput("粟島 宿泊相談 受付スクリプト 稼働中");
}
