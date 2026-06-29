/**
 * 粟島 民宿 予約相談フォーム → スプレッドシート記録 ＆ メール通知 ＆ リマインド
 *
 * ◆セットアップ
 *  1. Googleスプレッドシートを新規作成
 *  2. 拡張機能 → Apps Script を開き、このコードを全部貼り付けて保存
 *  3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *     - 実行ユーザー: 自分 / アクセスできるユーザー: 全員 → デプロイ
 *  4. 発行された「ウェブアプリのURL」を輝紀に渡す（フォームに設定）
 *
 * ◆リマインドメール（1週間前・2日前）を有効にする
 *  5. 左メニュー「トリガー（時計アイコン）」→ トリガーを追加
 *     - 実行する関数: sendReminders
 *     - イベントのソース: 時間主導型 → 日付ベースのタイマー → 午前8〜9時 など
 *     これで毎日自動チェックし、行く日の7日前・2日前に申込者へメールが飛びます。
 */
var NOTIFY_EMAIL = "honbo@reterras.co.jp";                  // 幹事（通知の宛先）
var SHEET_NAME   = "予約相談";
var FORM_URL     = "https://honbo-teruki.github.io/awashima-ferry/yado.html";
var FERRY_TIMETABLE_URL = "https://awashimakisen.co.jp/schedule.html";

var HEADER = ["受付日時","種別","代表者","電話","メール","連絡方法","行く日","泊数","部屋数","大人","子供",
              "食事","来る船","送迎","エリア希望","希望の宿","やりたいこと","その他希望","備考","同意","1週間前通知","2日前通知"];

function doPost(e){
  try{
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    // 見出し行を常に最新に保つ（列を追加しても自動で揃う）
    var head1 = sh.getLastRow()===0 ? [] : sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0];
    if(head1.join("|") !== HEADER.join("|")) sh.getRange(1,1,1,HEADER.length).setValues([HEADER]);
    var kind = d.kind || "新規";
    sh.appendRow([new Date(), kind, d.name, d.tel, d.email, d.contact, d.checkin, d.nights, d.rooms, d.adults, d.kids,
                  d.meal, d.boat, d.pickup, d.area, d.yado, d.experiences, d.expOther, d.note, (d.agree||""), "", ""]);

    var details =
      "代表者　： " + d.name + "\n" +
      "電話　　： " + d.tel + "\n" +
      "メール　： " + (d.email || "（なし）") + "\n" +
      "希望の連絡方法： " + d.contact + "\n" +
      "行く日　： " + d.checkin + "　／　" + d.nights + "泊\n" +
      "部屋数　： " + d.rooms + "室\n" +
      "人数　　： 大人" + d.adults + "名・子供" + d.kids + "名\n" +
      "食事　　： " + d.meal + "\n" +
      "来る船　： " + d.boat + "\n" +
      "送迎　　： " + d.pickup + "\n" +
      "エリア希望： " + d.area + "\n" +
      "希望の宿： " + d.yado + "\n" +
      "やりたいこと： " + ((d.experiences || "") + (d.expOther ? "（その他:" + d.expOther + "）" : "") || "特になし") + "\n" +
      "備考　　： " + (d.note || "なし");

    var tag = (kind === "修正版") ? "・修正版" : "";

    // ① 幹事（自分）へ通知
    MailApp.sendEmail(NOTIFY_EMAIL, "【粟島 宿泊相談" + tag + "】" + d.name + "様",
      (kind === "修正版" ? "※修正版（前回の内容を更新したものです）\n\n" : "") +
      "新しい宿泊相談が届きました。\n\n" + details + "\n\n（フォームから自動送信されています）",
      {name:"粟島 宿泊相談フォーム", replyTo:(d.email || NOTIFY_EMAIL)});

    // ② 相手（申込者）へ受付控え（メールが入力されていれば）
    if(d.email){
      MailApp.sendEmail(d.email, "【粟島 宿泊相談を受け付けました" + tag + "】" + d.name + "様",
        d.name + " 様\n\n" +
        "粟島の宿泊相談を受け付けました。\n幹事が宿の空き状況を確認し、追ってご連絡します。\n\n" +
        "――― ご入力内容 ―――\n" + details + "\n\n" +
        "■ 内容を修正したいとき\n下記フォームからもう一度送信してください（最新版として受け付けます）。\n" + FORM_URL + "\n\n" +
        "※これは予約確定ではありません。確定後の変更・キャンセルは各自でお宿へ直接ご連絡ください。\n" +
        "※お心当たりがない場合はこのメールを破棄してください。",
        {name:"粟島 宿泊相談（幹事）", replyTo:NOTIFY_EMAIL});
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * リマインドメール：毎日1回トリガーで実行。
 * 行く日の「7日前」と「2日前」に、申込者へ時刻表確認の念押しメールを送る。
 * 送信済みは列に記録して二重送信を防止。
 */
function sendReminders(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh || sh.getLastRow() < 2) return;
  var vals = sh.getDataRange().getValues();
  var head = vals[0], col = {};
  head.forEach(function(h,i){ col[h] = i; });
  var tz = ss.getSpreadsheetTimeZone();
  var today = new Date(); today.setHours(0,0,0,0);

  for(var r = 1; r < vals.length; r++){
    var row = vals[r];
    var email = row[col["メール"]];
    var go = parseDate_(row[col["行く日"]]);
    if(!email || !go) continue;
    go.setHours(0,0,0,0);
    var days = Math.round((go - today) / 86400000);

    if(days === 7 && !row[col["1週間前通知"]]){
      sendReminderMail_(row, col, "1週間前", days, go, tz);
      sh.getRange(r+1, col["1週間前通知"]+1).setValue(new Date());
    }
    if(days === 2 && !row[col["2日前通知"]]){
      sendReminderMail_(row, col, "2日前", days, go, tz);
      sh.getRange(r+1, col["2日前通知"]+1).setValue(new Date());
    }
  }
}

function sendReminderMail_(row, col, whenLabel, days, go, tz){
  var name = row[col["代表者"]];
  var email = row[col["メール"]];
  var goTxt = Utilities.formatDate(go, tz, "M月d日(EEE)");
  var subject = "【粟島ご旅行リマインド】" + name + "様 " + goTxt + "ご出発（" + whenLabel + "）";
  var body =
    name + " 様\n\n" +
    "ご予定の粟島へのご旅行が近づいてきました（あと" + days + "日／" + goTxt + "）。\n\n" +
    "■ 出発前に、必ず最新の運航時刻をご確認ください\n" +
    "粟島汽船は日によって時刻が異なり、天候・海象により変更になることもあります。\n" +
    "当日の便を下記の時刻表で改めてご確認のうえ、くれぐれもお間違えのないようお越しください。\n" +
    "▼粟島汽船 時刻表（公式）\n" + FERRY_TIMETABLE_URL + "\n\n" +
    "――― ご予約内容 ―――\n" +
    "行く日　： " + row[col["行く日"]] + "　／　" + row[col["泊数"]] + "泊\n" +
    "来る船　： " + row[col["来る船"]] + "\n" +
    "人数　　： 大人" + row[col["大人"]] + "名・子供" + row[col["子供"]] + "名\n" +
    "希望の宿： " + row[col["希望の宿"]] + "\n" +
    "送迎　　： " + row[col["送迎"]] + "\n\n" +
    "※予約確定後の変更・キャンセルは各自でお宿へ直接ご連絡ください。\n" +
    "良い粟島旅を！🏝️";
  MailApp.sendEmail(email, subject, body, {name:"粟島 宿泊リマインド", replyTo:NOTIFY_EMAIL});
}

// "YYYY-MM-DD" 文字列 または Date を Date に変換
function parseDate_(v){
  if(!v) return null;
  if(Object.prototype.toString.call(v) === "[object Date]") return new Date(v.getTime());
  var m = String(v).match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? new Date(+m[1], +m[2]-1, +m[3]) : null;
}

// 動作確認用: デプロイ後にURLをブラウザで開くとこの文字が出ればOK
function doGet(){
  return ContentService.createTextOutput("粟島 宿泊相談 受付スクリプト 稼働中");
}
