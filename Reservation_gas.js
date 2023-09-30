const form = FormApp.getActiveForm();
const spreadSheet = SpreadsheetApp.openById("記録するスプシのID");//ごめんここだけ、自分たちで使うスプレッドシートのIDを設定お願いします！
const reserveSheet = spreadSheet.getSheetByName("予約");
const listSheet = spreadSheet.getSheetByName("日程");

// *--------------------*
//   フォーム送信時の処理
// *--------------------*
function receivedApplication(e) {

  // フォームの送信内容
  const email = e.response.getRespondentEmail();
  const items = e.response.getItemResponses();
  const name = ( items[0].getItem().getTitle() === "代表者氏名" ? items[0].getResponse() : "" );
  let result = "";
  const preferredTime = ( items[1].getItem().getTitle() === "希望する時間を選んでください" ? items[1].getResponse() : "" );
  const preferredMember = Number(( items[2].getItem().getTitle() === "人数を選んでください" ? items[2].getResponse() : "" ));
  console.log(name);
  console.log(preferredTime);
  console.log(preferredMember);

  if ( preferredTime ){
    // 日程シートチェック＆更新
    result = checkAvailability(preferredTime,preferredMember);
    // 予約シートに申込内容を書き込み
    if ( result === "OK" ){ reserveSheet.appendRow([name, preferredTime, preferredMember, result]); }
    // フォームを更新
    editForm();
  } else {
    result = "NG";
  }

  // メール送信
  sendEmail(email, name, preferredTime, preferredMember, result);

}

// *---------------------*
//   予約状況をチェックする
// *---------------------*
function checkAvailability(preferredTime,preferredMember){

  // listシートを配列に格納
  const list = listSheet.getDataRange().getValues();
  list.shift();

  // リストから予約日を探し、定員を確認する
  for ( let i = 0; i < list.length; i++ ){
    if ( list[i][0] == preferredTime ){
      // 予約済 + 予約人数　<= 定員であればOK
      console.log(list[i][2] + preferredMember)
      console.log(list[i][1])
      if ( list[i][2] + preferredMember <= list[i][1]){
        listSheet.getRange(i + 2, 3).setValue( list[i][2] +preferredMember);
        console.log("ここはだめ")
        return "OK";
      } else {
        console.log("ここならいい")
        return "NG";
      }
    }
  }

  // 希望日がリストに存在しなかった場合はNG
  return "NG"; 

}

// *----------------------*
//   予約結果をメール送信する
// *----------------------*
function sendEmail(email, name, preferredTime ,preferredMember, result){

  const mailTitle = "【未定研22】予約結果について";
  let mailBody;
  if ( result === "OK" ){
    mailBody = "予約が完了いたしました。\n"
             + `代表者氏名：${name} 様\n`
             + `予約時間：${preferredTime}\n`
             + `人数：${preferredMember}人\n`
             + "上映場所：九州大学大橋キャンパス５号館４階共同製図室\n"
             + "\n"
             + "ご予約ありがとうございます！\n\n"
             + "受付は予約時間の５分前から開始いたします。\n"
             + "当日受付にて本予約完了メールを確認いたしますので、ご準備をお願いいたします。\n"
             + "予約時間より５分以上遅れますとキャンセル扱いとなりますので、お気をつけください。\n\n"
             + "当日劇場にてお待ちしております！！\n"
             + "未定研一同\n\n"
             + "キャンセルの際は以下のフォームからお願い致します。\n"
             + "https://forms.gle/u97ZwsroxG8mWj846"
  } else {
    mailBody = "定員超過のため予約できませんでした。\n"
             + "お手数ですが、下記のフォームから再度ご予約をお願いいたします。\n"
             + form.getPublishedUrl();
  }

  // 結果メール送信
  GmailApp.sendEmail(email, mailTitle, mailBody);

}

// *-----------------*
//   フォームを更新する
// *-----------------*
function editForm(){

  let infoText = "";      // 「空き状況」部分のテキスト
  let choiceValues = [];  // 「参加希望日」部分の選択肢

  // 日程シートを配列に格納
  const list = listSheet.getDataRange().getValues();
  list.shift();

  // 日程シートのデータからフォームの内容を作成する
  for ( const record of list ){
    if ( record[2] < record[1] ){
      // 空きがある日程は「空き状況」に記載＋選択肢として設定
      infoText += `${record[0]} ： 残り ${record[1] - record[2]} 名\n`;
      choiceValues.push(record[0]);
    } else {
      // 空きがない日程は「空き状況」に満員を記載＋選択肢にはしない
      infoText += `${record[0]} ： 満員（申込不可）\n`;
    }
  }

  // 選択肢が１つもない場合、選択肢は「全日程申込不可」としフォームをクローズ
  if ( !choiceValues.length ) {
    choiceValues.push("全日程申込不可");
    form.setAcceptingResponses(false);
  }

  // フォームに変更を反映
  const items = form.getItems();
  items[2].setHelpText(infoText);
  items[3].asListItem().setChoiceValues(choiceValues);
  
}