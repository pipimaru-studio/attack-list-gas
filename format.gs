// Webサイトからのコピペの際、表記ゆれを手作業で修正する作業の効率化とミスを防ぐため、下記入力ルールに基づいて、企業リストに入力された入力値を変換する
// 入力内容（および処理内容）
// １．電話番号変換（F列）：
  // ・全角数字→半角
  // ・ハイフンを半角に統一
  // ・不要な記号・文字は除去
    // 「03(1234)5678」→「03-1234-5678」
    // 「01-2345-6789（代表）」→「01-2345-6789」
    // 「0312345678」→「03-1234-5678」
// ２．社名（B列）：
  // ・カタカナは全角カタカナに統一
  // ・ローマ字や数字は半角に
  // ・「(株)」などの略称は正式名称に変換
  // ・会社名前後のスペースも削除
    // 「アールズ 株式会社」 → 「アールズ株式会社」
    // 「（株）フューチャーリンク」 → 「株式会社フューチャーリンク」
    // 「ﾃｸﾉﾄﾞﾘｰﾑＶＩＳＩＯＮ株式会社」 → 「テクノドリームVISION株式会社」
// 3．全項目→不要な空白を削除
// 4．入力日を自動で出力


function formatOnEdit(e) {

  // アタックリストの列番号を設定
  const companyCol = 2; //会社名：B列 
  const phoneCol = 6; //電話：F列 
  const dateCol = 11; // 入力日：K列 = 11

  const range = e.range;
  // 編集された値を取得
  const value = range.getValue(); // これならコピペでも大丈夫なことが多い
  // 編集されたシートを取得
  const sheet = range.getSheet();
  //編集されたセルの行数を取得
  const row = range.getRow();
  //編集されたセルの列数を取得
  const col = range.getColumn();
  // 変換後の文字列の格納先
  let converted_text;


  Logger.log("入力値：" + value);
  
 
  // 「アタックリスト」シートのB列が編集された場合
    if (col === companyCol) {
      // 会社名変換を実行
      converted_text = convert_companyName(value);
    }
    // 「アタックリスト」シートのF列が編集された場合
    else if (col === phoneCol) {
      //電話番号変換関数を実行
      converted_text = convert_phonenumber(value);
    }
    // 「アタックリスト」シートのその他の項目が編集された場合
    else{
      if (typeof value === "string") {
        converted_text = value
            .replace(/[\u200B-\u200D\uFEFF]/g, '')  // ゼロ幅スペースなど不可視文字の削除
            .replace(/　/g, '')                     // 全角スペースを削除
            .replace(/ /g, '')                      // 半角スペースを削除
            .replace(/\s+/g, '');                   // 改行・タブなど含む空白文字すべて削除
        converted_text = converted_text.trim();
        } else {
          // 文字列以外（数値や日付など）の場合はそのまま使う
          converted_text = value;
        }
    }
    
    // 変換後の値を編集されたセルに再出力
    sheet.getRange(row, col).setValue(converted_text);
    
    Logger.log("出力値：" + converted_text);
    
    // 入力日が空欄の場合、出力
    const dateCell = sheet.getRange(row, dateCol);
    if (dateCell.isBlank()) {
      dateCell.setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
    }

    //処理終了 
    return;

}


// アタックリストシートの会社名を一括で整形する（会社名変換のコード改修後、既存データの企業名を修正するため！）
function formatCompanyNames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("アタックリスト");
  const values = sheet.getRange("B2:B" + sheet.getLastRow()).getValues();
  const newValues = values.map(row => {
    const name = row[0];
    if (!name) return [""];
    return [convert_companyName(name)];
  });
  sheet.getRange("B2:B" + sheet.getLastRow()).setValues(newValues);
  SpreadsheetApp.getUi().alert("会社名の整形が完了しました！");
}