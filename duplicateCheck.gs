// アタックリストに記載する前に重複企業をあらかじめ確認し作業効率化、また同じ企業に営業をかけるミスを防ぐため、重複処理を実施する
// 【処理内容】１は手動実行。２〜３が自動化範囲。
// 1.データベースサイト（リサーチ対象の企業情報が載ったWebサイト）の１ページ分の企業名と詳細URLを取得。
    // 重複企業チェックシートのA・B列に貼り付け
// 2.企業名を整形し、アタックリストの企業名（B列）と照合
// 3.重複していたらA列の背景色を黄色に、C列に整形後の企業名を出力
// 3.重複していなければC列に整形後の企業名を出力、D列に検索リンクを出力


// ----------------【重複チェックシートのリセット】---------------- 
// 処理内容１の前に重複チェックシートのリセットを行う（カスタムメニューより呼び出す）
function resetDuplicateCheckSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("重複企業チェック");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;  // データがなければ終了

  // 2行目以降のA〜D列のデータを削除
  sheet.getRange("A2:D" + lastRow).clearContent();

  // 背景色だけをデフォルト（null)に戻す
  const range = sheet.getRange("A2:D" + lastRow);
  const defaultBackgrounds = Array.from({ length: lastRow - 1 }, () => [null, null, null,null]);
  range.setBackgrounds(defaultBackgrounds);
}

// ----------------【重複チェック】---------------- 
function checkDuplicateOnEdit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("重複企業チェック");
  const attackSheet = ss.getSheetByName("アタックリスト");

  // 「重複企業チェック」シートのA列（企業名）を取得
  const companyNames = sheet.getRange("A2:A" + sheet.getLastRow()).getValues();
  
  // 「アタックリスト」シートのB列（企業名）を取得し、空白を除外 → 整形
  const existingCompanies = attackSheet.getRange("B2:B" + attackSheet.getLastRow())
    .getValues()
    .flat()                             // 二次元配列を一次元に
    .filter(String)                     // 空文字や null を除外
    .map(name => name ? convert_companyName(name) : ""); // 整形（関数で変換）

  // 「重複企業チェック」シートの企業名リストを1件ずつ処理
  for (let i = 0; i < companyNames.length; i++) {
    let name = companyNames[i][0];
    
    // 企業名が空白ならスキップ
    if (!name) continue;
    
    // 整形処理をしてC列に表示
    let converted = convert_companyName(name);
    sheet.getRange(i + 2, 3).setValue(converted); // 整形後の会社名をC列にセット

    // アタックリスト内の企業名と照合（重複チェック）
    if (existingCompanies.includes(converted)) {
      sheet.getRange(i + 2, 1).setBackground("yellow"); // 重複：A列を黄色に
    } else {
      // 重複していない場合：Google検索URLをD列にセット
      const searchUrl = "https://www.google.com/search?q=" + encodeURIComponent(converted + " 公式サイト");
      sheet.getRange(i + 2, 4).setValue(searchUrl);
    }
  }
}
