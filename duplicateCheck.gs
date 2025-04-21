// 重複企業をあらかじめ確認し作業効率化、また同じ企業に営業をかけるミスを防ぐため、重複処理を実施する
// 【処理内容】
// 1.イベントリストサイトの１ページ分の企業名と詳細URLし取得し重複企業チェックシートのA・B列に貼り付け（別途、javascriptのスクリプトで取得後、手動で貼り付け）
// 2.企業名を整形し、アタックリストの企業名（B列）と照合
// 3.重複していたらA列の背景色を黄色に、C列に整形後の企業名を出力
// 3.重複していなければC列に整形後の企業名を出力、D列に検索リンクを出力


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