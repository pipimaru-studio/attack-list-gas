function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('データ整形')
    .addItem('アタックリストの会社名の一括整形', 'formatCompanyNames')
    .addToUi();

  // 「重複チェック」メニューを追加
  ui.createMenu('重複チェック')
    .addItem('重複企業チェックシートをリセット', 'resetDuplicateCheckSheet') 
    .addToUi();  

}
