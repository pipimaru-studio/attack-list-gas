function onEdit(e) {
  // 「mgmt」シートの実行判定欄が空欄の場合は処理を開始しない
  if (e.source.getSheetByName("mgmt").getRange("B11").getValue() === "") {
    Logger.log("トリガーが無効のため処理を終了します。");
    return;
  }

  const range = e.range;
  // 編集された値を取得
  const value = range.getValue(); // これならコピペでも大丈夫なことが多い
  // 編集されたシートを取得
  const sheet = range.getSheet();
   
  // 空文字の場合（値が取得できていなかった場合）は処理を行わない
  if (value === "") {
    Logger.log("値が取得できなかったため処理を終了します。");
    return;
  }

  // アタックリストが更新された場合：整形処理
  if (sheet.getName() === "アタックリスト"){ 
    formatOnEdit(e);
  }

  // 重複企業チェックシートのA列が更新された場合：重複チェック処理
  if (sheet.getName() === "重複企業チェック" && range.getColumn() === 1) {
    checkDuplicateOnEdit();
  }
}