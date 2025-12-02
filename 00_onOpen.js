function onOpen(){
    SpreadsheetApp.getUi()
        .createMenu('月次リポート')
        .addItem('サマリー作成', 'createSummary')
        .addToUi();
}