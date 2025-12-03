function onOpen(){
    SpreadsheetApp.getUi()
        .createMenu('月次リポート')
        .addItem('サマリー作成', 'createSummary')
        .addItem('部門別レポート','reportCreateByDept')
        .addToUi();
}