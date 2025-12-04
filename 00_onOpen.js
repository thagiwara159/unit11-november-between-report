function onOpen(){
    SpreadsheetApp.getUi()
        .createMenu('月次レポート')
        .addItem('🗒️サマリー作成', 'createSummary')
        .addItem('📒部門別レポート','reportCreateByDept')
        .addItem('📜PDF変換','convertToPdf')
        .addItem('🧑‍🤝‍🧑配信先設定','setupRecipients')
        .addItem('📥️テストメール','sendTestEmail')
        .addItem('📨レポート配信','sendReportWithPDF')
        .addItem('⚙️自動実行設定','setupAutoExecution')
        .addItem('🔧全工程テスト','testFullProcess')
        .addToUi();
}