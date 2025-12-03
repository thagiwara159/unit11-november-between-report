// 売上データの取得関数
function getCurrentMonthData(){
    try{
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sh = ss.getSheetByName('売上データ');
        // const lastRow = sh.getLastRow(); この2行は可読性を上げるために不要になった
        // const lastCol = sh.getLastColumn();

        // ヘッダーを除いてデータを最後まで取得
        // ⭐️ const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
        // 　　↑ここは自分で書いたがもっと可読性の良い書き方があった
        const data = sh.getDataRange().getValues(); 
        // ⭐️ ↑この書き方でデータを取得して後でSliceを使ったほうが可読性UP、変数も減らせる

        // 過去データを扱うため仮の日付を設定
        const fakeToday = new Date(2025,4,1); // 月は0スタートなので5月の場合は4にする
        
        /*const result = data.filter(row => {
         *    const rowData = new Date(row[0]);
         *   return rowData.getMonth() === fakeToday.getMonth(); 
         * }); ここも書いたがこの書き方だど年を比較してないので何年もデータがあった場合他の年の5月の
         *   データもとってきてしまう     あと、Sliceを使って書いたほうが可読性が上がるので書き換える
        */
        // 取得したデータを順番に入れて年と月を比較して === で厳密比較
        return data.slice(1).filter(row => {
            const date = new Date(row[0]); // [0]データの1番左の日付を年月比較のために変数に取得
            return date.getMonth() === fakeToday.getMonth() && // 月の比較
                date.getFullYear() === fakeToday.getFullYear(); // 年の比較
        });
        // ⭐️ データは全部取れた、あとは総売上と件数を計算する集計処理を足す
        // ⭐️ 集計は別の関数で。　ここの関数とは目的が違うから
    }catch(e){}
}

// 基本集計
function calculateSummary(){
    try{
        // 売上データ取得の関数を呼び出して変数に売上データを取得
        const data = getCurrentMonthData();

        /**
         * ⭐️ .reduce の説明　今回の場合は sum(累積値) 、dataの中のデータの row[4](値) を
         * sum の中に入れて、次のrow[4]の値を sum に入れて加算する sum は累積値なので
         * row[4] の値を次々受け入れ加算してくので値の合計値を出すのに使える　
         * ちなみに最後の 0 は sum の初期化、sum は 0 からスタートしますよの意味
         */
        const totalSales = data.reduce((sum, row) => sum + row[4], 0);
        
        // 売上データの取得した数を変数に代入 
        const totalCount = data.length;

        // 過去データを扱うため仮の日付を設定
        const fakeToday = new Date(2025,4).toLocaleDateString(
            'ja-jp', {year: 'numeric', month: 'long'
        }); // 月は0スタートなので5月の場合は4にする　出力値は、2025年5月
        
        // ⭐️ totalSales dataのrow[4]の金額の合計値、totalCount 取得したデータの量
        // period(ピリオド) 実際には集計日を出力したいが、今回は過去データのため
        // fakeToday を採用  以上をオブジェクトにして返してる
        return {
            totalSales: totalSales,
            totalCount: totalCount,
            period: fakeToday
        };
    }catch(e){}
}

function createSummary(){
    try{
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        // 基本集計関数を呼び出し変数に入れる
        const summary = calculateSummary();
        const existing = ss.getSheetByName('月次サマリー');
        // ⭐️ サマリーシートの作成、シートがあれば削除して常に最新状態を保つ
        // ⭐️ セルに集計結果（期間、総売上、件数）を見やすく配置する処理
        // ↓月次サマリーシートがあれば最新状態を保つために削除
        if (existing) ss.deleteSheet(existing);
        // 最新状態にするため月次サマリーシートを新たに作成、ヘッダー設定とデータ出力
        const sheet = ss.insertSheet('月次サマリー');
        sheet.getRange('A1').setValue(`${summary.period} 売上サマリー`);
        sheet.getRange('A3').setValue('総売上');   
        sheet.getRange('B3').setValue(summary.totalSales + '円');   
        sheet.getRange('A4').setValue('件数');   
        sheet.getRange('B4').setValue(summary.totalCount + '件');
             
        SpreadsheetApp.getUi().alert('サマリー作成完了！');

    }catch(e){}
}

// ⭐️ 集計結果が正しく動作するか確認、スプシのメニューから実行したいonOpen
// ⭐️ エラーがないかチェック　try-catch を入れる 解説を見た結果どこにもtry-catchはなかった

/**
 *      sheet.getRange('A1').setValue(`${summary.period} 売上サマリー`);
             .getRange('A3').setvalue('総売上');   
             .getRange('B3').setvalue(summary.totalSales + '円');   
             .getRange('A4').setvalue('件数');   
             .getRange('B4').setvalue(summary.totalCount + '件');
        ⭐️ GASでは一見できそうな書き方だけど、この書き方はNGです！
 */
/**
 * ⭐️ シートからすべてのデータを取ってきて比較する時、全データにたいしてSlice[1]するとヘッダーを除ける
 * ⭐️ .reduce を使って累計計算を行える 例：.reduce((sum, row) => sum + row[4], 0);
 * ⭐️ ↑説明 sum を 0 にしてから row[4] の値を追加して足し算の繰り返し sum にはどんどん値が溜まっていく
 * ⭐️ シートがあれば削除　例：if (sheet) ss.deleteSheet(sheet);
 * ⭐️ シートの作成　例：const sheet = ss.insertSheet('作成するシート名')
 * ⭐️ 日付データの日本語化（2025年5月1日）
 *      .toLocaleDateString('ja-jp', {year: 'numeric', month: 'long', day: 'numeric'});
*/