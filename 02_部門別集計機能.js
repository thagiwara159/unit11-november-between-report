// 月次部門別レポート
function createDepartmentSummary(){
    try{
        // 売上データの取得関数呼び出し
        const data = getCurrentMonthData();
        // データをオブジェクト化して変数に格納、最後にreturnする
        const results = {};
        /** ⭐️ 配置取得じゃなくてオブジェクト取得だった
         * でもこれはこれでちゃんと値は取れていた
            for (let i = 1; i < data.length; i++){
                const row = data[i];

                const amount = row[4];
                const department = row[6];

                results.push({
                    department: department,
                    amount: amount,
                });
            }
        */

        // ⭐️ ループで部門データを巡回して部門別の合計金額算出       
        // 部門ごとに集計
        /** ⭐️ ここの計算方法が面白かった!　部門([6])を上から順に取り込んで、連続して
         * 同じ部門が続いていたらその部門の金額[4]を足してく、次に違う部門が来たらその部門は
         * その部門で足していくので各部門ごとに合計金額が算出できるという仕組み。
         * ⭐️ 部署ごとのカゴにその部署の金額を入れていくイメージ
         */
        data.forEach(row => {
            const dept = row[6];
            results[dept] = (results[dept] || 0) + row[4]; // ⭐️ ||は「論理 or (または)」
        }); // ⭐️ この処理は項目(今回は部署)が増えてもコードを変えないでOK

        // ⭐️ オブジェクトを返す
        Logger.log(results);
        return results;

       

    }catch(e){}
}

function reportCreateByDept(){
    try{
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const existing = ss.getSheetByName('部門別レポート');
        // 既存のシートがあれば削除
        if (existing) ss.deleteSheet(existing);
        // シート作成
        const sheet = ss.insertSheet('部門別レポート');
        // ⭐️ ヘッダーと集計結果を配置
        // ⭐️ 部門名と売上額を行ごとに整理
        const summary = calculateSummary();
        // 各項目設定
        sheet.getRange('A1').setValue(`${summary.period} 部門別レポート`);
        sheet.getRange('A3').setValue('部門');
        sheet.getRange('B3').setValue('売上');

        let row = 4;
        // ⭐️ 関数呼び出しのとこの（）を付け忘れててここの処理が飛んでた
        // ⭐️ 関数のあとに（）を付けないと関数そのものを呼び出してるだけ
        // ⭐️ 関数（）で呼び出して関数を実行させる
        Object.entries(createDepartmentSummary()).forEach(([dept, amount]) => {
            sheet.getRange(row, 1).setValue(dept);
            sheet.getRange(row, 2).setValue(amount + '円');
            row++
        });

        SpreadsheetApp.getUi().alert('部門別レポート作成完了！');

    }catch(e){}
}

// ⭐️ onOpenに'部門別レポート'機能追加
// ⭐️ 動作テスト、前回のコードともに動作するかチェック
/**
 * ⭐️ オブジェクトは｛｝、配列は［］、Mapは new Map で作る
 * ⭐️ || は「論理 or (または)」　どちらか1つでもtrueならOK
 * ⭐️
 */