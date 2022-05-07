// ****************************
// 初期処理
// ****************************
WScript.Echo( "処理を開始します" );
var WshShell = new ActiveXObject("WScript.Shell");
var ExcelApp = new ActiveXObject( "Excel.Application" );

// デバッグ時は、Excel の本体を表示させて状況が解るようにする
ExcelApp.Visible = true;
// UI でチェックさせるようなダイアログを表示せずに実行する
ExcelApp.DisplayAlerts = false;

try {

    // ****************************
    // ブック追加
    // ****************************
    var Book = ExcelApp.Workbooks.Add();

    // 通常一つのシートが作成されています
    var Sheet = Book.Worksheets( 1 );

    // ****************************
    // シート名変更
    // ****************************
    Sheet.Name = "JScriptの処理";

    // ****************************
    // セルに値を直接セット
    // ****************************
    for( var i = 1; i <= 10; i++ )
    {
        Sheet.Cells(i, 1) = "処理 : " + i;
    }

    // ****************************
    // 1つのセルから
    // AutoFill で値をセット
    // ****************************
    Sheet.Cells(1, 2) = "子";
    // 基となるセル範囲
    var SourceRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(1,2));
    // オートフィルの範囲(基となるセル範囲を含む )
    var FillRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(10,2));
    SourceRange.AutoFill(FillRange);

    // ****************************
    // 保存
    // ****************************
    Book.SaveAs( WshShell.CurrentDirectory + "\\sample.xlsx" );

} catch (error) {
    ExcelApp.Quit();
    ExcelApp = null;
    WshShell.Popup(error.description);
    WScript.Quit();	
}

ExcelApp.Quit();
ExcelApp = null;

// ****************************
// ファイルの最後
// ****************************
WshShell.Popup("処理を終了します");
