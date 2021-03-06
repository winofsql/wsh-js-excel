var App = new ActiveXObject("Excel.Application");
var WshShell = new ActiveXObject( "WScript.Shell" );

// Excel を表示( 完成したらコメント化 )
App.Visible = true;
// 警告を出さない
App.DisplayAlerts = false;

// Minimize : GetSaveAsFilename を前面に出す為
// 本来、-4140 ですが WScript.Shell の Run と同じ 2 が使える
App.WindowState = 2

// ****************************
// 一つのファイルを開く
// ****************************
// https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.getopenfilename
var FilePath = App.GetOpenFilename("Excel ファイル (*.xlsx), *.xlsx,全て,*.*", 1,"ファイルの選択",null, false );
// 未選択の場合
if( FilePath === false ) {
    WshShell.Popup("ファイルの参照選択がキャンセルされました")
    App.Quit();
    App = null;
    WScript.Quit();
}

// ****************************
// ブックを開く
// ****************************
var Book = App.Workbooks.Open(FilePath);

// ****************************
// 最終シートを前にコピー
// ****************************
Book.Sheets(Book.Sheets.Count).Copy( Book.Sheets(Book.Sheets.Count) );

// コピーしたので アクティブになります
var Target = Book.ActiveSheet;

// 先頭列を文字列に設定
Target.Range("A:A").NumberFormatLocal = "@";

// セルに値をセット
Target.Cells(1, 1).Value = "社員コード";
Target.Cells(2, 1).Value = "0001"
Target.Cells(3, 1).Value = "0002";
Target.Cells(4, 1).Value = "0003";
Target.Cells(1, 2).Value = "社員名";
Target.Cells(2, 2).Value = "山田　太郎甚左衛門";
Target.Cells(3, 2).Value = "鈴木　一郎";
Target.Cells(4, 2).Value = "佐藤　洋子";

// 列幅自動調整
Target.Columns("B:B").EntireColumn.AutoFit();

// ****************************
// 保存
// ****************************
try {
    Book.SaveAs( FilePath );
    Book.Close();
}
catch (error) {
    WshShell.Popup( "Book.SaveAs でエラーが発生しました : " + error.description );
}

// 終了
App.Quit();

App = null;

WshShell.Run( "RunDLL32.EXE shell32.dll,ShellExec_RunDLL " + FilePath );
