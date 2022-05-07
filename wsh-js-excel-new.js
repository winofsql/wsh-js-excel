// ****************************
// 初期処理
// ****************************
var App = new ActiveXObject( "Excel.Application" );
var WshShell = new ActiveXObject( "WScript.Shell" );

// Excel を表示( 完成したらコメント化 )
App.Visible = true;
// 警告を出さない
App.DisplayAlerts = false;

// Minimize : GetSaveAsFilename を前面に出す為
// 本来、-4140 ですが WScript.Shell の Run と同じ 2 が使える
App.WindowState = 2

// ブック追加
App.Workbooks.Add();

 // ブックを取得( 一つしかないので、Count は 1 )
var Book = App.Workbooks( App.Workbooks.Count  );

Book.Sheets(1).Name = "最初のシート";
Book.Sheets.Add(null, Book.Sheets(1));
Book.Sheets(2).Name = "追加のシート";

// 先頭シートをアクティブにする
Book.Sheets(1).Activate();

// セルに値をセット
Book.Sheets(1).Cells(1, 1).Value = "社員名";
Book.Sheets(1).Cells(2, 1).Value = "山田　太郎甚左衛門";
Book.Sheets(1).Cells(3, 1).Value = "鈴木　一郎";
Book.Sheets(1).Cells(4, 1).Value = "佐藤　洋子";

// 列幅自動調整
Book.Sheets(1).Columns("A:A").EntireColumn.AutoFit();

// ****************************
// 参照
// 最後の 1 は、使用するフィルター
// の番号
// ****************************
var FilePath = App.GetSaveAsFilename(null,"Excel ファイル (*.xlsx), *.xlsx", 1);
if ( FilePath == false ) {
    WshShell.Popup( "Excel ファイルの保存選択がキャンセルされました" );
    App.Quit();
    App = null;
    WScript.Quit();
}

// ****************************
// 保存
// 拡張子を .xls で保存するには
// Book.SaveAs( FilePath, 56 )
// ****************************
try {
    Book.SaveAs( FilePath )
}
catch (error) {
    WshShell.Popup( "ERROR : " + error.description );
}

// 終了
App.Quit();

App = null;

WshShell.Run( "RunDLL32.EXE shell32.dll,ShellExec_RunDLL " + FilePath );
