var App = new ActiveXObject("Excel.Application");
var WshShell = new ActiveXObject( "WScript.Shell" );

// Excel を表示( 完成したらコメント化 )
App.Visible = true;
// 警告を出さない
App.DisplayAlerts = false;

// Minimize : GetSaveAsFilename を前面に出す為
// 本来、-4140 ですが WScript.Shell の Run と同じ 2 が使える
App.WindowState = 2

// 一つのファイルを開く
// https://docs.microsoft.com/ja-jp/office/vba/api/excel.application.getopenfilename
var filePath = App.GetOpenFilename("全て,*.*,CSV,*.csv", 1,"ファイルの選択",null, false );
// 未選択の場合
if( filePath === false ) {
    WshShell.Popup("ファイルの参照選択がキャンセルされました")
}
// 選択の場合
else {
    WshShell.Popup(filePath + " を選択しました");
}

// 終了
App.Quit();

App = null;
