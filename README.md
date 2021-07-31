# wsh-js-excel-open
ブックを参照して開いてセルに値をセットして Excel で起動
```xml
<?xml version="1.0" encoding="utf-8" ?>
```
GitHub 用に utf-8 で記述する為に wsf 形式を使用
```javascript
var App = new ActiveXObject( "Excel.Application" );
var WshShell = new ActiveXObject( "WScript.Shell" );
```
HTML アプリケーションでもコピペで使いたいので new ActiveXObject を使用
