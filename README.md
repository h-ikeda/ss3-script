# SS3-Script
Super Build／SS3 (UNION SYSTEM Inc.) をVisual Basic Script で操作するためのスクリプト集です。
## Usage
### 起動確認
SS3プログラムが正常に起動完了していることを確認します。  
起動時にデータを開いている必要があります。  
ライセンス確認では常に "はい(Y)" を選択します。
```vbs
Dim sh
Set sh = CreateObject("WScript.Shell")
sh.Run "Path\to\SS3Data\START.SS3", 1, False
result = sh.Run("ss3-script\SS3-起動確認.vbs", 0, True)
' 起動完了後、result = 0
```
### 終了
SS3プログラムを終了します。  
ダイアログが開いていてはいけません。
```vbs
Dim sh
Set sh = CreateObject("WScript.Shell")
result = sh.Run("ss3-script\SS3-終了.vbs", 0, True)
' 終了後、result = 0
```
