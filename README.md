# SS3-Script
Super Build／SS3 (UNION SYSTEM Inc.) をVisual Basic Script で操作するためのスクリプト集です。
## Usage
### 非起動確認
SS3プログラムが起動していないことを確認します。  
スクリプト実行前に確認して二重起動を防ぐために使用します。
```vbs
Dim sh
Set sh = CreateObject("WScript.Shell")
result = sh.Run("ss3-script\SS3-非起動確認.vbs", 0, True)
'起動していなければ、result = 0
```
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
### 解析
解析を実行します。  
SS3が起動し、データを開いておく必要があります。  
ダイアログが開いていてはいけません。  
解析は「10.断面算定」までを行います。  
設定は全て事前に行っておく必要があります。
```vbs
Dim sh
Set sh = CreateObject("WScript.Shell")
result = sh.Run("ss3-script\SS3-解析.vbs", 0, True)
' 解析完了後、result = 0
```
### CSV入力
CSVファイルからデータの読み込みを行います。  
SS3が起動し、データを開いておく必要があります。  
ダイアログが開いていてはいけません。  
全ての項目を選択した状態で読み込みが行われます。
```vbs
Dim sh
Set sh = CreateObject("WScript.Shell")
result = sh.Run("ss3-script\SS3-CSV入力.vbs Path\to\input.csv", 0, True)
' 読み込み完了後、result = 0
```
### CSV出力
解析結果をCSVファイルに出力します。  
SS3が起動し、データを開いておく必要があります。  
ダイアログが開いていてはいけません。  
すべての解析結果を出力します。
```vbs
Dim sh
Set sh = CreateObject("WScript.Shell")
result = sh.Run("ss3-script\SS3-CSV出力.vbs Path\to\output.csv", 0, True)
' 出力完了後、result = 0
```
### メモ追記
データ管理用のメモに追記します。  
最終行を改行して追記します。  
SS3が起動し、データを開いておく必要があります。  
ダイアログが開いていてはいけません。
```vbs
Dim sh
Set sh = CreateObject("WScript.Shell")
result = sh.Run("ss3-script\SS3-メモ追記.vbs ""Comments to be appended""", 0, True)
' メモ追記後、result = 0
```
