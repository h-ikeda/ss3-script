# SS3-Script
ユニオンシステム株式会社 (UNION SYSTEM Inc.) 製の一貫構造計算ソフトウェア (プログラム) である、Super Build／SS3 を Visual Basic Script で操作するための非公式スクリプト集です。
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
解析終了後に、出力画面は表示されません。
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
### 出力したCSVファイルの読み取り
CSV出力で出力したCSVファイルからデータを読み取る関数を利用できます。  
関数を利用するには、ExecuteGlobal関数で予めスクリプトの内容を読み込んでおきます。
```vbs
Dim lib
Set lib = WScript.CreateObject("Scripting.FileSystemObject").OpenTextFile("ss3-script\SS3-CSV読取.vbs", 1)
ExecuteGlobal lib.ReadAll
lib.Close
```
#### CSV_getData(title, path)
`path`で指定したCSVファイルから、`title`で指定した項目の内容を読み取ります。  
返り値は配列の配列です。
```vbs
Dim result = CSV_getData("節点変位", "path\to\output.csv")
WScript.Echo "節点No." & result(0)(0) & "のX方向の変位は、" & result(0)(5) & "cmです。"
```
#### CSV_getDataHeader(title, path)
`path`で指定したCSVファイルから、`title`で指定した項目のヘッダーを読み取ります。  
返り値は配列です。
```vbs
Dim header = CSV_getDataHeader("節点変位", "path\to\output.csv")
WScript.Echo "節点変位の最初の項目は、" & header(0)
```
## Contact
ikeda_hiroki@icloud.com
