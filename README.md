# NAME

kg-merge - Microsoft Forms の応答ファイルに学生の学籍番号を記入する

# DESCRIPTION

**kg-merge** は、Microsoft Excel の XLSX 形式で保存された Microsoft Form の応
答ファイルに、回答者の学籍番号を記入するスクリプトです。Microsoft Form からダ
ウンロードできる XLSX 形式の「回答ファイル」の D 列に記録されている回答者の ID
(Microsoft 365 の ID) を学生名簿 (「名簿ファイル」)から探索し、学生名簿に ID
が存在すれば C 列に学籍番号を記入します。C 列 (もともとは回答完了時刻が記録さ
れています) が上書きされることに注意してください。変更された回答ファイルは
「filename-merged.xlsx」というファイル名で保存されます。名簿ファイルは LUNA の
「名簿ダウンロード」からダウンロードされた、タブ区切りのテキストファイル (ただ
し拡張子は .XLS) を前提としています。

# REQUIREMENTS

- Python 3
- openpyxl (https://pypi.org/project/openpyxl/)

# INSTALLATION

```sh
pip3 install openpyxl
./kg-merge
```

Windows 10 (x64) 用のバイナリを用意しました。PyInstaller で単一の実行ファイル
に変換したものです。単一バイナリの実行ファイルです。PyInstaller でパッケージ化
していますので、Python のインストールや openpyxl モジュールのインストールは不
要です。

https://github.com/h-ohsaki/kg-merge/raw/master/win10-x64/kg-merge.exe

# USAGE

実行するとファイル選択ダイアログが開きますので、まず、LUNA からダウンロードし
た名簿ファイル (meibo-\*.xls) を選択してください。次に、再度ファイル選択ダイア
ログが開きますので、Microsoft Forms からダウンロードした応答ファイル (\*.xlsx)
を選択してください。応答ファイルと同じディレクトリに \*-merged.xlsx というファ
イルが作成されます。

# AVAILABILITY

最新版の **kg-merge** は https://github.com/h-ohsaki/kg-merge から入手できます。

# AUTHOR

Hiroyuki Ohsaki (ohsaki[atmark]lsnl.jp)
