# 提出物をチェックするためのPythonスクリプト
## 説明
プログラミング授業での学生の提出物をチェックするためのPythonスクリプトです。

## インストール
git clone https://github.com/(上げないかも)

## 必要なモジュール（要インストール）
pip install openpyxl
## 必要なモジュール（標準モジュール）
sys
os
difflib
### 4. 仕様
コマンドライン引数から答えのファイルと提出ディレクトリを受け取ります。
引数が正しくない場合はエラーメッセージを表示します。
提出されたファイルと答えのファイルを比較し、内容が完全に一致する場合は "〇"、そうでない場合は "△" と判定します。
チェック結果をExcelファイル（.xlsx）に出力します。
同時にコンソールにも結果を表示します。
### 5. 使い方
python assignment_checker.py 答えのファイル.py rensyu\提出ディレクトリ
python assignment_checker.py 答えのファイル.py  kadai\提出ディレクトリ

コマンドライン引数の第1引数に正解のファイル名
コマンドライン引数の第2引数に課題が入っている（一括ダウンロードした）フォルダ名
## 質問
質問・改善点がある場合は、メールアドレスまたはTeamsまでご連絡ください。