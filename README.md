# トウキ合戦

Quizizz/Kahoot! などのゲームベースの学習プラットフォーム向けの問題ファイル（Excel）を生成する小道具です。

## 技術情報

1. 開発言語：**python3**

2. ライブラリー：

   `エクセル処理` **openpyxl**

   ```shell
   pip install openpyxl
   ```

   `データベース` **sqlite3**

## 利用方法

1. **input/** フォルダの下に、"**quizizz.xlsx**" または "**kahoot.xlsx**" の入力ファイルを入れておいてください。

2. システムが Windows の場合は、**Run-Win.bat** をダブルクリックしてください。

   システムが macOS の場合は、**Run-Mac.command** をダブルクリックしてください。

3. 実行が完了したら、**output/** フォルダの下に、入力ファイルと同じ名前のファイルが生成されました。

   このファイルを Quizizz/Kahoot! にアップロードして、ゲームを楽しみましょう！

## 参考記事

1. [openpyxl による Excelファイル操作方法のまとめ](https://gammasoft.jp/support/how-to-use-openpyxl-for-excel-file/) | GAMMASOFT
2. [python3でsqlite3の操作。作成や読み出しなどの基礎。](https://qiita.com/saira/items/e08c8849cea6c3b5eb0c)| saira
3. [How can I add the sqlite3 module to Python?](https://stackoverflow.com/questions/19530974/how-can-i-add-the-sqlite3-module-to-python) | falsetru

