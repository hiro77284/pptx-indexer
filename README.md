#機能
PowerPointファイルへの階層的ナンバリングを行う python スクリプトです。
章・節番号を振りたい部分に特殊なコードを書いた pptx ファイルをソースとして、変換スクリプトをかけることによりその特殊コード部分を章・節番号に変換します。
目次生成、引用生成も可能です。

#使い方
詳しい使用法は現在準備中です。
コマンドラインでpythonスクリプトを起動して使います。
pythonの実行環境構築やコマンドライン操作等は自力でできることが前提です。

使用例を example/simple ディレクトリに入れてあります。

## 標準的な使用手順は下記の通り。

(1) ソースとなる pptx ファイルを用意する
 例： example/simple/Example1_source.pptx

(2) パラメータファイルを用意する
 例： example/simple/simpleExample1.yaml

(3) インデックスを生成する
 例： $ python3 PpIndexCollector.py example/simple/simpleExample1.yaml
 生成物： example/simple/simpleExample1.json

(4) コンバージョンを行う
 例： $ python3 PpIndexLabeller.py example/simple/simpleExample1.yaml
 生成物： example/simple/Example1.pptx
 （これが最終成果物）

## パラメータファイルの書き方
準備中です

## その他もろもろ
準備中です
