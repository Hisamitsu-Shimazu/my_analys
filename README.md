# README

dataフォルダにcsvまたはpickleファイルを入れ、batファイルを実行することでoutputディレクトリに集計結果をまとめるシステム<br>
複数ファイルの同時実行可能が可能。<br>

# フォルダ構成
.<br>
├ data　データファイル格納用<br>
├ output　集計結果出力先<br>
│　　└ img<br>
├ script<br>
│　　├ module<br>
│　　│　　├ edit_excel.py　　# openpyxl操作用<br>
│　　│　　└ excel_style.py　　# セルスタイル設定<br>
│　　└ my_analys.py　　# mainスクリプト<br>
├ my_analys.bat　　# 実行用バッチファイル<br>
└ README.md<br>

# 実行手順
1.パスの設定
* Anacondaのactive.batまでのパス
	* C:\Users\ユーザー\Anaconda\Scripts\activate.bat
* my_analys.batまでのパス

2.my_analys.batファイルの実行
