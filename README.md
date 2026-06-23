# 請求書自動化ツール

このプロジェクトは、Excel テンプレートと Python を使って
請求書を自動生成するツールです。

## 機能
- input_data.xlsx の内容を読み込み
- invoice_template.xlsx を元に請求書を生成
- output_invoice フォルダに PDF と Excel を自動出力

## 使用方法
1. input_data.xlsx に請求先や品目を入力
2. ターミナルで以下を実行  
   `python generate_invoice.py`
3. output_invoice に請求書が生成されます

## ファイル構成
- generate_invoice.py  
- input_data.xlsx  
- invoice_template.xlsx  
- output_invoice（生成された請求書）
