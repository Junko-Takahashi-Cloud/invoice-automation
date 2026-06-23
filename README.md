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

## このプロジェクトについて
Excel と Python を使って、請求書を自動生成するツールです。
手作業での請求書作成を効率化し、作業時間を大幅に短縮できます。

## 特徴
- input_data.xlsx に入力した内容を自動で読み込み
- invoice_template.xlsx を元に請求書を生成
- output_invoice フォルダに PDF と Excel を自動出力
- 小規模事業者や個人事業主の請求書作成を想定

## 使い方
1. input_data.xlsx に請求先や品目を入力  
2. ターミナルで以下を実行  
   `python generate_invoice.py`  
3. output_invoice に請求書が生成されます
