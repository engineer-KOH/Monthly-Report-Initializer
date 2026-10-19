# 🚀 Monthly-Report-Initializer

## 📌 概要（Overview）
このシステムは、月次および日々の成果報告に用いるExcelファイル（マクロ付きマスター）の初期化・セットアップを自動化するためのPythonツールです。  
カレンダー計算やテンプレートの動的配置を自動化し、毎月の面倒なレポート準備にかかる工数を削減します。

## 🎯 主な目的（Objective）
- **業務効率化**: 月初や日々の報告書フォーマットの準備・初期化作業の自動化
- **ヒューマンエラー防止**: シート名や日付、マスター情報の転記ミス・設定漏れの排除
- **マクロ連携**: VBAマクロを含んだExcelテンプレート（`.xlsm`）とのシームレスな統合

## 🛠️ 使用技術（Tech Stack）
- **言語**: Python
- **主要ライブラリ**: `openpyxl`, `pathlib`, `datetime` など
- **対象ファイル**: マクロ付きExcelワークブック（`.xlsm`）

## ⚙️ システム構成・処理フロー（Workflow）
- **入力 (Inputs)**: マクロ付きExcelマスターテンプレート（成果報告日報_氏名_マスター.xlsm）の読み込み
- **処理 (Process)**: Pythonによる日付・カレンダー情報の自動計算およびExcelシート構造の初期化・動的更新
- **出力 (Outputs)**: 新月・新期間用に適切に初期化・セットアップされた成果報告用Excelファイルの生成

## 🚀 使い方（Quick Start）
- **【開発者・セットアップ手順】**
-   **1.リポジトリのクローン**: 「git clone [https://github.com/mochapaparocku-crypto/Monthly-Report-Initializer.git](https://github.com/mochapaparocku-crypto/Monthly-Report-Initializer.git)」
-   **2.依存ライブラリのインストール**: 「pip install -r requirements.txt」
-   **3.スクリプトの実行**: 「python system/init_daily_report.py」

- **【エンドユーザー（実務担当者）の日常的な利用フロー】**
-   **1.成果報告日報_氏名_マスター.xlsm をExcelで開きます。**
-   **2.シート上に配置されている「日報生成」ボタン（VBAマクロ）をクリックします。**
-   **3.自動的に当月の日付やレイアウトが初期化され、その日の日報作成を即座に開始できます。**

## 📁 フォルダ構成（Repository Structure）
```text
Monthly-Report-Initializer/
├── system/
│   └── init_daily_report.py   # 日報・成果報告の初期化を行うメインスクリプト
├── 成果報告日報_氏名_マスター.xlsm  # 雛形となるマクロ付きExcelマスターファイル
└── .gitignore                 # 機密情報・不要ファイルの除外設定
