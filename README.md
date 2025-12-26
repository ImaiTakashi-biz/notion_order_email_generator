# Notion 注文書メール自動作成アプリ

Notionで管理している注文データを元に、Excelテンプレートを使用して注文書のPDFを自動作成し、担当者へメールで送信するデスクトップアプリケーションです。

## 主な機能

- Notionデータベースから「要発注」ステータスのデータを取得
- **部署名によるデータフィルタリング機能**
- 仕入先ごとに注文内容をグループ化
- 指定したExcelテンプレートを元に注文書PDFを自動作成
- 仕入先ごとにPDFを添付したメールを作成し、SMTP経由で送信
- 複数の送信元メールアカウントの切り替えに対応
- 処理済みのNotionデータに「発注日」を自動で記録
- **仕入先が未設定のデータを検出し、GUI上で警告**

## 動作要件

- **OS:** Windows
- **ソフトウェア:** Microsoft Excel
  - 本アプリケーションは内部でExcelを操作してPDFを作成するため、Excelのインストールが必須です。

## セットアップ手順

### 1. 依存ライブラリのインストール

以下のコマンドを実行して、必要なPythonライブラリをインストールします。

```bash
pip install -r requirements.txt
```

### 2. 環境変数ファイル (`.env`) の作成

プロジェクトのルートディレクトリに `.env` という名前のファイルを作成し、以下の内容を参考に、ご自身の環境に合わせて設定してください。

```dotenv
# Notion APIのインテグレーションシークレット
NOTION_API_TOKEN="secret_..."

# 注文管理DBのID
NOTION_DATABASE_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# 仕入先管理DBのID
NOTION_SUPPLIER_DATABASE_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# 注文書Excelテンプレートのフルパス
EXCEL_TEMPLATE_PATH="C:\path\to\your\template.xlsx"

# 作成したPDFの保存先フォルダのフルパス
PDF_SAVE_DIR="C:\path\to\save\pdfs"
```

### 3. アカウント設定ファイル (`email_accounts.json`) の作成

プロジェクトのルートディレクトリに `email_accounts.json` という名前のファイルを作成し、メールアカウントや部署の情報を設定します。

**ファイル構成の例:**
```json
{
  "smtp_server": "smtp.office365.com",
  "smtp_port": 587,
  "accounts": {
    "kato_aya": {
      "display_name": "加藤 彩",
      "sender": "a.katou@example.com",
      "password": "password_kato"
    },
    "taro_yamada": {
      "display_name": "山田 太郎",
      "sender": "yamada@example.com",
      "password": "password_yamada"
    }
  },
  "department_defaults": {
    "生産部": "kato_aya",
    "営業部": "taro_yamada"
  },
  "departments": [
    "生産部",
    "品質保証部",
    "営業部",
    "総務部"
  ]
}
```

- **`accounts`**:
  - アプリケーションで使用するメールアカウントを複数登録できます。
  - `"kato_aya"` のようなキー名は、他の人と重複しない**一意の名称**にしてください。
  - `display_name`はアプリのUIに表示される名前、`sender`はメールアドレス、`password`はパスワードです。
- **`department_defaults`**:
  - 部署名と、その部署が選択されたときにデフォルトで設定されるアカウントのキーを紐付けます。
- **`departments`**:
  - GUIの「部署名フィルター」に表示される部署のリストです。

## 実行方法

### 開発環境での実行

以下のコマンドでアプリケーションを起動します。

```bash
python main.py
```

### 配布用実行ファイルの作成（PyInstaller）

PyInstallerを使用して、単一の実行ファイル（.exe）としてビルドできます。

#### 1. PyInstallerのインストール

```bash
pip install pyinstaller
```

#### 2. 実行ファイルのビルド

以下のいずれかでビルドします（生成されるexe名は `OrderMailer.exe` 固定で、バージョンはexe名に含めません）。

- `build.bat` を実行
- または、以下のコマンドを実行

```bash
pyinstaller "OrderMailer.spec"
```

`OrderMailer.spec` に、配布に必要なファイル（`.env` / `email_accounts.json` / `app_icon.ico`）の同梱設定が含まれています。

**注意:** ビルド前に、プロジェクトルートに`.env`、`email_accounts.json`、`app_icon.ico`が存在することを確認してください。

ビルドが完了すると、`dist/OrderMailer.exe`が生成されます。

#### 3. 配布時のファイル構成

配布時は、実行ファイルのみを配布します。設定ファイル（`.env`と`email_accounts.json`）は実行ファイルに含まれています。

```
配布フォルダ/
└── OrderMailer.exe        # 実行ファイル（設定ファイルを含む）
```

**重要:**
- 配布物は`.exe`ファイル1つだけです
- `.env`と`email_accounts.json`は実行ファイルに含まれているため、別途配布する必要はありません
- 実行ファイルを実行すると、一時ディレクトリに設定ファイルが展開され、アプリケーションが読み込みます

## バージョン

現在のバージョン: `v1.0.0`（ビルド日: `2025-12-25`）

### バージョン管理方針
本アプリは Semantic Versioning に準拠し、
以下の形式でバージョンを管理します。

vメジャー.マイナー.パッチ（例：v1.2.3）

- メジャー：大きな仕様変更（DB構造変更など）
- マイナー：機能追加・画面追加
- パッチ：バグ修正・軽微な調整

### リリース手順
1. `version.py` の `APP_VERSION` を更新
2. `CHANGELOG.md` に変更内容を追記
3. exeをビルド（ファイル名は固定）
4. 配布・更新

## GUIの操作方法

1.  **部署名の選択:** 画面上部の「部署名フィルター」で、データを絞り込みたい部署を選択します（複数選択可）。
2.  **送信者アカウントの確認:** 部署を選択すると、デフォルトの送信者アカウントが自動で設定されます。必要に応じて、ドロップダウンリストから別のアカウントに変更することも可能です。
3.  **データ取得:** 「Notionからデータを取得」ボタンを押すと、Notionから発注対象のデータを取得し、左側のリストに仕入先名が表示されます。
4.  **仕入先の選択:** 左側のリストから仕入先を選択すると、右側のテーブルに発注内容が表示され、PDFの作成がバックグラウンドで開始されます。
5.  **プレビューと送信:** PDFの作成が完了すると、画面下部に宛先や担当者、添付ファイル名が表示されます。内容を確認し、問題がなければ「メール送信」ボタンを押してください。
6.  **Notionの更新:** メール送信後、Notionの対象ページの「発注日」を更新するか確認ダイアログが表示されます。「はい」を選択すると、発注日が今日の日付で記録されます。

## ファイル構成

```
/
├── .gitignore
├── main.py                    # アプリケーションのエントリーポイント
├── app_gui.py                 # 後方互換性のための統合ファイル
├── config.py                  # 設定情報(.env, .json)の読み込み、定数管理
├── email_service.py           # メール作成・送信処理
├── notion_api.py              # Notion APIとの連携処理
├── pdf_generator.py           # Excelテンプレートからの注文書PDF生成処理
├── settings_gui.py            # 設定画面のGUIとロジック
├── logger_config.py           # ロギング設定モジュール
├── cache_manager.py           # Notionデータ取得のキャッシュ管理
├── requirements.txt           # 依存ライブラリリスト
├── README.md                  # このファイル
├── CHANGELOG.md               # 変更履歴
├── version.py                 # バージョン情報の一元管理
├── controllers/               # コントローラーモジュール
│   ├── __init__.py
│   └── app_controller.py     # アプリケーションのメインコントローラー
├── ui/                        # UIコンポーネントモジュール
│   ├── __init__.py
│   ├── queue_io.py           # 標準出力のキューリダイレクト
│   ├── top_pane.py           # 上部UI（部署フィルター、アカウント選択）
│   ├── middle_pane.py        # 中央UI（仕入先リスト、注文データテーブル）
│   └── bottom_pane.py       # 下部UI（プレビュー、ログ表示）
└── tests/                     # 自動テストコード
    ├── test_email_service.py
    ├── test_notion_api.py
    └── test_pdf_generator.py
```

## 注意事項
- 本アプリケーションは`win32com`ライブラリを使用しているため、**Windows環境でのみ動作**します。
- `email_accounts.json` にメールアカウントのパスワードを直接記述します。ファイルの取り扱いには十分ご注意ください。
- **配布時**: PyInstallerでビルドした実行ファイルには、`.env`と`email_accounts.json`が含まれています。ビルド前にこれらのファイルを正しく設定しておく必要があります。設定が不適切な場合、アプリケーションは起動時にエラーダイアログを表示して終了します。
