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

以下のコマンドでアプリケーションを起動します。

```bash
python main.py
```

## GUIの操作方法

1.  **部署名の選択:** 画面上部の「部署名フィルター」で、データを絞り込みたい部署を選択します（複数選択可）。
2.  **送信者アカウントの確認:** 部署を選択すると、デフォルトの送信者アカウントが自動で設定されます。必要に応じて、ドロップダウンリストから別のアカウントに変更することも可能です。
3.  **データ取得:** 「Notionからデータを取得」ボタンを押すと、Notionから発注対象のデータを取得し、左側のリストに仕入先名が表示されます。
4.  **仕入先の選択:** 左側のリストから仕入先を選択すると、右側のテーブルに発注内容が表示され、PDFの作成がバックグラウンドで開始されます。
5.  **プレビューと送信:** PDFの作成が完了すると、画面下部に宛先や担当者、添付ファイル名が表示されます。内容を確認し、問題がなければ「メール送信」ボタンを押してください。
6.  **Notionの更新:** メール送信後、Notionの対象ページの「発注日」を更新するか確認ダイアログが表示されます。「はい」を選択すると、発注日が今日の日付で記録されます。

## 注意事項
- 本アプリケーションは`win32com`ライブラリを使用しているため、**Windows環境でのみ動作**します。
- `email_accounts.json` にメールアカウントのパスワードを直接記述します。ファイルの取り扱いには十分ご注意ください。