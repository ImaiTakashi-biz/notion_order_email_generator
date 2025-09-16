# Notion 注文書メール自動作成アプリ

Notionで管理している注文データを元に、Excelテンプレートを使用して注文書のPDFを自動作成し、担当者へメールで送信するデスクトップアプリケーションです。

## 主な機能

-   Notionデータベースから「要発注」ステータスのデータを取得
-   仕入先ごとに注文内容をグループ化
-   指定したExcelテンプレートを元に注文書PDFを自動作成
-   仕入先ごとにPDFを添付したメールを作成し、送信
-   処理済みのNotionデータに「発注日」を自動で記録
-   複数の送信元メールアカウントの切り替えに対応

## 動作要件

-   **OS:** Windows
-   **ソフトウェア:** Microsoft Excel
    -   本アプリケーションは内部でExcelを操作してPDFを作成するため、Excelのインストールが必須です。

## セットアップ手順

### 1. 依存ライブラリのインストール

以下のコマンドを実行して、必要なPythonライブラリをインストールします。

```bash
pip install -r requirements.txt
```

### 2. 環境変数ファイル (.env) の作成

プロジェクトのルートディレクトリに `.env` という名前のファイルを作成し、以下の内容を記述します。
これはAPIキーやパスワードなど、公開すべきでない情報を設定するためのファイルです。

```dotenv
# Notion APIのインテグレーションシークレット
NOTION_API_TOKEN="secret_..."

# 注文管理DBが含まれるNotionページのID
NOTION_DATABASE_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# --- 送信元メールアカウント情報 ---
# Office365の場合
SMTP_SERVER="smtp.office365.com"
SMTP_PORT="587"

# 送信者アカウント1 (xxの部分は任意の名称)
EMAIL_SENDER_TARO="taro.suzuki@example.com"
EMAIL_PASSWORD_TARO="password123"

# 送信者アカウント2 (複数設定可能)
EMAIL_SENDER_HANAKO="hanako.yamada@example.com"
EMAIL_PASSWORD_HANAKO="password456"
```

-   `NOTION_API_TOKEN`: Notionインテグレーションのシークレットトークン。
-   `NOTION_DATABASE_ID`: データを取得したいデータベースが含まれている親ページのID。
-   `EMAIL_SENDER_xx`: 送信元として使用するメールアドレス。`xx`の部分はアプリ内のアカウント選択肢に表示される名前になります。
-   `EMAIL_PASSWORD_xx`: 上記メールアドレスのパスワード。`xx`の部分は`EMAIL_SENDER_xx`と一致させてください。

### 3. 設定ファイル (config.json)

アプリケーションを初めて起動すると、`.env` と同じ階層に `config.json` ファイルが自動的に作成されます。
このファイルには、Excelテンプレートのパスなどが保存されます。

```json
{
    "EXCEL_TEMPLATE_PATH": "C:\path\to\your\template.xlsx",
    "PDF_SAVE_DIR": "C:\path\to\save\pdfs"
}
```

これらの値は、アプリケーションのGUI画面にある「設定」ボタンから変更・保存することが可能です。

## 実行方法

以下のコマンドでアプリケーションを起動します。

```bash
python main.py
```

## GUIの操作方法

1.  **送信者アカウントの選択:** 画面右上で、メールの送信元として使用するアカウントを選択します。
2.  **データ取得:** 「1. Notionからデータを取得」ボタンを押すと、Notionから発注対象のデータを取得し、左側のリストに仕入先名が表示されます。
3.  **仕入先の選択:** 左側のリストから仕入先を選択すると、右側のテーブルに発注内容が表示され、PDFの作成がバックグラウンドで開始されます。
4.  **プレビューと送信:** PDFの作成が完了すると、画面下部に宛先や担当者、添付ファイル名が表示されます。内容を確認し、問題がなければ「メール送信」ボタンを押してください。
5.  **Notionの更新:** メール送信後、Notionの対象ページの「発注日」を更新するか確認ダイアログが表示されます。「はい」を選択すると、発注日が今日の日付で記録されます。
6.  **設定の変更:** 「設定」ボタンから、ExcelテンプレートのパスとPDFの保存先フォルダをいつでも変更できます。
