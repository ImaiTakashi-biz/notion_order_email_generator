import os
import re
from datetime import datetime
import win32com.client as win32
import pythoncom

import config

def create_order_pdf(supplier_name, items, sales_contact, sender_info, selected_department=None, save_dir=None):
    """
    Excelテンプレートから注文書PDFを作成する内部関数。
    win32comを使用するため、Windows環境とExcelのインストールが必要。
    """
    excel = None
    workbook = None
    
    if not config.EXCEL_TEMPLATE_PATH or not os.path.exists(config.EXCEL_TEMPLATE_PATH):
        return None

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # win32comでテンプレートを開き、データを書き込む
        # win32comは絶対パスを要求するため、os.path.abspathで変換
        workbook = excel.Workbooks.Open(os.path.abspath(config.EXCEL_TEMPLATE_PATH))
        ws = workbook.ActiveSheet

        cells = config.AppConstants.EXCEL_CELLS
        # Formula Injection対策: 値の先頭にシングルクォートを付与して文字列として扱う
        ws.Range(cells['SUPPLIER_NAME']).Value = f"'{supplier_name} 御中"
        ws.Range(cells['SALES_CONTACT']).Value = f"'{sales_contact} 様"
        
        # 部署名を担当名前に追加
        if selected_department:
            ws.Range(cells['SENDER_NAME']).Value = f"'担当：{selected_department} {sender_info['name']}"
        else:
            ws.Range(cells['SENDER_NAME']).Value = f"'担当：{sender_info['name']}"
            
        ws.Range(cells['SENDER_EMAIL']).Value = f"'{sender_info["email"]}"

        for i, item in enumerate(items):
            row = cells['ITEM_START_ROW'] + i
            # Formula Injection対策: 値の先頭にシングルクォートを付与
            ws.Range(f"A{row}").Value = f"'{item.get("db_part_number", "")}"
            ws.Range(f"B{row}").Value = f"'{item.get("maker_name", "")}"
            ws.Range(f"C{row}").Value = f"'{item.get("quantity", 0)}"
        
        # PDFのファイルパスを生成
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        safe_supplier_name = re.sub(r'[\\/:*?"<>|]', '_', supplier_name)
        pdf_filename = f"{timestamp}_{safe_supplier_name}_注文書.pdf"
        
        # 保存ディレクトリを決定 (引数が優先)
        base_save_dir = save_dir if save_dir is not None else config.PDF_SAVE_DIR

        # 部署ごとのサブフォルダを考慮
        target_save_dir = base_save_dir
        if selected_department:
            department_dir = os.path.join(base_save_dir, selected_department)
            # 部署フォルダが存在する場合のみ、そちらを優先
            if os.path.exists(department_dir) and os.path.isdir(department_dir):
                target_save_dir = department_dir
            # 存在しない場合は、ベースの保存先に保存（自動作成はしない）
            elif not save_dir: # 通常保存の場合のみ作成
                 try:
                    os.makedirs(department_dir, exist_ok=True)
                    target_save_dir = department_dir
                 except OSError as e:
                    print(f"部署フォルダの作成に失敗: {e}")

        if not os.path.exists(target_save_dir):
            os.makedirs(target_save_dir)
        pdf_path = os.path.join(target_save_dir, pdf_filename)
        
        # 一時ファイルなしで直接PDFに変換
        ws.ExportAsFixedFormat(0, pdf_path)
        
        print(f"✅ PDFを作成しました: {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"❌ PDF作成中にエラーが発生しました: {e}")
        return None
    finally:
        if 'excel' in locals() and excel is not None:
            excel.Quit()

def generate_order_pdf_flow(supplier_name, items, sender_info, selected_department=None, save_dir=None):
    """
    UIからの情報を受け取り、PDF作成のフロー全体を管理する
    """
    try:
        if not items:
            error_message = "対象アイテムが見つかりません。"
            print(f"❌ PDF作成エラー: {error_message}")
            return None, None, error_message

        # create_order_pdfを呼び出す
        pdf_path = create_order_pdf(
            supplier_name,
            items,
            items[0]["sales_contact"], # 代表の連絡先を取得
            sender_info,
            selected_department=selected_department,
            save_dir=save_dir # 保存先を渡す
        )
        
        if not pdf_path:
            # create_order_pdf内でエラーが発生した場合
            return None, None, "ExcelでのPDF作成中にエラーが発生しました。詳細はコンソールログを確認してください。"

        # GUIのプレビュー更新に必要な情報を返す (成功)
        return pdf_path, items[0], None

    except Exception as e:
        error_message = f"予期せぬエラーが発生しました: {e}"
        print(f"❌ PDF生成フロー中に{error_message}")
        return None, None, error_message


