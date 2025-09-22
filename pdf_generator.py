import os
import re
from datetime import datetime
import win32com.client
import config
import traceback

def create_order_pdf(supplier_name, items, sales_contact, sender_info, selected_department=None):
    """
    Excelテンプレートから注文書PDFを作成する (win32comのみ使用)。
    この関数はwin32comを使用するため、Windows環境とExcelのインストールが必要です。
    また、スレッドから呼び出す場合は、呼び出し元でpythoncom.CoInitialize()と
    CoUninitialize()を処理する必要があります。
    """
    excel = None
    workbook = None
    
    if not config.EXCEL_TEMPLATE_PATH or not os.path.exists(config.EXCEL_TEMPLATE_PATH):
        return None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # win32comでテンプレートを開き、データを書き込む
        # win32comは絶対パスを要求するため、os.path.abspathで変換
        workbook = excel.Workbooks.Open(os.path.abspath(config.EXCEL_TEMPLATE_PATH))
        ws = workbook.ActiveSheet

        ws.Range("A5").Value = f"{supplier_name} 御中"
        ws.Range("A7").Value = f"{sales_contact} 様"
        ws.Range("D8").Value = f"担当：{sender_info['name']}"
        ws.Range("D14").Value = sender_info["email"]

        for i, item in enumerate(items):
            ws.Range(f"A{16+i}").Value = item["db_part_number"]
            ws.Range(f"B{16+i}").Value = item["maker_name"]
            ws.Range(f"C{16+i}").Value = item["quantity"]
        
        # PDFのファイルパスを生成
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        safe_supplier_name = re.sub(r'[\\/:*?"<>|]', '_', supplier_name)
        pdf_filename = f"{timestamp}_{safe_supplier_name}_注文書.pdf"
        
        # 保存ディレクトリを決定
        target_save_dir = config.PDF_SAVE_DIR
        if selected_department:
            department_dir = os.path.join(config.PDF_SAVE_DIR, selected_department)
            if os.path.exists(department_dir) and os.path.isdir(department_dir):
                target_save_dir = department_dir
            else:
                # 部署フォルダが存在しない場合は、デフォルトの保存先に保存
                print(f"警告: 部署フォルダ '{department_dir}' が見つからないため、デフォルトの保存先に保存します。")

        if not os.path.exists(target_save_dir):
            os.makedirs(target_save_dir)
        pdf_path = os.path.join(target_save_dir, pdf_filename)
        
        # 一時ファイルなしで直接PDFに変換
        ws.ExportAsFixedFormat(0, pdf_path)
        
        
        return pdf_path
        
    except Exception as e:
        print(f"PDF作成エラー ({supplier_name}): {e}")
        traceback.print_exc()
        return None
        
    finally:
        # クリーンアップ処理
        if workbook:
            workbook.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        # 念のためCOMオブジェクトの参照を解放