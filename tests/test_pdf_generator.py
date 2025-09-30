
import unittest
from unittest.mock import patch, MagicMock, ANY
import os
from datetime import datetime

# テスト対象のモジュール
import pdf_generator
import config

class TestPdfGenerator(unittest.TestCase):

    @patch('pdf_generator.os.path.exists')
    @patch('pdf_generator.win32.Dispatch')
    def test_create_order_pdf_success(self, mock_dispatch, mock_exists):
        """
        PDF作成が正常に成功するケースをテストする
        """
        # --- 準備 (Arrange) ---
        # os.path.existsが常にTrueを返すように設定（テンプレートファイルが存在する想定）
        mock_exists.return_value = True

        # win32comのExcelアプリケーション全体をモック化
        mock_excel_app = MagicMock()
        mock_workbook = MagicMock()
        mock_worksheet = MagicMock()

        # Dispatchがモックアプリを返すように設定
        mock_dispatch.return_value = mock_excel_app
        # Openがモックワークブックを返すように設定
        mock_excel_app.Workbooks.Open.return_value = mock_workbook
        # ActiveSheetがモックワークシートを返すように設定
        mock_workbook.ActiveSheet = mock_worksheet

        # テストデータ
        supplier_name = "テスト株式会社"
        items = [
            {"db_part_number": "PN-001", "maker_name": "メーカーA", "quantity": 10},
            {"db_part_number": "PN-002", "maker_name": "メーカーB", "quantity": 5},
        ]
        sales_contact = "営業担当者様"
        sender_info = {"name": "テスト太郎", "email": "test@example.com"}
        selected_department = "開発部"

        # --- 実行 (Act) ---
        result_path = pdf_generator.create_order_pdf(
            supplier_name, items, sales_contact, sender_info, selected_department
        )

        # --- 検証 (Assert) ---
        # 1. 正しいPDFパスが返されたか
        self.assertIsNotNone(result_path)
        self.assertIn(supplier_name, result_path)
        self.assertTrue(result_path.endswith("_注文書.pdf"))

        # 2. Excelオブジェクトが正しく呼び出されたか
        mock_dispatch.assert_called_once_with("Excel.Application")
        mock_excel_app.Workbooks.Open.assert_called_once_with(os.path.abspath(config.EXCEL_TEMPLATE_PATH))

        # 3. セルに正しい値が書き込まれたか (一部を抜粋してチェック)
        cells = config.AppConstants.EXCEL_CELLS
        mock_worksheet.Range(cells['SUPPLIER_NAME']).Value = f"'{supplier_name} 御中"
        mock_worksheet.Range(cells['SENDER_NAME']).Value = f"'担当：{selected_department} {sender_info['name']}"
        # 1行目のアイテム
        mock_worksheet.Range("A16").Value = f"'{items[0]['db_part_number']}"
        mock_worksheet.Range("C16").Value = f"'{items[0]['quantity']}"
        # 2行目のアイテム
        mock_worksheet.Range("A17").Value = f"'{items[1]['db_part_number']}"

        # 4. PDFとしてエクスポートされたか
        mock_worksheet.ExportAsFixedFormat.assert_called_once_with(0, result_path)

        # 5. Excelが正しく終了されたか
        mock_workbook.Close.assert_called_once_with(SaveChanges=False)
        mock_excel_app.Quit.assert_called_once()

    @patch('pdf_generator.os.path.exists')
    def test_fail_if_template_not_found(self, mock_exists):
        """
        Excelテンプレートが見つからない場合にNoneを返すことをテストする
        """
        # --- 準備 (Arrange) ---
        mock_exists.return_value = False # テンプレートが存在しない状況を再現

        # --- 実行 (Act) ---
        result = pdf_generator.create_order_pdf("s", [{"sales_contact":"c"}], "c", {"name":"n", "email":"e"})

        # --- 検証 (Assert) ---
        self.assertIsNone(result)

if __name__ == '__main__':
    unittest.main()
