import tempfile
import unittest
from unittest.mock import MagicMock, patch

import pdf_generator
import config


class TestPdfGenerator(unittest.TestCase):

    @patch('pdf_generator.SimpleDocTemplate')
    def test_create_order_pdf_success(self, mock_doc_template):
        mock_doc = MagicMock()
        mock_doc_template.return_value = mock_doc

        supplier_name = "Sample Supplier"
        items = [
            {"db_part_number": "PN-001", "maker_name": "MakerA", "quantity": 10, "remarks": ""},
            {"db_part_number": "PN-002", "maker_name": "MakerB", "quantity": 5, "remarks": ""},
        ]
        sales_contact = "John Doe"
        sender_info = {"name": "Alice Sender", "email": "test@example.com"}
        selected_department = "R&D"

        with tempfile.TemporaryDirectory() as tmpdir:
            result_path = pdf_generator.create_order_pdf(
                supplier_name,
                items,
                sales_contact,
                sender_info,
                selected_department,
                save_dir=tmpdir,
            )

        self.assertIsNotNone(result_path)
        self.assertTrue(result_path.startswith(tmpdir))
        self.assertTrue(result_path.endswith("_注文書.pdf"))
        self.assertIn("Sample Supplier", result_path)
        mock_doc_template.assert_called_once()
        mock_doc.build.assert_called_once()

    def test_returns_none_when_save_dir_unavailable(self):
        with patch.object(config, "PDF_SAVE_DIR", "", create=True):
            result = pdf_generator.create_order_pdf(
                "Fallback Supplier",
                [{"db_part_number": "A-1", "maker_name": "Maker", "quantity": 1, "remarks": ""}],
                "Fallback Contact",
                {"name": "Sender Name", "email": "sender@example.com"},
                selected_department=None,
            )
        self.assertIsNone(result)


if __name__ == '__main__':
    unittest.main()
