
import unittest
from unittest.mock import patch, MagicMock
import smtplib

# テスト対象のモジュール
import email_service
import config

class TestEmailService(unittest.TestCase):

    @patch('email_service.keyring.get_password')
    @patch('smtplib.SMTP')
    @patch('builtins.open', new_callable=unittest.mock.mock_open, read_data=b'dummy data')
    def test_prepare_and_send_order_email_success(self, mock_open, mock_smtp, mock_get_password):
        """
        メール送信が成功するケースをテストする
        """
        # --- 準備 (Arrange) ---
        # モックの設定
        mock_get_password.return_value = 'dummy_password'
        
        # smtplib.SMTPのインスタンスとメソッドをモック化
        mock_smtp_instance = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_smtp_instance

        # テストデータ
        account_key = 'test_account'
        sender_creds = {'sender': 'test@example.com', 'display_name': 'テスト担当'}
        items = [{
            'supplier_name': 'テスト仕入先',
            'sales_contact': 'テスト担当者',
            'email': 'recipient@example.com',
            'email_cc': 'cc@example.com'
        }]
        pdf_path = 'dummy.pdf'
        selected_department = 'テスト部署'

        # --- 実行 (Act) ---
        success = email_service.prepare_and_send_order_email(
            account_key, sender_creds, items, pdf_path, selected_department
        )

        # --- 検証 (Assert) ---
        self.assertTrue(success)
        # PDFファイルを開こうとしたか
        mock_open.assert_called_once_with('dummy.pdf', 'rb')
        # keyring.get_passwordが正しい引数で呼ばれたか
        mock_get_password.assert_called_once_with(email_service.SERVICE_NAME, 'test@example.com')
        # SMTPサーバーに接続しようとしたか
        mock_smtp.assert_called_once_with(config.SMTP_SERVER, config.SMTP_PORT)
        # サーバーのメソッドが呼ばれたか
        mock_smtp_instance.starttls.assert_called_once()
        mock_smtp_instance.login.assert_called_once_with('test@example.com', 'dummy_password')
        
        # sendmailが呼ばれ、その第一引数（送信元）を検証
        call_args, _ = mock_smtp_instance.sendmail.call_args
        self.assertEqual(call_args[0], 'test@example.com')
        # sendmailの第二引数（送信先リスト）を検証
        self.assertEqual(call_args[1], ['recipient@example.com', 'cc@example.com'])


    @patch('email_service.keyring.get_password')
    def test_send_fail_if_password_not_found(self, mock_get_password):
        """
        keyringにパスワードが保存されていない場合に失敗することをテストする
        """
        # --- 準備 (Arrange) ---
        mock_get_password.return_value = None # パスワードが見つからない状況を再現

        # テストデータ
        account_key = 'test_account'
        sender_creds = {'sender': 'test@example.com', 'display_name': 'テスト担当'}
        items = [{'email': 'recipient@example.com'}]
        pdf_path = 'dummy.pdf'

        # --- 実行 (Act) ---
        success = email_service.prepare_and_send_order_email(
            account_key, sender_creds, items, pdf_path
        )

        # --- 検証 (Assert) ---
        self.assertFalse(success)
        mock_get_password.assert_called_once_with(email_service.SERVICE_NAME, 'test@example.com')

    @patch('email_service.keyring.get_password')
    @patch('smtplib.SMTP')
    def test_send_fail_on_smtp_auth_error(self, mock_smtp, mock_get_password):
        """
        SMTP認証エラーが発生した場合に失敗することをテストする
        """
        # --- 準備 (Arrange) ---
        mock_get_password.return_value = 'wrong_password'
        
        # loginメソッドがSMTPAuthenticationErrorを発生させるように設定
        mock_smtp_instance = MagicMock()
        mock_smtp_instance.login.side_effect = smtplib.SMTPAuthenticationError(535, b'Authentication failed')
        mock_smtp.return_value.__enter__.return_value = mock_smtp_instance

        # テストデータ
        account_key = 'test_account'
        sender_creds = {'sender': 'test@example.com', 'display_name': 'テスト担当'}
        items = [{'email': 'recipient@example.com', 'supplier_name': 'テスト仕入先'}]
        pdf_path = 'dummy.pdf'

        # --- 実行 (Act) ---
        success = email_service.prepare_and_send_order_email(
            account_key, sender_creds, items, pdf_path
        )

        # --- 検証 (Assert) ---
        self.assertFalse(success)

if __name__ == '__main__':
    unittest.main()
