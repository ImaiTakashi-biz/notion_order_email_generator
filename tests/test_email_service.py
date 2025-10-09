import unittest
from unittest.mock import patch, MagicMock
import smtplib

# 対象モジュール
import email_service
import config

class TestEmailService(unittest.TestCase):

    @patch('email_service.os.path.exists')
    @patch('email_service.keyring.get_password')
    @patch('smtplib.SMTP')
    @patch('builtins.open', new_callable=unittest.mock.mock_open, read_data=b'dummy data')
    def test_prepare_and_send_order_email_success(self, mock_open, mock_smtp, mock_get_password, mock_exists):
        """正常系: メール送信が成功する"""
        mock_get_password.return_value = 'dummy_password'
        mock_exists.return_value = True

        mock_smtp_instance = MagicMock()
        mock_smtp.return_value.__enter__.return_value = mock_smtp_instance

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

        success, message = email_service.prepare_and_send_order_email(
            account_key, sender_creds, items, pdf_path, selected_department
        )

        self.assertTrue(success)
        self.assertIsNone(message)
        mock_open.assert_any_call('dummy.pdf', 'rb')
        mock_get_password.assert_called_once_with(email_service.SERVICE_NAME, 'test@example.com')
        mock_smtp.assert_called_once_with(config.SMTP_SERVER, config.SMTP_PORT)
        mock_smtp_instance.starttls.assert_called_once()
        mock_smtp_instance.login.assert_called_once_with('test@example.com', 'dummy_password')
        call_args, _ = mock_smtp_instance.sendmail.call_args
        self.assertEqual(call_args[0], 'test@example.com')
        self.assertEqual(call_args[1], ['recipient@example.com', 'cc@example.com'])

    @patch('email_service.keyring.get_password')
    def test_send_fail_if_password_not_found(self, mock_get_password):
        """Keyring にパスワードが無い場合は失敗"""
        mock_get_password.return_value = None

        account_key = 'test_account'
        sender_creds = {'sender': 'test@example.com', 'display_name': 'テスト担当'}
        items = [{'email': 'recipient@example.com'}]
        pdf_path = 'dummy.pdf'

        success, message = email_service.prepare_and_send_order_email(
            account_key, sender_creds, items, pdf_path
        )

        self.assertFalse(success)
        self.assertIn('パスワード', message)
        mock_get_password.assert_called_once_with(email_service.SERVICE_NAME, 'test@example.com')

    @patch('email_service.os.path.exists')
    @patch('email_service.keyring.get_password')
    @patch('smtplib.SMTP')
    @patch('builtins.open', new_callable=unittest.mock.mock_open, read_data=b'dummy data')
    def test_send_fail_on_smtp_auth_error(self, mock_open, mock_smtp, mock_get_password, mock_exists):
        """SMTP認証エラー時は失敗"""
        mock_get_password.return_value = 'wrong_password'
        mock_exists.return_value = True

        mock_smtp_instance = MagicMock()
        mock_smtp_instance.login.side_effect = smtplib.SMTPAuthenticationError(535, b'Authentication failed')
        mock_smtp.return_value.__enter__.return_value = mock_smtp_instance

        account_key = 'test_account'
        sender_creds = {'sender': 'test@example.com', 'display_name': 'テスト担当'}
        items = [{'email': 'recipient@example.com', 'supplier_name': 'テスト仕入先'}]
        pdf_path = 'dummy.pdf'

        success, message = email_service.prepare_and_send_order_email(
            account_key, sender_creds, items, pdf_path
        )

        self.assertFalse(success)
        self.assertIn('SMTP認証', message)

if __name__ == '__main__':
    unittest.main()
