import pytest
import json
import os
from unittest.mock import Mock, patch, mock_open
import pandas as pd

from modules import BC_webservice, get_app_version, export_excel, export_csv, send_email


class TestBCWebservice:
    def test_init(self):
        ws = BC_webservice("tenant", "client", "secret", "scope", "url")
        assert ws.tenant_id == "tenant"
        assert ws.client_id == "client"
        assert ws.client_secret == "secret"
        assert ws.scope == "scope"
        assert ws.token_url == "url"

    @patch('modules.requests.post')
    def test_get_token_success(self, mock_post):
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"access_token": "token123"}
        mock_post.return_value = mock_response

        ws = BC_webservice("tenant", "client", "secret", "scope", "url")
        token = ws._BC_webservice__get_token()
        assert token == "token123"
        mock_post.assert_called_once()

    @patch('modules.requests.post')
    def test_get_token_failure(self, mock_post):
        mock_response = Mock()
        mock_response.status_code = 400
        mock_post.return_value = mock_response

        ws = BC_webservice("tenant", "client", "secret", "scope", "url")
        with pytest.raises(SystemExit):
            ws._BC_webservice__get_token()

    @patch('modules.requests.get')
    @patch.object(BC_webservice, '_BC_webservice__get_token')
    def test_get_data_success(self, mock_get_token, mock_get):
        mock_get_token.return_value = "token123"
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"value": [{"Name": "Test", "Version": "1.0", "Is_Installed": True}]}
        mock_get.return_value = mock_response

        ws = BC_webservice("tenant", "client", "secret", "scope", "url")
        result = ws.get_data("env", "company", "ws", "?filter=test")
        assert result == {"value": [{"Name": "Test", "Version": "1.0", "Is_Installed": True}]}
        mock_get.assert_called_once()

    @patch('modules.requests.get')
    @patch.object(BC_webservice, '_BC_webservice__get_token')
    def test_get_data_500_error(self, mock_get_token, mock_get):
        mock_get_token.return_value = "token123"
        mock_response = Mock()
        mock_response.status_code = 500
        mock_response.text = "Server error"
        mock_get.return_value = mock_response

        ws = BC_webservice("tenant", "client", "secret", "scope", "url")
        with patch('builtins.print'):
            result = ws.get_data("env", "company", "ws", "?filter=test")
        assert result == {"value": [{"Name": "N/A", "Version": "N/A", "Is_Installed": False}]}


class TestGetAppVersion:
    @patch('modules.webdriver.Chrome')
    @patch('modules.Service')
    @patch('modules.ChromeDriverManager')
    def test_get_app_version_success(self, mock_cdm, mock_service, mock_chrome):
        mock_driver = Mock()
        mock_element = Mock()
        mock_element.text = "1.2.3"
        mock_driver.find_element.return_value = mock_element
        mock_chrome.return_value = mock_driver

        version = get_app_version("http://test.com")
        assert version == "1.2.3"
        mock_driver.get.assert_called_with("http://test.com")
        mock_driver.quit.assert_called_once()

    @patch('modules.webdriver.Chrome')
    @patch('modules.Service')
    @patch('modules.ChromeDriverManager')
    def test_get_app_version_error(self, mock_cdm, mock_service, mock_chrome):
        mock_driver = Mock()
        mock_driver.find_element.side_effect = Exception("Error")
        mock_chrome.return_value = mock_driver

        version = get_app_version("http://test.com")
        assert version == "Error"


class TestExportFunctions:
    def test_export_excel(self, tmp_path):
        data = [
            {"Country": "US", "Extension": "Test", "Environment": "prod", "Version": "1.0"}
        ]
        filename = str(tmp_path / "test.xlsx")
        result = export_excel(data, filename)
        assert os.path.exists(filename)
        # Check if it's a valid excel
        df = pd.read_excel(filename)
        assert len(df) > 0

    def test_export_csv(self, tmp_path):
        data = [
            {"Country": "US", "Extension": "Test", "Environment": "prod", "Version": "1.0"}
        ]
        filename = str(tmp_path / "test.csv")
        export_csv(data, filename)
        assert os.path.exists(filename)
        with open(filename, 'r') as f:
            content = f.read()
            assert "Country" in content


class TestSendEmail:
    @patch('modules.smtplib.SMTP')
    def test_send_email_success(self, mock_smtp):
        mock_server = Mock()
        mock_smtp.return_value.__enter__.return_value = mock_server

        send_email(
            smtp_server="smtp.test.com",
            smtp_port=587,
            user="user@test.com",
            password="pass",
            email_from="from@test.com",
            email_to="to@test.com",
            subject="Test",
            content="Content",
            file="test.xlsx"
        )

        mock_smtp.assert_called_with("smtp.test.com", 587)
        mock_server.starttls.assert_called_once()
        mock_server.login.assert_called_once_with("user@test.com", "pass")
        mock_server.send_message.assert_called_once()

    @patch('modules.smtplib.SMTP')
    def test_send_email_file_not_found(self, mock_smtp):
        with pytest.raises(SystemExit):
            send_email(
                smtp_server="smtp.test.com",
                smtp_port=587,
                user="user@test.com",
                password="pass",
                email_from="from@test.com",
                email_to="to@test.com",
                subject="Test",
                content="Content",
                file="nonexistent.xlsx"
            )