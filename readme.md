# BC Extensions Version Checker

A Python tool to audit and compare extension versions in Microsoft Business Central across multiple countries and environments. It retrieves version information from AppSource marketplace and installed extensions in staging/production environments, then exports the results to an Excel file for easy analysis.

## Features

- **Multi-Country Support**: Check extensions across different countries (e.g., Poland, Romania)
- **Environment Comparison**: Compare versions between AppSource, staging, and production environments
- **Automated Data Collection**: Uses web scraping for AppSource and REST API calls for Business Central
- **Excel Export**: Generates a pivot table Excel report showing version discrepancies
- **OAuth2 Authentication**: Secure API access to Business Central webservices

## Requirements

- Python 3.8+
- Microsoft Business Central API access with proper credentials

## Dependencies

1. Install the required packages using pip:

```bash
pip install -r requirements.txt
```

Required packages:
- requests
- playwright
- pandas
- openpyxl

2. Install Chromium headless

```bash
playwright install chromium
```

3. Install Linux dependencies (Debian/Ubuntu)

```bash
sudo apt install -y \
  libnss3 \
  libatk1.0-0 \
  libatk-bridge2.0-0 \
  libcups2 \
  libdrm2 \
  libxkbcommon0 \
  libxcomposite1 \
  libxdamage1 \
  libxrandr2 \
  libgbm1 \
  libasound2 \
  libpangocairo-1.0-0 \
  libpango-1.0-0 \
  libgtk-3-0
```

## Configuration

1. **Credentials Setup**: Update the credentials in `.env`:
   - `TENANT_ID`: Your Azure AD tenant ID
   - `PROD_CLIENT_ID` and `PROD_CLIENT_SECRET`: Production environment app registration credentials
   - `STAGE_CLIENT_ID` and `STAGE_CLIENT_SECRET`: Staging environment app registration credentials

2. **Country and Extension Configuration**: Modify the configuration in `data.json` to include:
   - Country codes
   - Environment URLs and company names
   - Extension names and AppSource URLs (if available)

3. **Recipients Configuration**: Update `recipients.json` with email addresses for report notifications.

## Email Configuration

The tool can automatically send the generated Excel report via email after completion. To enable this feature:

1. **SMTP Settings**: Add the following to your `.env` file:
   - `SMTP_SERVER`: Your SMTP server address (e.g., smtp.gmail.com)
   - `SMTP_PORT`: SMTP port (e.g., 587 for TLS)
   - `EMAIL`: Your sender email address
   - `PASSW`: Your email password or app-specific password

2. **Recipients**: Ensure `recipients.json` contains the list of email recipients.

3. **Email Content**: The email will include a predefined subject and message body with the Excel file attached.

Note: Use app-specific passwords for Gmail or similar services for security.

## Usage

1. Activate your virtual environment (Debian/Ubuntu):
   ```bash   
   source .venv/bin/activate
   ```

2. Run the main script:
   ```bash
   python extensions.py
   ```

3. The script will:
   - Fetch versions from AppSource for configured extensions
   - Query Business Central APIs for installed extension versions
   - Export results to `extensions_versions.xlsx`
   - Send email with results to recipients from `recipients.json`

## Output

The generated Excel file contains a pivot table with:
- **Rows**: Country and Extension name
- **Columns**: Environment (AppSource, stage, prod)
- **Values**: Version numbers

This allows easy identification of version mismatches between environments.

## Security Notes

- Store credentials securely and never commit them to version control
- Consider using Azure Key Vault or environment variables for production deployments
- The script uses headless Chrome for web scraping to avoid UI interference

## Troubleshooting

- **API Authentication Errors**: Verify credentials and Azure AD app registrations
- **Web Scraping Failures**: AppSource page structure may change; update XPath selectors in `modules.py`

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

GNU General Public License (GPL)