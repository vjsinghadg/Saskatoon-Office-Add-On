# ADGSentinel - Outlook Web Add-in (Office.js)

A modern Outlook Web Add-in built with Office.js that enables users to report phishing, spam, and legitimate emails directly from Outlook Web.

## Features

- **Report Phishing**: Report suspicious emails as phishing attempts
- **Report Spam**: Report emails as junk/spam
- **Report Legitimate**: Provide feedback on legitimate emails
- **Email Details Extraction**: Automatically extract and include email metadata, headers, URLs, and attachments info
- **GoPhish Integration**: Detection and reporting of simulated phishing campaigns
- **User Information Logging**: Capture reporter details for audit trail
- **URL Extraction**: Automatically detect and list URLs in emails with domain extraction

## Project Structure

```
OutlookWebAdd/
├── manifest.xml              # Office Add-in manifest configuration
├── function-file.html        # Entry point for command handlers
├── function-file.js          # Main logic for report handling
├── server.js                 # Express server for local development
├── package.json              # Node.js dependencies
├── certs/                    # SSL certificates for HTTPS (create manually)
│   ├── key.pem
│   └── cert.pem
├── public/                   # Static assets
│   ├── assets/               # Icons and images
│   │   ├── icon-16.png
│   │   ├── icon-32.png
│   │   ├── icon-80.png
│   │   ├── phishing-icon-16.png
│   │   ├── phishing-icon-32.png
│   │   ├── phishing-icon-80.png
│   │   ├── spam-icon-16.png
│   │   ├── spam-icon-32.png
│   │   ├── spam-icon-80.png
│   │   ├── legitimate-icon-16.png
│   │   ├── legitimate-icon-32.png
│   │   └── legitimate-icon-80.png
│   └── taskpane.html         # Optional: Taskpane UI
└── README.md                 # This file
```

## Prerequisites

- Node.js (v12 or later) and npm
- Outlook Web Access or Outlook Desktop (with Office.js support)
- HTTPS server (self-signed cert for development)
- Valid Microsoft 365 account

## Setup Instructions

### 1. Install Dependencies

```bash
npm install
```

### 2. Generate SSL Certificates (for local development)

```bash
mkdir -p certs
cd certs

# Using OpenSSL (macOS/Linux)
openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes

# Provide certificate details when prompted
```

### 3. Configure Settings

Edit `manifest.xml` and update:

- `<Id>`: Generate unique ID (UUID format)
- `<ProviderName>`: Your organization name
- `<DisplayName>`: Add-in display name

Edit `function-file.js` and update the `CONFIG` object:

```javascript
const CONFIG = {
  infosecEmail: "your-infosec@company.com", // InfoSec team email
  spamReportEmail: "spam-reports@company.com", // Spam report destination
  supportEmail: "addin-support@company.com", // Support email
  gophishUrl: "https://gophish.your-domain.com", // GoPhish URL (optional)
  gophishListenerPort: 3333, // GoPhish listener port
  gophishCustomHeader: "X-SENTINEL-AJSMN", // Custom header for GoPhish
};
```

### 4. Start Development Server

```bash
npm start
```

The server will start at `https://localhost:3000`

### 5. Create Static Assets (Optional)

Create placeholder icons in `public/assets/` directory (16x16, 32x32, 80x80 PNG):

- icon-\*.png (main add-in icon)
- phishing-icon-\*.png (phishing report icon)
- spam-icon-\*.png (spam report icon)
- legitimate-icon-\*.png (legitimate report icon)

### 6. Upload Add-in to Outlook

#### For Outlook Web:

1. Go to **Settings** (gear icon) > **View all Outlook settings**
2. Navigate to **Add-ins** > **Get Add-ins**
3. Select **My Add-ins**
4. Click **Upload My Add-in**
5. Choose **Upload from URL** and enter:
   ```
   https://localhost:3000/manifest.xml
   ```
   OR
   Upload the `manifest.xml` file directly

#### For Outlook Desktop (Mac/Windows):

1. Go to **File** > **Info** > **Manage Add-ins**
2. Click **+ Add Add-in** > **My Add-ins** > **Upload My Add-in**
3. Select or provide URL to manifest.xml

## Usage

### Reporting an Email

1. Open an email in Outlook Web or Outlook Desktop
2. Click the **Home** or **Message** tab
3. In the **InfoSec** group, click **ADGSentinel Report** dropdown
4. Select one of:
   - **Report Phishing**: Report as phishing attempt
   - **Report Spam**: Report as junk/spam
   - **Report Legitimate**: Report as safe/legitimate
5. A confirmation notification will appear

### Report Email Details

When a report is generated, the following information is included:

#### Email Information:

- Subject, From, To, CC, BCC
- Timestamp of report
- Reporter details (name, email, timezone)
- Number of attachments
- Read/unread status

#### Security Details:

- Complete email headers
- Extracted URLs with domain names
- Attachment information
- GoPhish simulated campaign detection

#### Original Email:

- Full email body (for reference)

## Configuration Details

### Environment Variables

Set these via system environment or `.env` file:

```
INFOSEC_EMAIL=security@company.com
SPAM_REPORT_EMAIL=spam@company.com
SUPPORT_EMAIL=support@company.com
PORT=3000
```

### Manifest Configuration

**Important manifest elements:**

```xml
<!-- Function file - handles command execution -->
<bt:Url id="functionfile" DefaultValue="https://localhost:3000/function-file/function-file.html"/>

<!-- Dropdown menu with report options -->
<Control xsi:type="Menu" id="reportMenu">
  <Items>
    <Item id="reportPhishing"> ... </Item>
    <Item id="reportSpam"> ... </Item>
    <Item id="reportLegitimate"> ... </Item>
  </Items>
</Control>
```

## Limitations & Notes

### Outlook Web Add-ins Limitations:

1. **Email Deletion**: Web add-ins cannot directly delete emails. Instead:
   - Emails are marked as read
   - Email is tagged with a category
   - User can manually delete or mark as junk

2. **Email Sending**: Create report email via reply mechanism rather than direct send

3. **Header Access**: Some email headers may not be accessible in web client

4. **Attachment Processing**: Limited ability to process attachment contents; size and hash information may not be available

### Desktop Add-ins (VSTO):

For desktop Outlook, see the original implementation:
[GitHub - Saskatoon Phishing Reporter](https://github.com/adg-tech/saskatoon-phishing-reporter)

## API Endpoints

| Endpoint                            | Method | Description                 |
| ----------------------------------- | ------ | --------------------------- |
| `/manifest.xml`                     | GET    | Office Add-in manifest      |
| `/function-file/function-file.html` | GET    | Command handler entry point |
| `/scripts/function-file.js`         | GET    | Main business logic         |
| `/api/config`                       | GET    | Configuration endpoint      |
| `/health`                           | GET    | Health check                |

## Security Considerations

1. **HTTPS Only**: Always use HTTPS in production
2. **Certificate Management**: Use valid, trusted certificates in production
3. **Email Recipient Validation**: Always validate email addresses before sending
4. **Rate Limiting**: Implement rate limiting to prevent abuse
5. **Authentication**: Consider adding additional authentication/authorization
6. **Data Privacy**: Ensure compliance with GDPR, CCPA, and other regulations

## Troubleshooting

### "Add-in didn't load" Error

1. Check browser console for JavaScript errors
2. Verify manifest.xml is accessible at HTTPS URL
3. Check that all image resources are accessible
4. Ensure Office.js library is loaded correctly

### "Command not found" Error

1. Verify `FunctionName` in manifest matches function name in JavaScript
2. Ensure `function-file.html` is properly configured in manifest
3. Check that JavaScript file is loadable via HTTPS

### Email Not Being Processed

1. Check browser console for errors
2. Verify user has permission to read email headers
3. Ensure email is properly selected before clicking report
4. Check that recipient email addresses are valid

### Report Email Not Sending

1. In Outlook Web, manual send is required (secure by design)
2. For desktop, ensure configured email addresses are valid
3. Check network connectivity and mail server access

## Development Tips

### Using Nodemon for Auto-Reload

```bash
npm run dev
```

The server will automatically restart when files change.

### Debugging in Outlook Web

1. Open Developer Tools (F12)
2. Go to Console tab
3. Look for messages like "Reporting email as Phishing"
4. Check for any JavaScript errors

### Testing Configuration Changes

```javascript
// In function-file.js
console.log("CONFIG:", CONFIG);
```

Check browser console to verify configuration is loaded correctly.

## Performance Optimization

For production deployments:

1. **Minify JavaScript**: Use webpack or terser

   ```bash
   npm install --save-dev webpack webpack-cli
   ```

2. **Enable Caching**: Configure cache headers in server.js
3. **CDN Integration**: Serve assets via CDN
4. **Email Storage**: Implement database for report history

## Monitoring & Logging

Add monitoring to track:

- Number of reports submitted
- Report types (phishing, spam, legitimate)
- User engagement
- Error rates

## License

GNU General Public License v3.0 or later

## Support

For issues or questions:

- Email: support@company.com
- Report issues in project repository
- Check Office.js documentation: https://docs.microsoft.com/en-us/office/dev/add-ins/overview/index

## Related Resources

- [Office.js API Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/index)
- [Outlook Add-in Requirements](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-add-ins)
- [Manifest Schema Reference](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifest)
- [Office Add-in Best Practices](https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/best-practices)

## Version History

### v1.0.0 (Current)

- Initial release
- Support for Phishing, Spam, and Legitimate reporting
- Email metadata extraction
- GoPhish integration ready
- Outlook Web and Desktop support

## Credits

Developed by ADG Tech
Based on original VSTO implementation by Abdulla Albreiki

---

**Last Updated**: February 27, 2026
