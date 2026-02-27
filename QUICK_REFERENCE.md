# Quick Reference Guide

## Project Structure

```
OutlookWebAdd/
├── manifest.xml                # Office Add-in configuration (REQUIRED)
├── function-file.html          # Command handler entry point (REQUIRED)
├── function-file.js            # Main business logic (REQUIRED)
├── server.js                   # Express server for serving files
├── package.json                # npm dependencies
├── .env.example                # Environment variables template
├── .env                        # Local environment config (create from .env.example)
├── certs/                      # SSL certificates
│   ├── key.pem                # Private key
│   └── cert.pem               # Public certificate
├── public/
│   └── assets/                # Icons and images
│       ├── icon-16.png
│       ├── icon-32.png
│       ├── icon-80.png
│       ├── phishing-icon-16.png
│       ├── phishing-icon-32.png
│       ├── phishing-icon-80.png
│       ├── spam-icon-16.png
│       ├── spam-icon-32.png
│       ├── spam-icon-80.png
│       ├── legitimate-icon-16.png
│       ├── legitimate-icon-32.png
│       └── legitimate-icon-80.png
├── logs/                      # Application logs
├── setupsh                     # Setup script (macOS/Linux)
├── setup.bat                   # Setup script (Windows)
├── README.md                   # Project documentation
├── INTEGRATION.md              # Integration guide with external services
├── COMPARISON.md               # VSTO vs Office.js comparison
└── DEPLOYMENT.md               # Deployment instructions
```

## Common Commands

```bash
# Setup and Installation
npm install                    # Install dependencies
./setup.sh                     # Run setup (macOS/Linux)
setup.bat                      # Run setup (Windows)

# Development
npm start                      # Start development server
npm run dev                    # Start with auto-reload (requires nodemon)

# SSL Certificates (macOS/Linux)
openssl req -x509 -newkey rsa:2048 -keyout certs/key.pem -out certs/cert.pem -days 365 -nodes

# Testing
curl -k https://localhost:3000/manifest.xml    # Test manifest access
curl -k https://localhost:3000/health          # Test server health

# Production
NODE_ENV=production npm start
```

## Configuration

### Most Important Settings

**In `function-file.js` (Line 5-12):**

```javascript
const CONFIG = {
  infosecEmail: "security@company.com", // ← CHANGE THIS
  spamReportEmail: "spam@company.com", // ← CHANGE THIS
  supportEmail: "support@company.com", // ← CHANGE THIS
  gophishUrl: "https://gophish.company.com",
  gophishListenerPort: 3333,
  gophishCustomHeader: "X-SENTINEL-AJSMN",
};
```

**In `manifest.xml` (Line 7-8):**

```xml
<Id>12345678-1234-1234-1234-123456789012</Id>  <!-- Generate new UUID -->
<ProviderName>ADGSentinel</ProviderName>       <!-- Your organization name -->
```

**In `.env`:**

```
INFOSEC_EMAIL=security@company.com
SPAM_REPORT_EMAIL=spam@company.com
SUPPORT_EMAIL=support@company.com
NODE_ENV=development
```

## Manifest URL

| Environment       | URL                                        |
| ----------------- | ------------------------------------------ |
| Local Development | `https://localhost:3000/manifest.xml`      |
| Staging           | `https://staging.company.com/manifest.xml` |
| Production        | `https://adgin.company.com/manifest.xml`   |

## Key API Endpoints

| Endpoint                                | Method | Purpose                |
| --------------------------------------- | ------ | ---------------------- |
| `GET /manifest.xml`                     | GET    | Office Add-in manifest |
| `GET /function-file/function-file.html` | GET    | Command handler page   |
| `GET /scripts/function-file.js`         | GET    | Main script            |
| `GET /api/config`                       | GET    | Configuration API      |
| `GET /health`                           | GET    | Health check           |

## Function Export (Office.js)

Functions must be exported globally for Office.js to find them:

```javascript
// In function-file.js
window.reportPhishing = reportPhishing;
window.reportSpam = reportSpam;
window.reportLegitimate = reportLegitimate;
```

## Icons Required

Create icons in `public/assets/` (PNG format):

```
Required Dimensions:
- 16x16 (toolbar)
- 32x32 (dropdown)
- 80x80 (ribbon)

For each report type:
- icon-{16|32|80}.png (main)
- phishing-icon-{16|32|80}.png
- spam-icon-{16|32|80}.png
- legitimate-icon-{16|32|80}.png

Total: 12 images
```

## Common Issues

| Error                         | Solution                                                                     |
| ----------------------------- | ---------------------------------------------------------------------------- |
| "Cannot read property 'item'" | Office.js not initialized - check Office.onReady()                           |
| "SSL certificate problem"     | Use self-signed cert for dev, valid cert for prod                            |
| "Manifest not found"          | Check `/manifest.xml` is accessible via HTTPS                                |
| "Headers are undefined"       | getAllInternetHeadersAsync() failed - some headers unavailable in web client |
| "Email not deleted"           | Web Add-ins can't directly delete - use categories instead                   |
| "Report not sending"          | Web Add-in uses reply compose - user must send manually                      |

## Testing Workflow

```
1. Start server:     npm start
2. Verify server:    https://localhost:3000/health
3. Open Outlook:     https://outlook.office.com
4. Upload manifest:  Settings > Get Add-ins > Upload > https://localhost:3000/manifest.xml
5. Select email:     Click email to open it
6. Click Report:     Home > ADGSentinel Report > [Option]
7. Check logs:       Browser console (F12) for debug messages
8. Verify email:     Check if report email received
```

## Debugging Tips

```javascript
// Add console logs to track execution
console.log("Function called:", functionName);
console.log("Email data:", emailData);
console.log("Config:", CONFIG);

// Check Office context
console.log("Office context:", Office.context);

// Monitor async operations
item.getAllInternetHeadersAsync((result) => {
  console.log("Headers result:", result);
});

// Monitor notifications
console.log("Showing notification:", message);
```

## Office.js Common API Calls

```javascript
// Get current mail item
const item = Office.context.mailbox.item;

// Get user profile
const userProfile = Office.context.mailbox.userProfile;
console.log("User email:", userProfile.emailAddress);

// Get body type
item.body.getTypeAsync((result) => {
  console.log("Body type:", result.value); // Office.MailboxEnums.BodyType.HTML
});

// Get body content
item.body.getAsync(Office.MailboxEnums.BodyType.HTML, (result) => {
  console.log("Body content:", result.value);
});

// Get all headers
item.getAllInternetHeadersAsync((result) => {
  if (result.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Headers:", result.value);
  }
});

// Show notification
item.notificationMessages.addAsync("myNotification", {
  type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
  message: "Message text",
  persistent: false,
});

// Get diagnostics
const diags = Office.context.mailbox.diagnostics;
console.log("Host:", diags.hostName); // "Outlook"
console.log("Host Version:", diags.hostVersion);
```

## Git Commands

```bash
# Clone repo
git clone <repo-url>
cd OutlookWebAdd

# Create feature branch
git checkout -b feature/your-feature

# Commit changes
git add .
git commit -m "feat: describe your change"

# Push to remote
git push origin feature/your-feature

# Create Pull Request
# Go to GitHub/GitLab and create PR

# Switch to production
git checkout main
git pull origin main
git merge develop
git push origin main
```

## Environment Variables Reference

```bash
# Email Configuration
INFOSEC_EMAIL=security@company.com
SPAM_REPORT_EMAIL=spam@company.com
SUPPORT_EMAIL=support@company.com

# Server
PORT=3000
NODE_ENV=development|production

# GoPhish (Optional)
GOPHISH_URL=https://gophish.company.com
GOPHISH_LISTENER_PORT=3333
GOPHISH_CUSTOM_HEADER=X-SENTINEL-AJSMN

# External APIs (Optional)
EMAIL_API_ENDPOINT=https://api.company.com/reports
EMAIL_API_TOKEN=token_here
SLACK_WEBHOOK_URL=https://hooks.slack.com/...
LOG_ENDPOINT=https://logging.company.com/api

# Features
ENABLE_GOPHISH_INTEGRATION=true
ENABLE_EXTERNAL_API=false
LOG_LEVEL=info
```

## Useful Resources

- **Office.js Docs**: https://docs.microsoft.com/en-us/office/dev/add-ins/
- **Outlook API**: https://docs.microsoft.com/en-us/outlook/add-ins/
- **Manifest Schema**: https://docs.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifest
- **GoPhish**: https://getgophish.com/
- **Express.js**: https://expressjs.com/
- **SSL Certificates**: https://letsencrypt.org/

## Support & Contact

- **Email**: adgin-support@company.com
- **Documentation**: See README.md, INTEGRATION.md, DEPLOYMENT.md
- **Issues**: Report via email or issue tracker

## Version Info

- **Current Version**: 1.0.0
- **Last Updated**: February 27, 2026
- **Compatibility**: Office.js 1.1+, Node.js 12+, npm 6+

---

**Remember**: Always update configuration files before deploying!
