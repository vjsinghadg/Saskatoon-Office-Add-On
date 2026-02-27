# ADGSentinel Add-in - Integration Guide

## Overview

This document provides detailed integration instructions for connecting the ADGSentinel Outlook Web Add-in with external services like email systems, GoPhish, and reporting backends.

## Email Reporting Flow

```
User selects email
    ↓
User clicks "ADGSentinel Report" > [Phishing/Spam/Legitimate]
    ↓
Office.js captures email metadata & headers
    ↓
Check for GoPhish simulated phishing campaign
    ↓
Create report email with extracted details
    ↓
Route to appropriate recipient (InfoSec/Spam/Support)
    ↓
Mark original email & notify user
```

## Configuration

### 1. InfoSec Email Configuration

In `function-file.js`, line 5-12:

```javascript
const CONFIG = {
  infosecEmail: "security-team@company.com",
  spamReportEmail: "spam-reports@company.com",
  supportEmail: "addin-support@company.com",
  // ... other settings
};
```

### 2. Email Routing Logic

The add-in automatically routes reports to:

| Report Type | Default Recipient        | Environment Variable |
| ----------- | ------------------------ | -------------------- |
| Phishing    | `CONFIG.infosecEmail`    | `INFOSEC_EMAIL`      |
| Spam        | `CONFIG.spamReportEmail` | `SPAM_REPORT_EMAIL`  |
| Legitimate  | `CONFIG.infosecEmail`    | `INFOSEC_EMAIL`      |

## Report Email Format

### Subject Line

```
[SENTINEL-PHISHING] Original email subject
[SENTINEL-SPAM] Original email subject
[SENTINEL-LEGITIMATE] Original email subject
```

### Body Structure

```html
Report Type: [Phishing|Spam|Legitimate] Report Time: ISO 8601 timestamp Reported
by: User Display Name (user@company.com) --- Email Information Subject: Original
subject From: sender@external.com To: recipient@company.com CC: [if applicable]
Attachments: 0-N --- URLs Found (N) - https://suspicious[.]domain[.]com -
https://another[.]bad[.]domain[.]com [URLs have : replaced with [:] for safety]
--- Email Headers [Full email header block - 2000+ characters typically] ---
Original Email Body [Complete email HTML or text body] --- ADGSentinel Report
Add-in v1.0 | Powered by Office.js
```

## GoPhish Integration

### Detection Method

The add-in detects GoPhish simulated phishing campaigns by checking for custom headers:

```javascript
// Custom header detection (line 260)
if (
  emailData.headers &&
  emailData.headers.indexOf(CONFIG.gophishCustomHeader) > -1
) {
  return true; // Simulated phishing detected
}
```

### GoPhish Setup

1. **Configure GoPhish Sending Profile**:
   - In GoPhish admin: Create new Sending Profile
   - Add custom header: `X-SENTINEL-AJSMN`
   - Set value to: `{{.RId}}` (GoPhish placeholder for tracking ID)

2. **Track Campaign Reports**:
   - GoPhish logs reports via webhook when custom header is detected
   - Maps reports to campaigns by RId

3. **Reporting URL**:
   ```
   https://gophish.your-domain.com:3333/api/report?rid={{.RId}}
   ```

### Example GoPhish Configuration

```json
{
  "sending_profile": "ADGSentinel Integration",
  "headers": {
    "X-SENTINEL-AJSMN": "{{.RId}}"
  },
  "from": "sender@company.com",
  "host": "mail.company.com:587",
  "username": "adcsender@company.com",
  "password": "encrypted_password"
}
```

## External API Integration

### Email API Integration

To send reports via external email API instead of Outlook:

```javascript
// Add to function-file.js after reportPhishingEmailToSecurityTeam()

async function sendViaExternalAPI(emailData, reportType) {
  const apiUrl = "https://api.company.com/reports/phishing";

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${API_TOKEN}`,
      },
      body: JSON.stringify({
        report_type: reportType,
        email_subject: emailData.subject,
        email_from: emailData.from,
        email_headers: emailData.headers,
        reporter_email: emailData.userInfo.email,
        timestamp: emailData.timestamp,
        is_simulated_phishing: emailData.isSimulatedPhishing,
        urls_found: extractUrls(emailData.body),
        attachment_count: emailData.attachmentCount,
      }),
    });

    if (!response.ok) {
      throw new Error(`API Error: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error("API Integration Error:", error);
    throw error;
  }
}
```

### Slack Notification Integration

```javascript
async function notifyViaSlack(emailData, reportType) {
  const webhookUrl = process.env.SLACK_WEBHOOK_URL;

  const message = {
    text: `New ${reportType} Report from ADGSentinel`,
    blocks: [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*${reportType} Email Report*`,
        },
      },
      {
        type: "section",
        fields: [
          {
            type: "mrkdwn",
            text: `*Reporter:*\n${emailData.userInfo.displayName}`,
          },
          {
            type: "mrkdwn",
            text: `*Report Type:*\n${reportType}`,
          },
          {
            type: "mrkdwn",
            text: `*From:*\n${emailData.from}`,
          },
          {
            type: "mrkdwn",
            text: `*Subject:*\n${emailData.subject.substring(0, 50)}...`,
          },
        ],
      },
    ],
  };

  try {
    const response = await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(message),
    });

    if (!response.ok) {
      console.error("Slack notification failed:", response.statusText);
    }
  } catch (error) {
    console.error("Slack integration error:", error);
  }
}
```

### Azure Service Bus Integration

```javascript
async function publishToServiceBus(emailData, reportType) {
  const endpoint = process.env.AZURE_SERVICE_BUS_ENDPOINT;
  const accessKey = process.env.AZURE_SERVICE_BUS_KEY;

  const message = {
    report_type: reportType,
    email_metadata: {
      subject: emailData.subject,
      from: emailData.from,
      headers: emailData.headers,
    },
    reporter: emailData.userInfo,
    timestamp: emailData.timestamp,
  };

  try {
    const response = await fetch(`${endpoint}/messages`, {
      method: "POST",
      headers: {
        Authorization: `SharedAccessSignature ${accessKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(message),
    });

    if (!response.ok) {
      throw new Error(`Service Bus Error: ${response.statusText}`);
    }
  } catch (error) {
    console.error("Azure Service Bus error:", error);
  }
}
```

## Database Integration

### Store Report History

```javascript
async function storeReportInDatabase(emailData, reportType) {
  const dbUrl = process.env.DB_API_ENDPOINT;

  const report = {
    type: reportType,
    email_id: emailData.messageId,
    subject: emailData.subject,
    from: emailData.from,
    reporter_email: emailData.userInfo.email,
    reporter_name: emailData.userInfo.displayName,
    timestamp: emailData.timestamp,
    is_simulated_phishing: emailData.isSimulatedPhishing,
    urls_count: extractUrls(emailData.body).length,
    attachment_count: emailData.attachmentCount,
    status: "reported",
  };

  try {
    const response = await fetch(`${dbUrl}/api/reports`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${DB_TOKEN}`,
      },
      body: JSON.stringify(report),
    });

    return await response.json();
  } catch (error) {
    console.error("Database storage error:", error);
    throw error;
  }
}
```

## Analytics & Metrics

### Track Report Statistics

```javascript
async function trackMetrics(reportType, success) {
  const analyticsUrl = "https://analytics.company.com/api/track";

  const event = {
    event_type: "phishing_report",
    report_type: reportType,
    success: success,
    timestamp: new Date().toISOString(),
    user: Office.context.mailbox.userProfile.emailAddress,
  };

  try {
    await fetch(analyticsUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(event),
      keepalive: true,
    });
  } catch (error) {
    console.warn("Analytics tracking failed:", error);
  }
}
```

## Error Handling & Logging

### Centralized Error Logging

```javascript
async function logError(error, context) {
  const loggingUrl = process.env.LOG_ENDPOINT;

  const errorReport = {
    error: error.message,
    stack: error.stack,
    context: context,
    timestamp: new Date().toISOString(),
    user: Office.context.mailbox.userProfile.emailAddress,
    browser: navigator.userAgent,
  };

  try {
    await fetch(loggingUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(errorReport),
    });
  } catch (logError) {
    console.error("Failed to log error:", logError);
  }
}
```

## Security Headers

### Add Security Headers to Manifest

```xml
<VersionOverrides>
  <Hosts>
    <Host xsi:type="MailHost">
      <DesktopFormFactor>
        <FunctionFile resid="functionfile"/>
        <!-- CSP Header would be set by server -->
      </DesktopFormFactor>
    </Host>
  </Hosts>
</VersionOverrides>
```

### Server-Side Security Headers

```javascript
// In server.js
app.use((req, res, next) => {
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "SAMEORIGIN");
  res.setHeader(
    "Content-Security-Policy",
    "default-src 'self' https://appsforoffice.microsoft.com",
  );
  next();
});
```

## Testing Integrations

### Test Email Reporting

```bash
# Test local server
curl -k https://localhost:3000/health

# Test manifest accessibility
curl -k https://localhost:3000/manifest.xml
```

### Test GoPhish Integration

1. Create test campaign in GoPhish
2. Add custom header in Sending Profile
3. Make report from Outlook
4. Check GoPhish portal for report update

### Test External APIs

```javascript
// Add test button in Dev Tools
async function testExternalAPI() {
  const testData = {
    subject: "Test Email",
    from: "test@example.com",
    body: "Test body with https://test.com link",
    headers: "Test-Header: value",
    userInfo: { displayName: "Test User", email: "test@company.com" },
  };

  try {
    const result = await sendViaExternalAPI(testData, "Phishing");
    console.log("API Test Result:", result);
  } catch (error) {
    console.error("API Test Failed:", error);
  }
}
```

## Troubleshooting Integration Issues

| Issue                       | Solution                                                            |
| --------------------------- | ------------------------------------------------------------------- |
| Emails not reaching InfoSec | Verify email address in CONFIG and Exchange settings                |
| GoPhish detection fails     | Check custom header spelling matches exactly                        |
| External API timeouts       | Increase timeout, check network access, verify API endpoint         |
| Reports missing attachments | Web add-ins have limited attachment access                          |
| Header extraction fails     | Some headers unavailable in web client; use desktop for full access |

## Performance Considerations

1. **Batch Report Processing**: For high-volume scenarios, queue reports
2. **Caching**: Cache email headers temporarily to reduce API calls
3. **Async Processing**: Use async/await to prevent blocking
4. **Rate Limiting**: Implement on server side to prevent abuse

---

For additional support, contact: adgin-support@company.com
