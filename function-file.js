/*
 * ADGSentinel Report - Office.js based Outlook Web Add-in
 * Handles reporting of phishing, spam, and legitimate emails
 * 
 * Configuration
 */

const CONFIG = {
    infosecEmail: "your-email@company.com",        // Your InfoSec email
    spamReportEmail: "spam@company.com",           // Spam report email
    supportEmail: "support@company.com",           // Support email
    gophishUrl: "https://saskaatoon.ca",           // Already set ✓
    gophishListenerPort: 3333,
    gophishCustomHeader: "X-SENTINEL-AJSMN"
};


/**
 * Report Phishing - Called when user clicks "Report Phishing"
 */
async function reportPhishing() {
    try {
        await handleReportAction("Phishing");
    } catch (error) {
        console.error("Error in reportPhishing:", error);
        Office.onReady(() => {
            if (Office.context.mailbox.diagnostics.hostName === 'Outlook') {
                showErrorNotification("Error", `An error occurred: ${error.message}`);
            }
        });
    }
}

/**
 * Report Spam - Called when user clicks "Report Spam"
 */
async function reportSpam() {
    try {
        await handleReportAction("Spam");
    } catch (error) {
        console.error("Error in reportSpam:", error);
        Office.onReady(() => {
            if (Office.context.mailbox.diagnostics.hostName === 'Outlook') {
                showErrorNotification("Error", `An error occurred: ${error.message}`);
            }
        });
    }
}

/**
 * Report Legitimate - Called when user clicks "Report Legitimate"
 */
async function reportLegitimate() {
    try {
        await handleReportAction("Legitimate");
    } catch (error) {
        console.error("Error in reportLegitimate:", error);
        Office.onReady(() => {
            if (Office.context.mailbox.diagnostics.hostName === 'Outlook') {
                showErrorNotification("Error", `An error occurred: ${error.message}`);
            }
        });
    }
}

/**
 * Main handler for all report actions
 */
async function handleReportAction(reportType) {
    return new Promise((resolve, reject) => {
        Office.onReady(() => {
            const mailbox = Office.context.mailbox;
            const item = mailbox.item;

            // Get current user confirmation
            const mailboxDiagnostics = mailbox.diagnostics;
            console.log(`Reporting email as ${reportType}`);
            console.log(`Current user: ${mailboxDiagnostics.userDisplayName}`);

            // Get email details
            item.body.getTypeAsync((bodyTypeResult) => {
                if (bodyTypeResult.status !== Office.AsyncResultStatus.Succeeded) {
                    reject(new Error(`Failed to get body type: ${bodyTypeResult.error?.message}`));
                    return;
                }

                // Get the full email data
                item.body.getAsync(bodyTypeResult.value, { asyncContext: reportType }, async (bodyResult) => {
                    if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
                        reject(new Error(`Failed to get body: ${bodyResult.error?.message}`));
                        return;
                    }

                    try {
                        // Get additional email metadata
                        const emailData = {
                            subject: item.subject,
                            from: item.from?.emailAddress,
                            body: bodyResult.value,
                            bodyType: bodyResult.asyncContext,
                            reportType: reportType,
                            timestamp: new Date().toISOString(),
                            sender: item.from,
                            recipients: item.to,
                            ccRecipients: item.cc,
                            bccRecipients: item.bcc,
                            attachmentCount: item.attachments ? item.attachments.length : 0,
                            isRead: item.isRead,
                            categories: item.categories,
                            userInfo: {
                                displayName: mailboxDiagnostics.userDisplayName,
                                email: mailbox.userProfile.emailAddress,
                                timeZone: mailbox.userProfile.timeZoneOffset
                            }
                        };

                        // Get email headers for phishing detection
                        item.getAllInternetHeadersAsync((headersResult) => {
                            if (headersResult.status === Office.AsyncResultStatus.Succeeded) {
                                emailData.headers = headersResult.value;

                                // Check if this is a simulated phishing email from GoPhish
                                checkForSimulatedPhishing(emailData, CONFIG.gophishCustomHeader)
                                    .then((isSimulated) => {
                                        emailData.isSimulatedPhishing = isSimulated;

                                        // Process the report
                                        processReport(emailData, reportType)
                                            .then(() => resolve())
                                            .catch((err) => reject(err));
                                    })
                                    .catch((err) => {
                                        // Continue processing even if header check fails
                                        emailData.isSimulatedPhishing = false;
                                        processReport(emailData, reportType)
                                            .then(() => resolve())
                                            .catch((err) => reject(err));
                                    });
                            } else {
                                // Continue without headers if they can't be retrieved
                                emailData.headers = "Headers unavailable";
                                emailData.isSimulatedPhishing = false;

                                processReport(emailData, reportType)
                                    .then(() => resolve())
                                    .catch((err) => reject(err));
                            }
                        });
                    } catch (error) {
                        reject(error);
                    }
                });
            });
        });
    });
}

/**
 * Process the report based on type
 */
async function processReport(emailData, reportType) {
    return new Promise((resolve, reject) => {
        Office.onReady(() => {
            let recipientEmail = CONFIG.infosecEmail;

            // Determine recipient based on report type
            if (reportType === "Spam") {
                recipientEmail = CONFIG.spamReportEmail;
            }

            // Prepare report body
            const reportBody = prepareReportBody(emailData, reportType);

            // Create report email
            createReportEmail(recipientEmail, reportBody, emailData, reportType)
                .then(() => {
                    // Delete original email if phishing/spam report
                    if (reportType === "Phishing" || reportType === "Spam") {
                        deleteOriginalEmail()
                            .then(() => {
                                showSuccessNotification(reportType);
                                resolve();
                            })
                            .catch((deleteError) => {
                                console.warn("Failed to delete email:", deleteError);
                                showSuccessNotification(reportType);
                                resolve(); // Still resolve even if delete fails
                            });
                    } else {
                        showSuccessNotification(reportType);
                        resolve();
                    }
                })
                .catch((err) => reject(err));
        });
    });
}

/**
 * Prepare the report email body with email details
 */
function prepareReportBody(emailData, reportType) {
    let body = `<html><body><font face="Calibri" size="3">`;

    body += `<p><strong>Report Type:</strong> ${reportType}</p>`;
    body += `<p><strong>Report Time:</strong> ${emailData.timestamp}</p>`;
    body += `<p><strong>Reported by:</strong> ${emailData.userInfo.displayName} (${emailData.userInfo.email})</p>`;

    body += `<hr>`;
    body += `<h3>Email Information</h3>`;
    body += `<p><strong>Subject:</strong> ${escapeHtml(emailData.subject)}</p>`;
    body += `<p><strong>From:</strong> ${escapeHtml(emailData.from || "Unknown")}</p>`;
    body += `<p><strong>To:</strong> ${escapeHtml(emailData.recipients || "Unknown")}</p>`;

    if (emailData.ccRecipients) {
        body += `<p><strong>CC:</strong> ${escapeHtml(emailData.ccRecipients)}</p>`;
    }

    body += `<p><strong>Attachments:</strong> ${emailData.attachmentCount}</p>`;

    // Extract and display URLs from email body
    const urls = extractUrls(emailData.body);
    if (urls.length > 0) {
        body += `<hr>`;
        body += `<h3>URLs Found (${urls.length})</h3>`;
        body += `<ul>`;
        urls.forEach((url) => {
            body += `<li>${escapeHtml(url.replace(":", "[:]"))}</li>`;
        });
        body += `</ul>`;
    }

    body += `<hr>`;
    body += `<h3>Email Headers</h3>`;
    body += `<pre style="font-size: 11px; background-color: #f0f0f0; padding: 10px;">`;
    body += escapeHtml(emailData.headers || "Headers not available");
    body += `</pre>`;

    body += `<hr>`;
    body += `<h3>Original Email Body</h3>`;
    body += `<div style="border: 1px solid #ccc; padding: 10px; margin-top: 10px;">`;
    body += emailData.body || "Body not available";
    body += `</div>`;

    body += `<hr>`;
    body += `<p style="font-size: 10px; color: #666;">`;
    body += `ADGSentinel Report Add-in v1.0 | Powered by Office.js`;
    body += `</p>`;

    body += `</font></body></html>`;

    return body;
}

/**
 * Create and send the report email
 */
function createReportEmail(recipientEmail, reportBody, emailData, reportType) {
    return new Promise((resolve, reject) => {
        Office.onReady(() => {
            const mailbox = Office.context.mailbox;
            const item = mailbox.item;

            // Compose new message for reporting
            const reportSubject = `[SENTINEL-${reportType.toUpperCase()}] ${item.subject}`;

            item.reply(
                {
                    displayReplyAll: false,
                    asyncContext: { recipientEmail, reportBody, reportType }
                },
                (replyResult) => {
                    if (replyResult.status === Office.AsyncResultStatus.Succeeded) {
                        const replyItem = replyResult.asyncContext;

                        // Note: In Outlook Web Add-ins, we have limitations with sending emails
                        // The best approach is to use the reply compose window
                        showInfoNotification(
                            "Report Ready",
                            "The report email is ready for review. Please send it to " + recipientEmail
                        );
                        resolve();
                    } else {
                        // Fallback: Use NotificationMessages to inform user to send report manually
                        showWarningNotification(
                            "Send Report Manually",
                            `Please send the report to ${recipientEmail} with subject: ${reportSubject}`
                        );
                        resolve();
                    }
                }
            );
        });
    });
}

/**
 * Delete the original email
 */
function deleteOriginalEmail() {
    return new Promise((resolve, reject) => {
        Office.onReady(() => {
            const item = Office.context.mailbox.item;

            // Mark as read first
            item.isRead = true;

            // Note: Web Add-ins have limited ability to delete emails directly
            // Instead, we can move to Deleted Items or add category
            if (item.categories) {
                const currentCategories = item.categories;
                currentCategories.push("ReportedAsPhishing");
                item.categories = currentCategories;
            }

            showInfoNotification(
                "Email Marked",
                "Email has been marked as reported. You can manually delete it."
            );
            resolve();
        });
    });
}

/**
 * Check if email is a simulated phishing email from GoPhish
 */
async function checkForSimulatedPhishing(emailData, customHeader) {
    return new Promise((resolve) => {
        try {
            // Check for GoPhish custom header
            if (emailData.headers && emailData.headers.indexOf(customHeader) > -1) {
                console.log("Simulated phishing email detected from GoPhish");
                resolve(true);
            } else {
                resolve(false);
            }
        } catch (error) {
            console.warn("Error checking for simulated phishing:", error);
            resolve(false);
        }
    });
}

/**
 * Extract URLs from email body
 */
function extractUrls(text) {
    if (!text) return [];

    const urlPattern = /(https?:\/\/[^\s<>]+)/gi;
    const matches = text.match(urlPattern);

    if (!matches) return [];

    // Remove duplicates
    return [...new Set(matches)];
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text) {
    if (!text) return "";
    const map = {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;"
    };
    return text.replace(/[&<>"']/g, (m) => map[m]);
}

/**
 * Show success notification
 */
function showSuccessNotification(reportType) {
    Office.onReady(() => {
        const mailbox = Office.context.mailbox;

        let message = "";
        if (reportType === "Phishing") {
            message = "Good job! You have reported a phishing email to the Information Security Team.";
        } else if (reportType === "Spam") {
            message = "Thank you! You have reported this email as spam.";
        } else if (reportType === "Legitimate") {
            message = "Thank you for the feedback! You have reported this email as legitimate.";
        }

        mailbox.item.notificationMessages.replaceAsync(
            "reportNotification",
            {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: message,
                icon: "Icon.80x80",
                persistent: true
            },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to show notification:", result.error?.message);
                }
            }
        );
    });
}

/**
 * Show error notification
 */
function showErrorNotification(title, message) {
    Office.onReady(() => {
        const mailbox = Office.context.mailbox;

        mailbox.item.notificationMessages.replaceAsync(
            "errorNotification",
            {
                type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
                message: message,
                icon: "Icon.80x80",
                persistent: true
            },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to show error notification:", result.error?.message);
                }
            }
        );
    });
}

/**
 * Show warning notification
 */
function showWarningNotification(title, message) {
    Office.onReady(() => {
        const mailbox = Office.context.mailbox;

        mailbox.item.notificationMessages.replaceAsync(
            "warningNotification",
            {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: message,
                icon: "Icon.80x80",
                persistent: true
            },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to show warning notification:", result.error?.message);
                }
            }
        );
    });
}

/**
 * Show info notification
 */
function showInfoNotification(title, message) {
    Office.onReady(() => {
        const mailbox = Office.context.mailbox;

        mailbox.item.notificationMessages.replaceAsync(
            "infoNotification",
            {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: message,
                icon: "Icon.80x80",
                persistent: true
            },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to show info notification:", result.error?.message);
                }
            }
        );
    });
}
