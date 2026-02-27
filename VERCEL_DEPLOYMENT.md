# Vercel Deployment Guide - ADGSentinel Outlook Add-in

## Overview

This project has been restructured to work with Vercel's serverless architecture. Instead of running a traditional Express.js server, all endpoints are now serverless functions in the `/api` directory.

## What Changed

- **Removed:** `https.createServer()` and certificate file handling (incompatible with serverless)
- **Added:** Individual serverless functions in `/api/` directory:
  - `api/manifest.xml.js` - Serves manifest.xml
  - `api/function-file-html.js` - Serves function-file.html
  - `api/function-file-js.js` - Serves function-file.js
  - `api/config.js` - Provides configuration endpoint
  - `api/health.js` - Health check endpoint
- **Updated:** `vercel.json` to route requests to serverless functions
- **Preserved:** Local development still works with `npm start` (HTTPS on localhost:3000)

## Prerequisites

1. Vercel account (free tier: vercel.com)
2. Git/GitHub account with your project uploaded
3. Node.js 14+ installed locally
4. npm package manager

## Deployment Steps

### Step 1: Push Code to GitHub

```bash
# Initialize git (if not already done)
git init
git add -A
git commit -m "Convert to Vercel serverless architecture"

# Add GitHub remote (replace YOUR_USERNAME/YOUR_REPO)
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git branch -M main
git push -u origin main
```

### Step 2: Create Vercel Project

1. Go to https://vercel.com/dashboard
2. Click "Add New" → "Project"
3. Import your GitHub repository
4. Select your OutlookWebAdd repository
5. Click "Import"

### Step 3: Configure Environment Variables

In Vercel Dashboard:

1. Go to your project settings
2. Click "Environment Variables"
3. Add the following variables:
   ```
   INFOSEC_EMAIL = infosec@company.com
   SPAM_REPORT_EMAIL = spam-report@company.com
   SUPPORT_EMAIL = support@company.com
   GOPHISH_URL = https://saskaatoon.ca
   ```
4. Save changes

### Step 4: Deploy

1. Vercel will auto-deploy when you push to main branch
2. Your project URL will be: `https://your-project.vercel.app`
3. Manifest will be at: `https://your-project.vercel.app/manifest.xml`

### Step 5: Configure Custom Domain (Optional but Recommended)

1. In Vercel Dashboard, go to "Domains"
2. Click "Add Domain"
3. Enter your domain (e.g., `saskatoonapi.adgstaging.in`)
4. Follow DNS configuration instructions (CNAME or A record)
5. DNS propagation takes 24-48 hours

### Step 6: Update Manifest for Production

If using custom domain, update [manifest.xml](manifest.xml):

```xml
<!-- Change all occurrences of: -->
https://your-project.vercel.app/
<!-- To: -->
https://saskatoonapi.adgstaging.in/
```

## Testing the Deployment

### Via Browser

```bash
# Manifest endpoint
curl https://your-project.vercel.app/manifest.xml

# Health check
curl https://your-project.vercel.app/health

# Config endpoint
curl https://your-project.vercel.app/api/config
```

### Via Outlook Web

1. Navigate to Outlook Web (outlook.office.com)
2. Click Settings ⚙️ → "Get Add-ins"
3. Select "My Add-ins" → "Upload My Add-in"
4. Enter manifest URL: `https://your-project.vercel.app/manifest.xml`
5. Click "Upload"
6. Add the add-in to your mailbox
7. Test the "ADGSentinel Report" dropdown in the ribbon

## Local Development

### Run Locally (HTTPS on localhost:3000)

```bash
# Install dependencies
npm install

# Generate self-signed certificates (one time)
npm run setup

# Start local server
npm start
```

### Test with Vercel CLI (Optional)

```bash
# Install Vercel CLI
npm install -g vercel

# Run locally simulating Vercel environment
vercel dev
```

## Troubleshooting

### Issue: "Function failed to compile"

**Solution:** Confirm all files in `/api/` directory use ES module syntax (`export default function`)

### Issue: "Cannot find module"

**Solution:** Ensure `manifest.xml`, `function-file.html`, and `function-file.js` exist in project root

### Issue: "404 Not Found" for manifest.xml

**Solution:** Verify vercel.json routes are correct and deployment completed successfully

### Issue: "Environment variables not loading"

**Solution:** Re-deploy project after adding environment variables (Vercel doesn't hot-reload env vars)

## File Structure After Restructuring

```
OutlookWebAdd/
├── api/
│   ├── manifest.xml.js          # NEW: Serverless function for manifest
│   ├── function-file-html.js    # NEW: Serverless function for HTML
│   ├── function-file-js.js      # NEW: Serverless function for JavaScript
│   ├── config.js                # NEW: Serverless function for config
│   └── health.js                # NEW: Serverless function for health
├── public/
│   └── assets/                  # Icons and images
├── manifest.xml
├── function-file.html
├── function-file.js
├── server.js                    # UPDATED: Supports both local & serverless
├── package.json
├── vercel.json                  # UPDATED: New route configuration
└── README.md
```

## Key Changes in server.js

- **Lines 1-55:** Express app setup (shared by local & serverless)
- **Lines 57-82:** Local HTTPS server startup (runs on `npm start`)
- **Line 85:** `module.exports = app` exports for Vercel serverless

## Why These Changes?

- **Serverless compatibility:** Vercel can't run traditional Express servers with PORT binding
- **No certificate needed:** Vercel auto-provisions HTTPS via Let's Encrypt
- **Scalability:** Serverless functions scale automatically
- **Cost:** Only pay for actual usage, not always-running servers
- **Local development:** Still works with HTTPS for testing before deployment

## Next Steps

1. Push code to GitHub
2. Deploy via Vercel Dashboard
3. Configure custom domain (if using saskatoonapi.adgstaging.in)
4. Upload manifest to Microsoft 365 Admin Center
5. Test in Outlook Web with real users

## Support

For issues with Vercel deployment:

- Check Vercel logs: Dashboard → Deployments → Function logs
- Review vercel.json configuration
- Verify environment variables are set
- Ensure all API route files export default handler functions
