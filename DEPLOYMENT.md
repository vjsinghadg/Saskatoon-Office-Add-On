# Deployment Guide - ADGSentinel Outlook Web Add-in

## Table of Contents

1. [Pre-Deployment Checklist](#pre-deployment-checklist)
2. [Development Deployment](#development-deployment)
3. [Staging Deployment](#staging-deployment)
4. [Production Deployment](#production-deployment)
5. [Azure App Service](#azure-app-service)
6. [GitHub Pages](#github-pages)
7. [AWS Deployment](#aws-deployment)
8. [Docker Deployment](#docker-deployment)
9. [Troubleshooting](#troubleshooting)

---

## Pre-Deployment Checklist

Before deploying, ensure:

- [ ] Node.js 12+ installed
- [ ] npm dependencies installed (`npm install`)
- [ ] SSL certificates generated (development) or obtained (production)
- [ ] `.env` file configured with correct email addresses
- [ ] `manifest.xml` ID is unique (UUID format)
- [ ] All email addresses in CONFIG are valid
- [ ] Icons created in `public/assets/` (16x16, 32x32, 80x80 PNG)
- [ ] Code tested locally (`npm start`)
- [ ] HTTPS working with valid certificate
- [ ] Office.js library accessible (CDN)
- [ ] All external API endpoints accessible
- [ ] Firewall/network rules allow HTTPS (port 443)

---

## Development Deployment

### Local Setup

```bash
# 1. Install dependencies
npm install

# 2. Generate self-signed certificates
mkdir -p certs
openssl req -x509 -newkey rsa:2048 -keyout certs/key.pem -out certs/cert.pem \
    -days 365 -nodes

# 3. Configure .env
cp .env.example .env
# Edit .env with your settings

# 4. Start development server
npm start
```

Server runs at: `https://localhost:3000`

### Testing in Outlook Web

1. Go to: `https://outlook.office.com`
2. Click **Settings** (gear icon) > **View all Outlook settings**
3. Go to **Add-ins** section
4. Click **Get Add-ins**
5. Select **My Add-ins**
6. Click **Upload My Add-in**
7. Choose **Upload from URL**
8. Enter: `https://localhost:3000/manifest.xml`

**Note:** Self-signed certificate warning is normal. Click "Advanced" and "Proceed to localhost"

---

## Staging Deployment

### Prerequisites

- Staging domain (e.g., `staging.company.com`)
- Valid SSL certificate for staging domain
- Staging web server or hosting service

### Steps

#### 1. Create Staging Configuration

```bash
# Create staging .env
cp .env.example .env.staging

# Edit .env.staging with staging settings
INFOSEC_EMAIL=staging-security@company.com
PORT=3000
NODE_ENV=staging
```

#### 2. Update Manifest for Staging

Create `manifest-staging.xml`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp ...>
  <Id>87654321-4321-4321-4321-210987654321</Id>
  <ProviderName>ADGSentinel Staging</ProviderName>
  <DisplayName DefaultValue="ADGSentinel Report (Staging)"/>

  <!-- Update resource URLs to staging domain -->
  <Resources>
    <bt:Urls>
      <bt:Url id="functionfile" DefaultValue="https://staging.company.com/function-file/function-file.html"/>
    </bt:Urls>
    <bt:Images>
      <bt:Image id="icon16" DefaultValue="https://staging.company.com/assets/icon-16.png"/>
      <!-- ... other URLs ... -->
    </bt:Images>
  </Resources>

  <!-- ... rest of manifest ... -->
</OfficeApp>
```

#### 3. Deploy to Staging Server

```bash
# Using SSH/SCP
scp -r . user@staging.company.com:/var/www/staging-addin/

# SSH into server
ssh user@staging.company.com

# Install and start
cd /var/www/staging-addin/
npm install --production
npm start
```

#### 4. Test in Outlook

Upload `https://staging.company.com/manifest-staging.xml` to Outlook

---

## Production Deployment

### Prerequisites

- Production domain with SSL certificate (Let's Encrypt or paid)
- Production web server or cloud service
- Backup and disaster recovery plan
- Monitoring and alerting setup

### Steps

#### 1. Prepare Production Configuration

```bash
# Create production .env
cp .env.example .env.production

# Edit with production values
cat << 'EOF' > .env.production
INFOSEC_EMAIL=security-team@company.com
SPAM_REPORT_EMAIL=spam-reports@company.com
SUPPORT_EMAIL=addin-support@company.com
PORT=3000
NODE_ENV=production
LOG_LEVEL=warn
EOF
```

#### 2. Update Production Manifest

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp ...>
  <Id>12345678-1234-1234-1234-123456789012</Id>
  <ProviderName>ADGSentinel</ProviderName>
  <DisplayName DefaultValue="ADGSentinel Report"/>

  <Resources>
    <bt:Urls>
      <bt:Url id="functionfile" DefaultValue="https://adgin.company.com/function-file/function-file.html"/>
    </bt:Urls>
    <bt:Images>
      <bt:Image id="icon16" DefaultValue="https://adgin.company.com/assets/icon-16.png"/>
    </bt:Images>
  </Resources>
</OfficeApp>
```

#### 3. Deploy to Production Server

Using **Linux/Unix server**:

```bash
# 1. SSH into production server
ssh deploy@adgin.company.com

# 2. Navigate to application directory
cd /opt/adgin

# 3. Pull latest code (if using git)
git pull origin main

# 4. Install dependencies
npm install --production

# 5. Stop old process
sudo systemctl stop adgin

# 6. Start new process
sudo systemctl start adgin

# 7. Verify status
sudo systemctl status adgin
```

#### 4. Setup systemd Service (Linux)

Create `/etc/systemd/system/adgin.service`:

```ini
[Unit]
Description=ADGSentinel Outlook Web Add-in
After=network.target

[Service]
Type=simple
User=adgin
WorkingDirectory=/opt/adgin
ExecStart=/usr/bin/node /opt/adgin/server.js
Restart=on-failure
RestartSec=10

Environment="NODE_ENV=production"
EnvironmentFile=/opt/adgin/.env.production

StandardOutput=append:/var/log/adgin/output.log
StandardError=append:/var/log/adgin/error.log

[Install]
WantedBy=multi-user.target
```

Enable and start:

```bash
sudo systemctl daemon-reload
sudo systemctl enable adgin
sudo systemctl start adgin
```

#### 5. Setup Nginx Reverse Proxy

Create `/etc/nginx/sites-available/adgin.conf`:

```nginx
server {
    listen 443 ssl http2;
    server_name adgin.company.com;

    ssl_certificate /etc/letsencrypt/live/adgin.company.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/adgin.company.com/privkey.pem;

    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers HIGH:!aNULL:!MD5;

    # Security headers
    add_header X-Content-Type-Options "nosniff" always;
    add_header X-Frame-Options "SAMEORIGIN" always;
    add_header X-XSS-Protection "1; mode=block" always;
    add_header Referrer-Policy "no-referrer-when-downgrade" always;

    # CSP for Office Add-in
    add_header Content-Security-Policy "default-src 'self' https://appsforoffice.microsoft.com https://office.com; script-src 'self' https://appsforoffice.microsoft.com" always;

    location / {
        proxy_pass https://localhost:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_cache_bypass $http_upgrade;
    }

    # Cache static assets
    location ~* \.(png|jpg|jpeg|gif|ico|css|js)$ {
        expires 1y;
        add_header Cache-Control "public, immutable";
    }

    # Gzip compression
    gzip on;
    gzip_types text/plain text/css application/json application/javascript text/xml application/xml;
}

# Redirect HTTP to HTTPS
server {
    listen 80;
    server_name adgin.company.com;
    return 301 https://$server_name$request_uri;
}
```

Enable:

```bash
sudo ln -s /etc/nginx/sites-available/adgin.conf /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl restart nginx
```

#### 6. Setup SSL with Let's Encrypt

```bash
sudo apt-get install certbot python3-certbot-nginx
sudo certbot certonly --nginx -d adgin.company.com
sudo certbot renew --dry-run  # Test auto-renewal
```

---

## Azure App Service

### Prerequisites

- Azure subscription
- Azure CLI installed
- Resource group created

### Steps

#### 1. Create App Service

```bash
# Variables
RG="adgin-rg"
APP_NAME="adgin-outlook-app"
PLAN="adgin-plan"
REGION="eastus"

# Create resource group
az group create --name $RG --location $REGION

# Create App Service plan
az appservice plan create --name $PLAN --resource-group $RG --sku B1 --is-linux

# Create web app
az webapp create --resource-group $RG --plan $PLAN --name $APP_NAME --runtime "NODE|16-lts"
```

#### 2. Configure Environment

```bash
az webapp config appsettings set \
  --resource-group $RG \
  --name $APP_NAME \
  --settings \
    INFOSEC_EMAIL="security@company.com" \
    SPAM_REPORT_EMAIL="spam@company.com" \
    NODE_ENV="production" \
    PORT="8080"
```

#### 3. Deploy Code

```bash
# Initialize git (if not already)
git init
git add .
git commit -m "Initial commit"

# Add Azure remote
az webapp deployment source config-local-git \
  --resource-group $RG \
  --name $APP_NAME

# Get Git URL
az webapp deployment source show \
  --resource-group $RG \
  --name $APP_NAME \
  --query url

# Push to Azure
git remote add azure <GIT_URL>
git push azure main
```

#### 4. Setup Custom Domain

```bash
az webapp config hostname add \
  --resource-group $RG \
  --webapp-name $APP_NAME \
  --hostname adgin.company.com
```

#### 5. Add SSL Certificate

```bash
az webapp config ssl bind \
  --resource-group $RG \
  --name $APP_NAME \
  --certificate-thumbprint <CERT_THUMBPRINT> \
  --ssl-type SNI
```

---

## GitHub Pages (Static Hosting)

Not suitable for Node.js backend, but useful for documentation/taskpane UI.

```bash
# Create gh-pages branch
git checkout -b gh-pages

# Push documentation
git push origin gh-pages

# Enable in GitHub Settings > Pages
```

---

## AWS Deployment

### EC2 Instance

```bash
# Launch EC2 instance (Ubuntu 20.04)
# Security group: Allow 80, 443, 22

ssh -i key.pem ubuntu@ec2-instance-ip

# Install Node.js
curl -fsSL https://deb.nodesource.com/setup_16.x | sudo -E bash -
sudo apt-get install -y nodejs

# Clone repository
git clone https://github.com/adg-tech/outlook-web-addin.git
cd outlook-web-addin

# Setup
npm install --production
npm start
```

### Elastic Beanstalk

```bash
# Install EB CLI
pip install awsebcli

# Initialize
eb init -p "Node.js 16 running on 64bit Amazon Linux 2" --region us-east-1

# Create environment
eb create production --envtype=LoadBalanced

# Deploy
eb deploy

# Monitor
eb logs
```

### Lambda with API Gateway

Not recommended for WebSocket/long-lived connections, but possible for stateless operations.

---

## Docker Deployment

### Dockerfile

```dockerfile
FROM node:16-alpine

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm ci --only=production

# Copy application
COPY . .

# Create certs directory
RUN mkdir -p certs

# Expose port
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=40s --retries=3 \
  CMD node -e "require('https').get('https://localhost:3000/health', (r) => {if (r.statusCode !== 200) throw new Error(r.statusCode)})"

# Start application
CMD ["npm", "start"]
```

### docker-compose.yml

```yaml
version: "3.8"

services:
  adgin:
    build: .
    container_name: adgin-outlook-addin
    ports:
      - "3000:3000"
    environment:
      NODE_ENV: production
      INFOSEC_EMAIL: ${INFOSEC_EMAIL}
      SPAM_REPORT_EMAIL: ${SPAM_REPORT_EMAIL}
      SUPPORT_EMAIL: ${SUPPORT_EMAIL}
    volumes:
      - ./certs:/app/certs:ro
      - ./logs:/app/logs
    restart: unless-stopped

  # Optional: Nginx reverse proxy
  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf:ro
      - ./certs:/etc/nginx/certs:ro
    depends_on:
      - adgin
    restart: unless-stopped
```

Build and run:

```bash
docker-compose up -d
```

---

## Monitoring & Logging

### PM2 (Process Manager)

```bash
npm install -g pm2

# Start with PM2
pm2 start server.js --name "adgin" --env production

# Monitor
pm2 monit

# View logs
pm2 logs adgin

# Setup auto-restart on server reboot
pm2 startup
pm2 save
```

### Logging Service

```javascript
// In server.js - add logging
const fs = require("fs");
const path = require("path");

// Ensure logs directory exists
const logsDir = path.join(__dirname, "logs");
if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir);
}

// Log all requests
app.use((req, res, next) => {
  const log = `${new Date().toISOString()} - ${req.method} ${req.path}\n`;
  fs.appendFileSync(path.join(logsDir, "access.log"), log);
  next();
});

// Log errors
app.use((err, req, res, next) => {
  const errorLog = `${new Date().toISOString()} - ERROR: ${err.message}\n${err.stack}\n`;
  fs.appendFileSync(path.join(logsDir, "error.log"), errorLog);
  res.status(500).json({ error: "Internal Server Error" });
});
```

---

## Troubleshooting

### Certificate Issues

**Problem**: "SSL certificate problem"

**Solution**:

```bash
# Verify certificate
openssl x509 -in certs/cert.pem -text -noout

# Check cert expiry
openssl x509 -in certs/cert.pem -noout -dates
```

### Port Already in Use

**Problem**: `EADDRINUSE: address already in use :::3000`

**Solution**:

```bash
# Find process using port 3000
lsof -i :3000

# Kill process
kill -9 <PID>

# Or use different port
PORT=3001 npm start
```

### Manifest Not Found

**Problem**: `404 manifest.xml not found`

**Solutions**:

1. Verify manifest.xml exists in root directory
2. Check HTTPS is enabled
3. Verify correct URL in Outlook settings

### Email Not Sending

**Problem**: Report email not reaching InfoSec

**Solutions**:

1. Verify CONFIG email addresses are correct
2. Check Exchange Online access
3. Review server logs for errors
4. Test email connectivity with external API

---

**For support, contact**: adgin-support@company.com
