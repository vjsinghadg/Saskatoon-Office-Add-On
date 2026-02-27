#!/bin/bash

# ADGSentinel Quick Start Setup Script
# This script sets up the development environment for the Outlook Web Add-in

set -e

echo "🚀 ADGSentinel Office.js Add-in Setup"
echo "======================================"
echo ""

# Check Node.js installation
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js 12+ first."
    exit 1
fi

echo "✅ Node.js $(node --version) detected"

# Check npm installation
if ! command -v npm &> /dev/null; then
    echo "❌ npm is not installed. Please install npm first."
    exit 1
fi

echo "✅ npm $(npm --version) detected"

# Create necessary directories
echo ""
echo "📁 Creating directory structure..."
mkdir -p public/assets

echo "✅ Directories created"

# Install dependencies
echo ""
echo "📦 Installing npm dependencies..."
npm install

echo "✅ Dependencies installed"

# Generate SSL certificates
echo ""
echo "🔐 Generating SSL certificates for HTTPS..."

if [ ! -d "certs" ]; then
    mkdir -p certs
    
    # Generate self-signed certificate (valid for 365 days)
    openssl req -x509 -newkey rsa:2048 -keyout certs/key.pem -out certs/cert.pem \
        -days 365 -nodes \
        -subj "/C=US/ST=State/L=City/O=Organization/CN=localhost"
    
    echo "✅ SSL certificates generated in ./certs/"
else
    echo "✅ Certificates directory exists"
fi

# Create .env file if it doesn't exist
echo ""
echo "⚙️  Checking configuration..."

if [ ! -f ".env" ]; then
    echo "Creating .env file from template..."
    cp .env.example .env
    echo "⚠️  Please update .env with your email configuration:"
    echo "   - INFOSEC_EMAIL"
    echo "   - SPAM_REPORT_EMAIL"
    echo "   - SUPPORT_EMAIL"
    echo "   - GOPHISH_URL (optional)"
else
    echo "✅ .env file already exists"
fi

# Summary
echo ""
echo "======================================"
echo "✅ Setup Complete!"
echo ""
echo "Next steps:"
echo "1. Update your email addresses in .env file"
echo "2. Update configuration in function-file.js (CONFIG object)"
echo "3. Create icons in public/assets/ directory (optional)"
echo "4. Run 'npm start' to start the development server"
echo ""
echo "Server will be available at: https://localhost:3000"
echo "Manifest URL: https://localhost:3000/manifest.xml"
echo ""
echo "To upload to Outlook:"
echo "1. Open Outlook Web Settings > Get Add-ins"
echo "2. Click 'My Add-ins' > 'Upload My Add-in'"
echo "3. Choose 'Upload from URL' or upload manifest.xml directly"
echo ""
echo "======================================"
