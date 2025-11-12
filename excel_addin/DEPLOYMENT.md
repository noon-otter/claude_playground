# Deployment Guide

How to deploy the Domino Excel Governance Add-in to production.

## Overview

The add-in consists of:
1. **Static files** (HTML/JS/CSS) - hosted on web server
2. **Manifest file** - distributed to users
3. **Domino API** - backend service (separate deployment)

## Prerequisites

- Web hosting (Azure, AWS, etc.)
- HTTPS domain (required by Office)
- Domino API endpoint

## Build for Production

### 1. Update Configuration

**Update API endpoints:**

Edit `src/commands/commands.js`:
```javascript
const DOMINO_API_BASE = 'https://domino.your-company.com/api';
```

Edit `src/utils/domino-api.js`:
```javascript
const DOMINO_API_BASE = 'https://domino.your-company.com/api';
```

Or use environment variables:

```bash
# .env.production
VITE_DOMINO_API_URL=https://domino.your-company.com/api
```

**Update manifest.xml:**

Replace all `localhost:3000` with your production domain:

```xml
<bt:Url id="CommandsFile.Url" DefaultValue="https://excel-addin.your-company.com/commands.html"/>
<bt:Url id="Taskpane.Url" DefaultValue="https://excel-addin.your-company.com/index.html"/>
<bt:Url id="RegisterModal.Url" DefaultValue="https://excel-addin.your-company.com/register.html"/>
```

Update AppDomains:
```xml
<AppDomains>
  <AppDomain>https://excel-addin.your-company.com</AppDomain>
</AppDomains>
```

Update icon URLs:
```xml
<bt:Image id="Icon.16x16" DefaultValue="https://excel-addin.your-company.com/assets/icon-16.png"/>
<bt:Image id="Icon.32x32" DefaultValue="https://excel-addin.your-company.com/assets/icon-32.png"/>
<bt:Image id="Icon.80x80" DefaultValue="https://excel-addin.your-company.com/assets/icon-80.png"/>
```

**Generate unique App ID:**

In `manifest.xml`, replace the placeholder ID:
```xml
<Id>12345678-1234-1234-1234-123456789012</Id>
```

With a real GUID:
```bash
# Generate on Mac/Linux
uuidgen

# Or use online generator
# https://www.uuidgenerator.net/
```

### 2. Build

```bash
npm run build
```

Output goes to `dist/` folder:
```
dist/
├── index.html
├── commands.html
├── register.html
├── assets/
│   ├── index-[hash].js
│   ├── commands-[hash].js
│   ├── register-[hash].js
│   └── icons...
```

### 3. Test Build Locally

```bash
npm run preview
```

Visit `http://localhost:4173` to verify build works.

## Deployment Options

### Option 1: Azure Static Web Apps (Recommended)

**Pros:**
- Free tier available
- Auto HTTPS
- CDN included
- Easy CI/CD with GitHub

**Steps:**

1. Create Azure Static Web App
```bash
az login
az staticwebapp create \
  --name domino-excel-addin \
  --resource-group your-rg \
  --location eastus2
```

2. Deploy
```bash
# Using Azure CLI
az staticwebapp deploy \
  --name domino-excel-addin \
  --resource-group your-rg \
  --source ./dist
```

Or connect to GitHub for auto-deploy.

3. Get URL
```
https://domino-excel-addin.azurestaticapps.net
```

### Option 2: AWS S3 + CloudFront

**Steps:**

1. Create S3 bucket
```bash
aws s3 mb s3://domino-excel-addin
```

2. Upload files
```bash
aws s3 sync ./dist s3://domino-excel-addin --acl public-read
```

3. Configure CloudFront
- Create distribution
- Point to S3 bucket
- Enable HTTPS
- Set custom domain (optional)

4. Get URL
```
https://d1234567890.cloudfront.net
```

### Option 3: Netlify

**Steps:**

1. Install Netlify CLI
```bash
npm install -g netlify-cli
```

2. Deploy
```bash
netlify deploy --prod --dir=dist
```

3. Get URL
```
https://domino-excel-addin.netlify.app
```

### Option 4: Internal Server (IIS/Apache/Nginx)

**Requirements:**
- HTTPS certificate
- Static file hosting
- CORS headers configured

**Nginx config example:**
```nginx
server {
    listen 443 ssl;
    server_name excel-addin.your-company.com;

    ssl_certificate /path/to/cert.pem;
    ssl_certificate_key /path/to/key.pem;

    root /var/www/excel-addin;
    index index.html;

    # CORS headers
    add_header Access-Control-Allow-Origin *;
    add_header Access-Control-Allow-Methods "GET, POST, OPTIONS";

    location / {
        try_files $uri $uri/ =404;
    }
}
```

Upload `dist/` contents to `/var/www/excel-addin`.

## Distribute to Users

### Option A: Centralized Deployment (Enterprise)

**For IT Admins:**

1. Go to Microsoft 365 Admin Center
2. Settings → Integrated apps → Upload custom apps
3. Upload `manifest.xml`
4. Choose deployment:
   - Entire organization
   - Specific groups/users
5. Deploy

Users will see add-in automatically in Excel.

### Option B: Self-Service (Individuals/Teams)

**For End Users:**

1. Host `manifest.xml` on accessible URL
   - SharePoint
   - Internal file server
   - Azure Blob Storage

2. Send instructions:
   ```
   1. Download manifest.xml from: [URL]
   2. Open Excel
   3. Insert → Get Add-ins → My Add-ins
   4. Upload My Add-in
   5. Browse to downloaded manifest.xml
   6. Click Upload
   ```

### Option C: AppSource (Public)

To publish on Microsoft AppSource:

1. Register as Microsoft Partner
2. Submit add-in for validation
3. Wait for approval (~1-2 weeks)
4. Users find in Office Store

**Not recommended for internal governance tool.**

## Configure CORS on Domino API

The Excel add-in runs in browser, so Domino API needs CORS headers.

**FastAPI example:**
```python
from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://excel-addin.your-company.com"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

**Express example:**
```javascript
const cors = require('cors');

app.use(cors({
    origin: 'https://excel-addin.your-company.com',
    credentials: true
}));
```

## SSL Certificate

Office Add-ins **require HTTPS** in production.

Options:
- Let's Encrypt (free, auto-renew)
- Azure/AWS managed certificates
- Corporate certificate authority

## Monitoring

### Add-in Usage
- Monitor API calls to Domino
- Track registration events
- Count active users

### Error Tracking
Consider adding:
- Sentry.io for JavaScript errors
- Application Insights (Azure)
- CloudWatch (AWS)

Example with Sentry:
```javascript
// src/main.jsx
import * as Sentry from "@sentry/react";

Sentry.init({
  dsn: "https://your-sentry-dsn",
  environment: "production"
});
```

## Versioning

Update version in `manifest.xml`:
```xml
<Version>1.1.0.0</Version>
```

Office caches manifest - users need to:
1. Remove old add-in
2. Re-upload new manifest

Or use centralized deployment to push updates.

## Rollback

If deployment fails:

1. Revert to previous `dist/` build
2. Re-deploy
3. Or keep old version URL live:
   - `v1.domino-excel-addin.com`
   - `v2.domino-excel-addin.com`

Update manifest to point to stable version.

## Security Checklist

- [ ] HTTPS enabled
- [ ] CORS restricted to add-in domain only
- [ ] No sensitive data in client-side code
- [ ] API authentication enabled (OAuth/JWT)
- [ ] Content Security Policy headers
- [ ] Manifest signed (optional)
- [ ] Regular dependency updates

## Performance

Office Add-ins should be fast:

- Minimize bundle size (code splitting)
- Use CDN for assets
- Compress images
- Enable gzip/brotli compression
- Cache static assets

Check bundle size:
```bash
npm run build
du -sh dist/
```

## Testing in Production

Before full rollout:

1. Deploy to staging environment
2. Test with pilot group (5-10 users)
3. Monitor for errors
4. Gather feedback
5. Fix issues
6. Roll out to everyone

## Support

Document for users:
- How to install
- How to use
- Troubleshooting guide
- Contact for help

Example support doc:
```markdown
# Excel Governance Add-in Support

## Installation Issues
If add-in won't load, try:
1. Clear Excel cache
2. Restart Excel
3. Re-upload manifest

## Questions
Email: excel-support@company.com
Slack: #excel-governance
```

## Compliance

For regulated industries:

- Log all deployments
- Track which users have add-in
- Document data retention policies
- Audit event storage
- Ensure GDPR/SOC2 compliance

## Updates

When updating the add-in:

1. Update version in `manifest.xml`
2. Rebuild: `npm run build`
3. Deploy new `dist/` files
4. Distribute new manifest (if changed)
5. Notify users of changes

Consider backward compatibility for events/API.
