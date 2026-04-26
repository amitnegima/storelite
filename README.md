# StoreLite — Vercel Deployment Guide

## Step 1: Deploy to Vercel

1. Go to https://vercel.com → Sign up / Login with GitHub
2. Click "Add New..." → "Project"
3. Select "Import Third-Party Git Repository" OR just use the upload method:
   - Go to https://vercel.com/new
   - At the bottom, click "Deploy" under "Clone Template" section
   - OR easier: Install Vercel CLI and run `vercel` in this folder

### Easiest Method (drag & drop):
1. Go to https://app.netlify.com (Netlify also works with Vercel-style deploys)
2. OR use Vercel CLI:
   ```
   npm i -g vercel
   cd deploy
   vercel
   ```
3. Follow prompts → set project name: storelite
4. Your site is live at: https://storelite.vercel.app

## Step 2: Set Up Settings Tab (REQUIRED)

Open your Master Registry Google Sheet and create a "Settings" tab:

| Key | Value |
|-----|-------|
| AdminPassword | YourStrongPassword123 |

OR for hash-based auth (more secure):

| Key | Value |
|-----|-------|
| AdminPasswordHash | (SHA-256 hash from Tools panel) |

## Step 3: Update Apps Script (if not done already)

### Registry Script (Master Registry sheet):
1. Open Master Registry → Extensions → Apps Script
2. Paste contents of `registry-apps-script.js`
3. Deploy → New deployment → Web app → Anyone
4. Copy URL → it's already set in admin.html

### Store Script (each store's sheet):
1. Open store's Google Sheet → Extensions → Apps Script
2. Paste contents of `store-apps-script.js`
3. Deploy → New deployment → Web app → Anyone
4. Add URL as `OrderScript` in that store's Config tab

## Your URLs After Deploy

| Page | URL |
|------|-----|
| Landing page | https://storelite.vercel.app |
| Demo store | https://storelite.vercel.app/?store=grocery-srinagar |
| Any store | https://storelite.vercel.app/?store=SLUG |
| Dashboard | https://storelite.vercel.app/dashboard.html?store=SLUG |
| Admin | https://storelite.vercel.app/admin.html |

## Files

| File | Size | Purpose |
|------|------|---------|
| index.html | 42K | Landing page + signup |
| store.html | 105K | Customer store |
| dashboard.html | 60K | Shopkeeper panel |
| admin.html | 49K | SaaS admin + tools |
| manifest.json | 654B | PWA config |
| sw.js | 2.3K | Service worker |
| icon-192.png | 753B | App icon |
| icon-512.png | 4.2K | App icon |
| store-apps-script.js | 7.9K | Per-store Apps Script |
| registry-apps-script.js | 17K | Master registry Apps Script |
