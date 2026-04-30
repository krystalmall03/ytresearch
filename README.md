# YT Research Tool — by Krystal Research

A free YouTube channel research tool that pulls full channel data and exports to Excel, PDF, or Word.

## Setup (5 minutes)

### 1. Get a YouTube API Key (Free)
1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project (or use existing)
3. Go to **APIs & Services → Enable APIs**
4. Search for **"YouTube Data API v3"** and enable it
5. Go to **APIs & Services → Credentials**
6. Click **Create Credentials → API Key**
7. Copy your key

### 2. Configure
```bash
cp .env.example .env
# Open .env and paste your API key
```

### 3. Install & Run
```bash
npm install
npm start
```

Open http://localhost:3000 in your browser.

## Features
- ✅ Analyze any YouTube channel by URL, handle, or channel ID
- ✅ Subscribers, views, video count, upload frequency
- ✅ Average views/likes/comments (last 20 videos)
- ✅ Estimated lifetime revenue range
- ✅ Full recent video table with thumbnails
- ✅ Export to **Excel (.xlsx)**, **Word (.docx)**, **PDF**
- ✅ 100% free, no login required

## Deploy to the web

**Render.com (free tier):**
1. Push to GitHub
2. Connect repo on render.com
3. Add `YOUTUBE_API_KEY` as environment variable
4. Deploy

**Railway.app:**
Same steps — connect repo, set env var, deploy.

## Ko-fi / Tips
This tool is free forever. If it helps you, consider leaving a tip:
https://ko-fi.com

---
Powered by [Krystal Research](https://krystalresearch.org) · Tamale, Ghana
