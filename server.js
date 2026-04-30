require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } = require('docx');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'frontend/public')));

const YT_API_KEY = process.env.YOUTUBE_API_KEY;
const YT_BASE = 'https://www.googleapis.com/youtube/v3';

// ─── Category map & RPM data (verified 2025-2026) ─────────

const CATEGORIES = {
  '1':'Film & Animation','2':'Autos & Vehicles','10':'Music','15':'Pets & Animals',
  '17':'Sports','18':'Short Movies','19':'Travel & Events','20':'Gaming',
  '21':'Videoblogging','22':'People & Blogs','23':'Comedy','24':'Entertainment',
  '25':'News & Politics','26':'Howto & Style','27':'Education','28':'Science & Technology',
  '29':'Nonprofits & Activism','30':'Movies','43':'Shows','44':'Trailers'
};

// RPM ranges by category (USD, 2025-2026 verified data)
// Sources: OutlierKit, MilX, FluxNote, Lenos creator dashboards
const RPM_BY_CATEGORY = {
  '1':  { low: 2, high: 6, label: 'Film & Animation' },
  '2':  { low: 4, high: 10, label: 'Autos & Vehicles' },
  '10': { low: 1, high: 3, label: 'Music' },
  '15': { low: 2, high: 5, label: 'Pets & Animals' },
  '17': { low: 2, high: 6, label: 'Sports' },
  '19': { low: 3, high: 8, label: 'Travel & Events' },
  '20': { low: 2, high: 5, label: 'Gaming' },
  '22': { low: 2, high: 6, label: 'People & Blogs' },
  '23': { low: 2, high: 5, label: 'Comedy' },
  '24': { low: 2, high: 5, label: 'Entertainment' },
  '25': { low: 5, high: 12, label: 'News & Politics' },
  '26': { low: 3, high: 8, label: 'Howto & Style' },
  '27': { low: 5, high: 12, label: 'Education' },
  '28': { low: 5, high: 12, label: 'Science & Technology' },
  '29': { low: 1, high: 4, label: 'Nonprofits & Activism' },
  'default': { low: 2, high: 6, label: 'General' },
  // Special sub-niche overrides (applied if keywords match)
  'finance': { low: 10, high: 25, label: 'Personal Finance' },
  'investing': { low: 12, high: 30, label: 'Investing' },
  'real_estate': { low: 8, high: 18, label: 'Real Estate' },
  'insurance': { low: 15, high: 35, label: 'Insurance' },
  'legal': { low: 10, high: 20, label: 'Legal' },
  'saas': { low: 8, high: 15, label: 'SaaS/Software' },
  'health': { low: 5, high: 12, label: 'Health & Wellness' },
  'asmr': { low: 5, high: 11, label: 'ASMR/Ambient' },
  'true_crime': { low: 8, high: 12, label: 'True Crime' },
};

function detectSubNiche(title, description, tags) {
  const text = `${title} ${description} ${(tags||[]).join(' ')}`.toLowerCase();
  if (/\b(invest|stock|crypto|trading|dividend|portfolio|etf|401k)\b/.test(text)) return 'investing';
  if (/\b(financ|money|budget|credit card|debt|savings|wealth|income|earn)\b/.test(text)) return 'finance';
  if (/\b(real estate|property|mortgage|housing|rent|landlord)\b/.test(text)) return 'real_estate';
  if (/\b(insurance|coverage|premium|deductible|liability)\b/.test(text)) return 'insurance';
  if (/\b(lawyer|attorney|legal|court|lawsuit|law firm)\b/.test(text)) return 'legal';
  if (/\b(saas|software|app review|tutorial|tool|plugin)\b/.test(text)) return 'saas';
  if (/\b(asmr|ambient|sleep|meditation|relaxation|soak)\b/.test(text)) return 'asmr';
  if (/\b(true crime|murder|mystery|investigation|cold case|forensic)\b/.test(text)) return 'true_crime';
  if (/\b(health|medical|doctor|nurse|wellness|nutrition|diet|fitness)\b/.test(text)) return 'health';
  return null;
}

function estimateRevenue(views, categoryId, title, description, tags) {
  views = parseInt(views) || 0;
  const subNiche = detectSubNiche(title||'', description||'', tags);
  const rpm = subNiche ? RPM_BY_CATEGORY[subNiche] :
              (RPM_BY_CATEGORY[categoryId] || RPM_BY_CATEGORY['default']);
  return {
    low: `$${Math.round(views/1000*rpm.low).toLocaleString()}`,
    high: `$${Math.round(views/1000*rpm.high).toLocaleString()}`,
    rpmLow: rpm.low,
    rpmHigh: rpm.high,
    niche: rpm.label,
    note: `Based on ${rpm.label} niche RPM ($${rpm.low}–$${rpm.high}/1K views). Actual revenue varies by audience geography, watch time, and monetization status.`
  };
}

// ─── URL Parsing ───────────────────────────────────────────

function parseInput(input) {
  input = input.trim();
  if (/^UC[\w-]{22}$/.test(input)) return { type: 'channelId', value: input };
  const videoMatch = input.match(/(?:youtube\.com\/watch\?.*v=|youtu\.be\/|youtube\.com\/shorts\/)([\w-]{11})/);
  if (videoMatch) return { type: 'videoId', value: videoMatch[1] };
  const handleMatch = input.match(/(?:youtube\.com\/@|^@)([\w.-]+)/);
  if (handleMatch) return { type: 'handle', value: handleMatch[1] };
  const channelMatch = input.match(/youtube\.com\/channel\/(UC[\w-]{22})/);
  if (channelMatch) return { type: 'channelId', value: channelMatch[1] };
  const customMatch = input.match(/youtube\.com\/(?:c\/|user\/)([\w.-]+)/);
  if (customMatch) return { type: 'handle', value: customMatch[1] };
  if (/^[\w.-]+$/.test(input)) return { type: 'handle', value: input };
  return null;
}

async function resolveChannelId(parsed) {
  if (parsed.type === 'channelId') return parsed.value;
  if (parsed.type === 'videoId') {
    const res = await axios.get(`${YT_BASE}/videos`, { params: { part: 'snippet', id: parsed.value, key: YT_API_KEY } });
    if (!res.data.items?.length) throw new Error('Video not found');
    return res.data.items[0].snippet.channelId;
  }
  try {
    const r = await axios.get(`${YT_BASE}/channels`, { params: { part: 'id', forHandle: parsed.value, key: YT_API_KEY } });
    if (r.data.items?.length) return r.data.items[0].id;
  } catch (e) {}
  const r = await axios.get(`${YT_BASE}/search`, { params: { part: 'snippet', q: parsed.value, type: 'channel', maxResults: 1, key: YT_API_KEY } });
  if (r.data.items?.length) return r.data.items[0].snippet.channelId;
  throw new Error('Channel not found');
}

// ─── Helpers ───────────────────────────────────────────────

function fmt(n) {
  n = parseInt(n) || 0;
  if (n >= 1e9) return (n/1e9).toFixed(2)+'B';
  if (n >= 1e6) return (n/1e6).toFixed(2)+'M';
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K';
  return n.toString();
}

function parseDuration(iso) {
  if (!iso) return '0:00';
  const m = iso.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!m) return '0:00';
  const h=parseInt(m[1]||0), mn=parseInt(m[2]||0), s=parseInt(m[3]||0);
  return h>0 ? `${h}:${String(mn).padStart(2,'0')}:${String(s).padStart(2,'0')}` : `${mn}:${String(s).padStart(2,'0')}`;
}

function durationSeconds(iso) {
  if (!iso) return 0;
  const m = iso.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!m) return 0;
  return (parseInt(m[1]||0)*3600)+(parseInt(m[2]||0)*60)+parseInt(m[3]||0);
}

function mapVideoData(v) {
  const catId = v.snippet.categoryId;
  const views = parseInt(v.statistics?.viewCount)||0;
  const likes = parseInt(v.statistics?.likeCount)||0;
  const comments = parseInt(v.statistics?.commentCount)||0;
  return {
    id: v.id,
    title: v.snippet.title,
    description: v.snippet.description,
    publishedAt: v.snippet.publishedAt,
    thumbnail: v.snippet.thumbnails?.medium?.url || v.snippet.thumbnails?.default?.url,
    thumbnailHigh: v.snippet.thumbnails?.maxres?.url || v.snippet.thumbnails?.high?.url || v.snippet.thumbnails?.default?.url,
    categoryId: catId,
    category: CATEGORIES[catId] || 'Unknown',
    tags: v.snippet.tags || [],
    defaultLanguage: v.snippet.defaultLanguage || v.snippet.defaultAudioLanguage || 'N/A',
    duration: parseDuration(v.contentDetails?.duration),
    durationSeconds: durationSeconds(v.contentDetails?.duration),
    definition: (v.contentDetails?.definition || 'sd').toUpperCase(),
    caption: v.contentDetails?.caption === 'true' ? 'Yes' : 'No',
    licensedContent: v.contentDetails?.licensedContent ? 'Yes' : 'No',
    projection: v.contentDetails?.projection || 'rectangular',
    views, likes, comments, commentCount: comments,
    likeRate: views > 0 ? ((likes/views)*100).toFixed(2)+'%' : '0%',
    commentRate: views > 0 ? ((comments/views)*100).toFixed(3)+'%' : '0%',
    engagementRate: views > 0 ? (((likes+comments)/views)*100).toFixed(2)+'%' : '0%',
    url: `https://youtube.com/watch?v=${v.id}`
  };
}

// ─── Data Fetchers ─────────────────────────────────────────

async function getChannelData(channelId) {
  const res = await axios.get(`${YT_BASE}/channels`, { params: { part: 'snippet,statistics,brandingSettings,contentDetails,topicDetails', id: channelId, key: YT_API_KEY } });
  if (!res.data.items?.length) throw new Error('Channel not found');
  return res.data.items[0];
}

async function getRecentVideos(channelId, maxResults = 20) {
  const ch = await axios.get(`${YT_BASE}/channels`, { params: { part: 'contentDetails', id: channelId, key: YT_API_KEY } });
  const uploadsId = ch.data.items[0].contentDetails.relatedPlaylists.uploads;
  const pl = await axios.get(`${YT_BASE}/playlistItems`, { params: { part: 'contentDetails', playlistId: uploadsId, maxResults, key: YT_API_KEY } });
  const ids = pl.data.items.map(i => i.contentDetails.videoId);
  const vids = await axios.get(`${YT_BASE}/videos`, { params: { part: 'snippet,statistics,contentDetails', id: ids.join(','), key: YT_API_KEY } });
  return vids.data.items;
}

async function getVideoData(videoId) {
  const res = await axios.get(`${YT_BASE}/videos`, { params: { part: 'snippet,statistics,contentDetails,topicDetails', id: videoId, key: YT_API_KEY } });
  if (!res.data.items?.length) throw new Error('Video not found');
  return res.data.items[0];
}

async function getVideoComments(videoId, maxResults = 15) {
  try {
    const res = await axios.get(`${YT_BASE}/commentThreads`, { params: { part: 'snippet', videoId, maxResults, order: 'relevance', key: YT_API_KEY } });
    return res.data.items || [];
  } catch (e) { return []; }
}

// ─── Channel Analysis ──────────────────────────────────────

app.get('/api/analyze/channel', async (req, res) => {
  try {
    const { url } = req.query;
    if (!url) return res.status(400).json({ error: 'URL required' });
    if (!YT_API_KEY) return res.status(500).json({ error: 'YouTube API key not configured' });
    const parsed = parseInput(url);
    if (!parsed) return res.status(400).json({ error: 'Could not parse that URL' });

    const channelId = await resolveChannelId(parsed);
    const [channelData, rawVideos] = await Promise.all([getChannelData(channelId), getRecentVideos(channelId, 20)]);
    const stats = channelData.statistics, snippet = channelData.snippet;
    const videos = rawVideos.map(mapVideoData);

    const totalViews = videos.reduce((s,v) => s+v.views, 0);
    const totalLikes = videos.reduce((s,v) => s+v.likes, 0);
    const totalComments = videos.reduce((s,v) => s+v.comments, 0);
    const len = videos.length || 1;

    const dates = videos.map(v => new Date(v.publishedAt)).sort((a,b) => b-a);
    let uploadFreq = 'N/A';
    if (dates.length >= 2) {
      const dpv = (dates[0]-dates[dates.length-1])/(864e5)/(dates.length-1);
      uploadFreq = dpv<1?'Multiple/day':dpv<2?'Daily':dpv<4?'Every 2-3 days':dpv<8?'Weekly':dpv<15?'Bi-weekly':'Monthly or less';
    }

    // Use most common category for channel RPM
    const catCounts = {};
    videos.forEach(v => { catCounts[v.categoryId] = (catCounts[v.categoryId]||0)+1; });
    const primaryCat = Object.entries(catCounts).sort((a,b)=>b[1]-a[1])[0]?.[0] || 'default';

    const revenue = estimateRevenue(stats.viewCount, primaryCat, snippet.title, snippet.description);

    res.json({
      mode:'channel', channelId, name:snippet.title, handle:snippet.customUrl||'',
      description:snippet.description, country:snippet.country||'N/A', createdAt:snippet.publishedAt,
      thumbnail:snippet.thumbnails?.high?.url||snippet.thumbnails?.default?.url,
      subscribers:parseInt(stats.subscriberCount)||0, subscribersFormatted:fmt(stats.subscriberCount),
      totalViews:parseInt(stats.viewCount)||0, totalViewsFormatted:fmt(stats.viewCount),
      videoCount:parseInt(stats.videoCount)||0, hiddenSubscribers:stats.hiddenSubscriberCount||false,
      avgViews:Math.round(totalViews/len), avgLikes:Math.round(totalLikes/len), avgComments:Math.round(totalComments/len),
      uploadFrequency:uploadFreq, estimatedRevenueTotal:revenue,
      primaryCategory: CATEGORIES[primaryCat] || 'General',
      topics:channelData.topicDetails?.topicCategories?.map(t=>t.split('/').pop().replace(/_/g,' '))||[],
      keywords:channelData.brandingSettings?.channel?.keywords||'',
      recentVideos: videos
    });
  } catch (err) {
    const msg = err.response?.data?.error?.message||err.message;
    console.error('Channel Error:', msg);
    res.status(500).json({ error: msg });
  }
});

// ─── Video Analysis ────────────────────────────────────────

app.get('/api/analyze/video', async (req, res) => {
  try {
    const { url } = req.query;
    if (!url) return res.status(400).json({ error: 'URL required' });
    if (!YT_API_KEY) return res.status(500).json({ error: 'YouTube API key not configured' });
    const parsed = parseInput(url);
    if (!parsed) return res.status(400).json({ error: 'Could not parse that URL' });
    if (parsed.type !== 'videoId') return res.status(400).json({ error: 'Please paste a video URL for video analysis' });

    const rawVideo = await getVideoData(parsed.value);
    const video = mapVideoData(rawVideo);
    const channel = await getChannelData(rawVideo.snippet.channelId);
    const chStats = channel.statistics;
    const commentThreads = await getVideoComments(parsed.value, 15);

    const revenue = estimateRevenue(video.views, video.categoryId, video.title, video.description, video.tags);

    res.json({
      mode:'video', ...video,
      estimatedRevenue: revenue,
      channel: {
        id: rawVideo.snippet.channelId, name: channel.snippet.title,
        handle: channel.snippet.customUrl||'',
        subscribers: fmt(chStats.subscriberCount), totalViews: fmt(chStats.viewCount),
        videoCount: fmt(chStats.videoCount),
        thumbnail: channel.snippet.thumbnails?.default?.url,
        country: channel.snippet.country || 'N/A',
        createdAt: channel.snippet.publishedAt
      },
      topComments: commentThreads.map(c => {
        const s = c.snippet.topLevelComment.snippet;
        return { author:s.authorDisplayName, text:s.textDisplay, likes:parseInt(s.likeCount)||0, publishedAt:s.publishedAt };
      })
    });
  } catch (err) {
    const msg = err.response?.data?.error?.message||err.message;
    console.error('Video Error:', msg);
    res.status(500).json({ error: msg });
  }
});

// ─── Excel Export ──────────────────────────────────────────

app.post('/api/export/excel', async (req, res) => {
  try {
    const data = req.body;
    const wb = new ExcelJS.Workbook(); wb.creator = 'YT Research';
    const hs = { font:{bold:true,color:{argb:'FFFFFFFF'}}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF1A1A2E'}} };
    const ls = { font:{bold:true}, fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FFE8E8F0'}} };
    const addRows = (ws, rows) => rows.forEach(r => {
      const row = ws.addRow(r);
      if (r[1]===''&&r[0]) { row.getCell(1).style=hs; row.getCell(2).style=hs; }
      else if (r[0]&&r[1]!==undefined) row.getCell(1).style=ls;
    });

    if (data.mode==='video') {
      const ws = wb.addWorksheet('Video Analysis');
      ws.columns = [{key:'m',width:30},{key:'v',width:60}];
      ws.addRow(['VIDEO ANALYSIS REPORT']).font={bold:true,size:16};
      ws.addRow([`Generated: ${new Date().toLocaleDateString()}`]).font={italic:true,color:{argb:'FF666666'}};
      ws.addRow([]);
      addRows(ws, [
        ['VIDEO INFO',''],['Title',data.title],['Published',new Date(data.publishedAt).toLocaleDateString()],
        ['Duration',data.duration],['Category',data.category],['Definition',data.definition],
        ['Captions',data.caption],['Language',data.defaultLanguage],['Licensed',data.licensedContent],
        ['Tags',(data.tags||[]).slice(0,20).join(', ')],['URL',data.url],
        [''],['CHANNEL',''],['Channel',data.channel.name],['Handle',data.channel.handle],
        ['Subscribers',data.channel.subscribers],['Total Views',data.channel.totalViews],
        ['Videos',data.channel.videoCount],['Country',data.channel.country],
        [''],['PERFORMANCE',''],['Views',data.views.toLocaleString()],['Likes',data.likes.toLocaleString()],
        ['Comments',data.commentCount.toLocaleString()],['Like Rate',data.likeRate],
        ['Comment Rate',data.commentRate],['Engagement Rate',data.engagementRate],
        [''],['EST. REVENUE',''],['Low',data.estimatedRevenue.low],['High',data.estimatedRevenue.high],
        ['RPM Range',`$${data.estimatedRevenue.rpmLow}–$${data.estimatedRevenue.rpmHigh}`],
        ['Niche',data.estimatedRevenue.niche],['Note',data.estimatedRevenue.note],
      ]);
      if (data.topComments?.length) {
        const cs = wb.addWorksheet('Top Comments');
        cs.columns=[{header:'Author',key:'a',width:20},{header:'Comment',key:'t',width:60},{header:'Likes',key:'l',width:10},{header:'Date',key:'d',width:14}];
        cs.getRow(1).eachCell(c=>c.style=hs);
        data.topComments.forEach(c=>cs.addRow({a:c.author,t:c.text.replace(/<[^>]*>/g,''),l:c.likes,d:new Date(c.publishedAt).toLocaleDateString()}));
      }
    } else {
      const ws = wb.addWorksheet('Channel Overview');
      ws.columns=[{key:'m',width:30},{key:'v',width:50}];
      ws.addRow(['CHANNEL OVERVIEW REPORT']).font={bold:true,size:16};
      ws.addRow([`Generated: ${new Date().toLocaleDateString()}`]).font={italic:true,color:{argb:'FF666666'}};
      ws.addRow([]);
      addRows(ws, [
        ['CHANNEL INFO',''],['Name',data.name],['Handle',data.handle],['Country',data.country],
        ['Created',new Date(data.createdAt).toLocaleDateString()],['Topics',data.topics.join(', ')],
        ['Primary Category',data.primaryCategory],['Keywords',data.keywords],
        [''],['STATISTICS',''],['Subscribers',data.subscribersFormatted],['Total Views',data.totalViewsFormatted],
        ['Videos',data.videoCount],[''],['PERFORMANCE (Last 20)',''],
        ['Avg Views',data.avgViews.toLocaleString()],['Avg Likes',data.avgLikes.toLocaleString()],
        ['Avg Comments',data.avgComments.toLocaleString()],['Frequency',data.uploadFrequency],
        [''],['EST. REVENUE',''],['Low',data.estimatedRevenueTotal.low],['High',data.estimatedRevenueTotal.high],
        ['RPM Range',`$${data.estimatedRevenueTotal.rpmLow}–$${data.estimatedRevenueTotal.rpmHigh}`],
        ['Niche',data.estimatedRevenueTotal.niche],['Note',data.estimatedRevenueTotal.note],
      ]);
      const vs = wb.addWorksheet('Recent Videos');
      vs.columns=[{header:'Title',key:'t',width:50},{header:'Date',key:'d',width:14},{header:'Category',key:'cat',width:20},
        {header:'Views',key:'v',width:12},{header:'Likes',key:'l',width:12},{header:'Comments',key:'c',width:12},
        {header:'Like Rate',key:'lr',width:10},{header:'Engagement',key:'er',width:10},
        {header:'Duration',key:'dr',width:10},{header:'Definition',key:'def',width:8},
        {header:'Captions',key:'cap',width:8},{header:'Tags',key:'tags',width:40},{header:'URL',key:'u',width:45}];
      vs.getRow(1).eachCell(c=>c.style=hs);
      data.recentVideos.forEach(v=>vs.addRow({
        t:v.title,d:new Date(v.publishedAt).toLocaleDateString(),cat:v.category,
        v:v.views,l:v.likes,c:v.comments,lr:v.likeRate,er:v.engagementRate,
        dr:v.duration,def:v.definition,cap:v.caption,tags:(v.tags||[]).slice(0,10).join(', '),u:v.url
      }));
    }
    const buf = await wb.xlsx.writeBuffer();
    const name = (data.mode==='video'?data.title:data.name).replace(/[^a-z0-9]/gi,'_');
    res.set({'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','Content-Disposition':`attachment; filename="${name}_Report.xlsx"`});
    res.send(buf);
  } catch (err) { res.status(500).json({error:err.message}); }
});

// ─── Word Export ───────────────────────────────────────────

app.post('/api/export/word', async (req, res) => {
  try {
    const data = req.body;
    const mk = (l,v) => new TableRow({children:[
      new TableCell({children:[new Paragraph({children:[new TextRun({text:l,bold:true})]})],width:{size:35,type:WidthType.PERCENTAGE}}),
      new TableCell({children:[new Paragraph({text:String(v)})],width:{size:65,type:WidthType.PERCENTAGE}}),
    ]});
    let children;
    if (data.mode==='video') {
      children = [
        new Paragraph({text:'YouTube Video Analysis Report',heading:HeadingLevel.TITLE}),
        new Paragraph({text:`Generated ${new Date().toLocaleDateString()}`}), new Paragraph({text:''}),
        new Paragraph({text:'Video Details',heading:HeadingLevel.HEADING_1}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
          mk('Title',data.title),mk('Published',new Date(data.publishedAt).toLocaleDateString()),
          mk('Duration',data.duration),mk('Category',data.category),mk('Definition',data.definition),
          mk('Captions',data.caption),mk('Language',data.defaultLanguage),mk('Licensed Content',data.licensedContent),
          mk('Channel',data.channel.name),mk('Channel Subs',data.channel.subscribers),
        ]}), new Paragraph({text:''}),
        new Paragraph({text:'Performance',heading:HeadingLevel.HEADING_1}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
          mk('Views',data.views.toLocaleString()),mk('Likes',data.likes.toLocaleString()),
          mk('Comments',data.commentCount.toLocaleString()),mk('Like Rate',data.likeRate),
          mk('Comment Rate',data.commentRate),mk('Engagement Rate',data.engagementRate),
        ]}), new Paragraph({text:''}),
        new Paragraph({text:'Revenue Estimate',heading:HeadingLevel.HEADING_1}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
          mk('Low',data.estimatedRevenue.low),mk('High',data.estimatedRevenue.high),
          mk('RPM Range',`$${data.estimatedRevenue.rpmLow}–$${data.estimatedRevenue.rpmHigh}`),
          mk('Detected Niche',data.estimatedRevenue.niche),mk('Note',data.estimatedRevenue.note),
        ]}),
      ];
      if (data.tags?.length) {
        children.push(new Paragraph({text:''}), new Paragraph({text:'Tags',heading:HeadingLevel.HEADING_1}),
          new Paragraph({text:data.tags.join(', ')}));
      }
    } else {
      children = [
        new Paragraph({text:'YouTube Channel Research Report',heading:HeadingLevel.TITLE}),
        new Paragraph({text:`Generated ${new Date().toLocaleDateString()}`}), new Paragraph({text:''}),
        new Paragraph({text:'Channel Overview',heading:HeadingLevel.HEADING_1}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
          mk('Channel',data.name),mk('Handle',data.handle),mk('Country',data.country),
          mk('Created',new Date(data.createdAt).toLocaleDateString()),mk('Primary Category',data.primaryCategory),
        ]}), new Paragraph({text:''}),
        new Paragraph({text:'Statistics',heading:HeadingLevel.HEADING_1}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
          mk('Subscribers',data.subscribersFormatted),mk('Total Views',data.totalViewsFormatted),
          mk('Videos',data.videoCount),mk('Frequency',data.uploadFrequency),
          mk('Avg Views',data.avgViews.toLocaleString()),mk('Avg Likes',data.avgLikes.toLocaleString()),
        ]}), new Paragraph({text:''}),
        new Paragraph({text:'Revenue Estimate (Lifetime)',heading:HeadingLevel.HEADING_1}),
        new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
          mk('Low',data.estimatedRevenueTotal.low),mk('High',data.estimatedRevenueTotal.high),
          mk('RPM Range',`$${data.estimatedRevenueTotal.rpmLow}–$${data.estimatedRevenueTotal.rpmHigh}`),
          mk('Detected Niche',data.estimatedRevenueTotal.niche),mk('Note',data.estimatedRevenueTotal.note),
        ]}),
      ];
    }
    const doc = new Document({sections:[{children}]});
    const buf = await Packer.toBuffer(doc);
    const name = (data.mode==='video'?data.title:data.name).replace(/[^a-z0-9]/gi,'_');
    res.set({'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':`attachment; filename="${name}_Report.docx"`});
    res.send(buf);
  } catch (err) { res.status(500).json({error:err.message}); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`YT Research running → http://localhost:${PORT}`));
