const express = require('express');
const cors = require('cors');
const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const app = express();
const PORT = process.env.PORT || 3002;

// Detect if running on cloud or localhost
const BASE_URL = process.env.RAILWAY_PUBLIC_DOMAIN 
    ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}`
    : process.env.RENDER_EXTERNAL_URL 
    ? process.env.RENDER_EXTERNAL_URL
    : `http://localhost:${PORT}`;

// Azure AD Configuration
const config = {
    clientId: process.env.AZURE_CLIENT_ID || '4dc5efa6-49cb-4991-9cbf-1867596a2110',
    clientSecret: process.env.AZURE_CLIENT_SECRET || 'vqs8Q~s1bjrgYq1T8XV1lcySV_gwf9~DhXJEQcV~',
    tenantId: process.env.AZURE_TENANT_ID || 'abd34cc0-2911-44eb-abea-713c02c9b3e8',
    redirectUri: `${BASE_URL}/auth/callback`
};

// MSAL Configuration for server-side auth
const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientSecret: config.clientSecret
    }
};

const msalClient = new ConfidentialClientApplication(msalConfig);

// Store tokens (in production, use a database)
let tokenCache = {};

app.use(cors());
app.use(express.json());

// Health check
app.get('/', (req, res) => {
    res.json({
        name: 'Teams Recordings API',
        version: '1.0.0',
        status: 'running',
        endpoints: {
            auth: '/auth/login',
            recordings: '/api/recordings',
            transcripts: '/api/transcripts',
            calendar: '/api/calendar',
            download: '/api/download/:fileId'
        }
    });
});

// ============ AUTH ENDPOINTS ============

// Start OAuth flow
app.get('/auth/login', async (req, res) => {
    const authUrl = await msalClient.getAuthCodeUrl({
        scopes: ['User.Read', 'Files.Read.All', 'Calendars.Read'],
        redirectUri: config.redirectUri
    });
    res.redirect(authUrl);
});

// OAuth callback
app.get('/auth/callback', async (req, res) => {
    try {
        const tokenResponse = await msalClient.acquireTokenByCode({
            code: req.query.code,
            scopes: ['User.Read', 'Files.Read.All', 'Calendars.Read'],
            redirectUri: config.redirectUri
        });
        
        tokenCache = {
            accessToken: tokenResponse.accessToken,
            expiresOn: tokenResponse.expiresOn,
            account: tokenResponse.account
        };
        
        res.send(`
            <html>
            <body style="font-family: Segoe UI; text-align: center; padding: 50px;">
                <h1>✅ Authentication Successful!</h1>
                <p>You can now use the API endpoints.</p>
                <p>Access Token stored. You can close this window.</p>
                <script>setTimeout(() => window.close(), 3000);</script>
            </body>
            </html>
        `);
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Check auth status
app.get('/auth/status', (req, res) => {
    const isAuthenticated = tokenCache.accessToken && new Date() < new Date(tokenCache.expiresOn);
    res.json({
        authenticated: isAuthenticated,
        user: tokenCache.account?.username || null,
        expiresOn: tokenCache.expiresOn || null
    });
});

// ============ HELPER FUNCTION ============

async function callGraphAPI(endpoint) {
    if (!tokenCache.accessToken) {
        throw new Error('Not authenticated. Visit /auth/login first.');
    }
    
    // Check if token expired
    if (new Date() >= new Date(tokenCache.expiresOn)) {
        throw new Error('Token expired. Please re-authenticate at /auth/login');
    }
    
    const response = await axios.get(`https://graph.microsoft.com/v1.0${endpoint}`, {
        headers: { 'Authorization': `Bearer ${tokenCache.accessToken}` }
    });
    
    return response.data;
}

// ============ API ENDPOINTS FOR N8N ============

// GET /api/recordings - List all recordings
app.get('/api/recordings', async (req, res) => {
    try {
        const data = await callGraphAPI(
            "/me/drive/root/search(q='.mp4')?$select=id,name,size,webUrl,lastModifiedDateTime,@microsoft.graph.downloadUrl&$top=50"
        );
        
        const recordings = (data.value || []).filter(f => 
            f.name.toLowerCase().includes('recording') || f.name.endsWith('.mp4')
        );
        
        res.json({
            success: true,
            count: recordings.length,
            recordings: recordings.map(r => ({
                id: r.id,
                name: r.name,
                size: r.size,
                sizeFormatted: (r.size / 1024 / 1024).toFixed(2) + ' MB',
                date: r.lastModifiedDateTime,
                webUrl: r.webUrl,
                downloadUrl: r['@microsoft.graph.downloadUrl']
            }))
        });
    } catch (error) {
        res.status(error.message.includes('authenticate') ? 401 : 500).json({
            success: false,
            error: error.message
        });
    }
});

// GET /api/transcripts - List all transcripts
app.get('/api/transcripts', async (req, res) => {
    try {
        const data = await callGraphAPI(
            "/me/drive/root/search(q='transcript')?$select=id,name,size,webUrl,lastModifiedDateTime,@microsoft.graph.downloadUrl&$top=50"
        );
        
        const transcripts = (data.value || []).filter(f => 
            f.name.toLowerCase().includes('transcript') || 
            f.name.endsWith('.vtt') || 
            f.name.endsWith('.docx')
        );
        
        res.json({
            success: true,
            count: transcripts.length,
            transcripts: transcripts.map(t => ({
                id: t.id,
                name: t.name,
                size: t.size,
                date: t.lastModifiedDateTime,
                webUrl: t.webUrl,
                downloadUrl: t['@microsoft.graph.downloadUrl']
            }))
        });
    } catch (error) {
        res.status(error.message.includes('authenticate') ? 401 : 500).json({
            success: false,
            error: error.message
        });
    }
});

// GET /api/calendar - Get upcoming Teams meetings
app.get('/api/calendar', async (req, res) => {
    try {
        const now = new Date().toISOString();
        const later = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
        
        const data = await callGraphAPI(
            `/me/calendarView?startDateTime=${now}&endDateTime=${later}&$select=subject,start,end,isOnlineMeeting,onlineMeetingUrl&$top=20&$orderby=start/dateTime`
        );
        
        res.json({
            success: true,
            count: data.value?.length || 0,
            meetings: (data.value || []).map(e => ({
                subject: e.subject,
                start: e.start,
                end: e.end,
                isTeamsMeeting: e.isOnlineMeeting,
                joinUrl: e.onlineMeetingUrl
            }))
        });
    } catch (error) {
        res.status(error.message.includes('authenticate') ? 401 : 500).json({
            success: false,
            error: error.message
        });
    }
});

// GET /api/download/:fileId - Get download URL for a specific file
app.get('/api/download/:fileId', async (req, res) => {
    try {
        const data = await callGraphAPI(
            `/me/drive/items/${req.params.fileId}?$select=id,name,@microsoft.graph.downloadUrl`
        );
        
        res.json({
            success: true,
            id: data.id,
            name: data.name,
            downloadUrl: data['@microsoft.graph.downloadUrl']
        });
    } catch (error) {
        res.status(error.message.includes('authenticate') ? 401 : 500).json({
            success: false,
            error: error.message
        });
    }
});

// GET /api/file/:fileId/content - Download file content directly
app.get('/api/file/:fileId/content', async (req, res) => {
    try {
        const data = await callGraphAPI(
            `/me/drive/items/${req.params.fileId}?$select=name,@microsoft.graph.downloadUrl`
        );
        
        // Redirect to the download URL
        res.redirect(data['@microsoft.graph.downloadUrl']);
    } catch (error) {
        res.status(error.message.includes('authenticate') ? 401 : 500).json({
            success: false,
            error: error.message
        });
    }
});

// Start server
app.listen(PORT, () => {
    console.log(`
╔═══════════════════════════════════════════════════════════╗
║         Teams Recordings API Server Started!              ║
╠═══════════════════════════════════════════════════════════╣
║  Server running at: http://localhost:${PORT}                 ║
║                                                           ║
║  STEP 1: Authenticate first:                              ║
║  👉 http://localhost:${PORT}/auth/login                      ║
║                                                           ║
║  API Endpoints for n8n:                                   ║
║  • GET /api/recordings    - List all recordings           ║
║  • GET /api/transcripts   - List all transcripts          ║
║  • GET /api/calendar      - Upcoming Teams meetings       ║
║  • GET /api/download/:id  - Get file download URL         ║
╚═══════════════════════════════════════════════════════════╝
    `);
});
