const express = require('express');
const cors = require('cors');
const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const app = express();
const PORT = process.env.PORT || 3002;

const BASE_URL = process.env.RAILWAY_PUBLIC_DOMAIN 
    ? `https://${process.env.RAILWAY_PUBLIC_DOMAIN}`
    : process.env.RENDER_EXTERNAL_URL 
    ? process.env.RENDER_EXTERNAL_URL
    : `http://localhost:${PORT}`;

const config = {
    clientId: process.env.AZURE_CLIENT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    tenantId: process.env.AZURE_TENANT_ID,
    redirectUri: `${BASE_URL}/auth/callback`
};

const msalClient = new ConfidentialClientApplication({
    auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientSecret: config.clientSecret
    }
});

let tokenCache = {};

app.use(cors());
app.use(express.json());

// AUTH ROUTES
app.get('/auth/login', async (req, res) => {
    const authUrl = await msalClient.getAuthCodeUrl({
        scopes: [
            'User.Read',
            'OnlineMeetings.Read',
            'OnlineMeetingTranscript.Read.All',
            'Calendars.Read',
            'Files.Read.All'
        ],
        redirectUri: config.redirectUri
    });
    res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
    try {
        const tokenResponse = await msalClient.acquireTokenByCode({
            code: req.query.code,
            scopes: [
                'User.Read',
                'OnlineMeetings.Read',
                'OnlineMeetingTranscript.Read.All',
                'Calendars.Read',
                'Files.Read.All'
            ],
            redirectUri: config.redirectUri
        });
        
        tokenCache = {
            accessToken: tokenResponse.accessToken,
            expiresOn: tokenResponse.expiresOn,
            account: tokenResponse.account
        };

        res.send("<h1>Authentication successful 🎉</h1>");
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

async function graphGetBeta(path) {
    if (!tokenCache.accessToken) throw new Error("Not authenticated");
    const url = `https://graph.microsoft.com/beta${path}`;
    const r = await axios.get(url, {
        headers: {
            Authorization: `Bearer ${tokenCache.accessToken}`
        }
    });
    return r.data;
}

// REAL TRANSCRIPTS API
app.get('/api/transcripts', async (req, res) => {
    try {
        // 1️⃣ Get the current user’s online meetings
        const meetings = await graphGetBeta('/me/onlineMeetings');

        let results = [];

        // 2️⃣ Loop each meeting and fetch transcripts
        for (const meeting of meetings.value || []) {
            try {
                const transcripts = await graphGetBeta(
                    `/me/onlineMeetings/${meeting.id}/transcripts`
                );

                for (const t of transcripts.value || []) {
                    // 3️⃣ Fetch actual transcript text
                    let textResponse = null;
                    try {
                        const textRes = await axios.get(t.contentUrl);
                        textResponse = textRes.data;
                    } catch (fetchErr) {
                        textResponse = null;
                    }

                    results.push({
                        meetingId: meeting.id,
                        meetingSubject: meeting.subject,
                        transcriptId: t.id,
                        created: t.createdDateTime,
                        contentUrl: t.contentUrl,
                        text: textResponse
                    });
                }
            } catch {}
        }

        res.json({
            success: true,
            total: results.length,
            transcripts: results
        });
    } catch (err) {
        res.status(500).json({ success:false, error: err.message });
    }
});

app.listen(PORT, () => {
    console.log(`Server listening on port ${PORT}`);
});
