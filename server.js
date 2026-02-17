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

/* ================= AUTH ================= */

app.get('/auth/login', async (req, res) => {
    const authUrl = await msalClient.getAuthCodeUrl({
        scopes: [
            'User.Read',
            'Files.Read.All',
            'Calendars.Read',
            'OnlineMeetingTranscript.Read',
            'OnlineMeetings.Read'
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
                'Files.Read.All',
                'Calendars.Read',
                'OnlineMeetingTranscript.Read',
                'OnlineMeetings.Read'
            ],
            redirectUri: config.redirectUri
        });

        tokenCache = {
            accessToken: tokenResponse.accessToken,
            expiresOn: tokenResponse.expiresOn,
            account: tokenResponse.account
        };

        res.send("<h1>Authentication Successful</h1>");
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

async function callGraph(endpoint) {
    if (!tokenCache.accessToken) throw new Error("Not authenticated");
    const response = await axios.get(`https://graph.microsoft.com/v1.0${endpoint}`, {
        headers: { Authorization: `Bearer ${tokenCache.accessToken}` }
    });
    return response.data;
}

/* ================= RECORDINGS ================= */

app.get('/api/recordings', async (req, res) => {
    try {
        const data = await callGraph(
            "/me/drive/root/search(q='.mp4')?$select=id,name,size,webUrl,lastModifiedDateTime,@microsoft.graph.downloadUrl&$top=50"
        );

        res.json({
            success: true,
            recordings: data.value || []
        });

    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

/* ================= REAL TRANSCRIPTS ================= */

app.get('/api/transcripts', async (req, res) => {
    try {
        // 1️⃣ Get all user meetings
        const meetings = await callGraph("/me/onlineMeetings?$top=50");

        let allTranscripts = [];

        for (const meeting of meetings.value || []) {
            try {
                const transcripts = await callGraph(
                    `/me/onlineMeetings/${meeting.id}/transcripts`
                );

                for (const transcript of transcripts.value || []) {
                    allTranscripts.push({
                        meetingId: meeting.id,
                        meetingSubject: meeting.subject,
                        transcriptId: transcript.id,
                        createdDateTime: transcript.createdDateTime,
                        contentUrl: transcript.contentUrl
                    });
                }

            } catch (innerError) {
                // Skip meetings without transcripts
            }
        }

        res.json({
            success: true,
            count: allTranscripts.length,
            transcripts: allTranscripts
        });

    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

/* ================= DOWNLOAD TRANSCRIPT CONTENT ================= */

app.get('/api/transcripts/:meetingId/:transcriptId', async (req, res) => {
    try {
        const { meetingId, transcriptId } = req.params;

        const response = await axios.get(
            `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`,
            {
                headers: { Authorization: `Bearer ${tokenCache.accessToken}` }
            }
        );

        res.setHeader("Content-Type", "text/vtt");
        res.send(response.data);

    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

/* ================= START SERVER ================= */

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
