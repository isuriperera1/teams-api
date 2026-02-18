const express = require('express');
const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const app = express();
const PORT = process.env.PORT || 3002;

// ================== MSAL ==================
const msalClient = new ConfidentialClientApplication({
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET
    }
});

// ================== GET APP TOKEN ==================
async function getAppToken() {
    const result = await msalClient.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"]
    });
    return result.accessToken;
}

// ================== HEALTH ==================
app.get("/", (req, res) => {
    res.json({
        status: "running",
        endpoints: {
            recordings: "/api/recordings",
            download: "/api/download/:itemId"
        }
    });
});

// ================== LIST RECORDINGS ==================
app.get("/api/recordings", async (req, res) => {
    try {
        const token = await getAppToken();
        const userId = process.env.TEAMS_USER_ID;

        const response = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/search(q='.mp4')?$select=id,name,lastModifiedDateTime,webUrl`,
            {
                headers: { Authorization: `Bearer ${token}` }
            }
        );

        const recordings = (response.data.value || []).map(file => ({
            meetingName: file.name,
            meetingDate: file.lastModifiedDateTime,
            itemId: file.id,
            stableDownloadUrl: `${req.protocol}://${req.get('host')}/api/download/${file.id}`,
            webUrl: file.webUrl
        }));

        res.json({
            total: recordings.length,
            recordings
        });

    } catch (err) {
        res.status(500).json({
            error: err.response?.data || err.message
        });
    }
});

// ================== STABLE DOWNLOAD ENDPOINT ==================
app.get("/api/download/:itemId", async (req, res) => {
    try {
        const token = await getAppToken();
        const userId = process.env.TEAMS_USER_ID;
        const itemId = req.params.itemId;

        const fileResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${itemId}/content`,
            {
                headers: { Authorization: `Bearer ${token}` },
                responseType: 'stream'
            }
        );

        res.setHeader('Content-Type', 'video/mp4');
        fileResponse.data.pipe(res);

    } catch (err) {
        res.status(500).json({
            error: err.response?.data || err.message
        });
    }
});

// ================== START SERVER ==================
app.listen(PORT, () => {
    console.log(`
=========================================
Teams Stable Recording Server Running
=========================================
GET /api/recordings
GET /api/download/:itemId
=========================================
`);
});
