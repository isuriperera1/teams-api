const express = require('express');
const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const app = express();
const PORT = process.env.PORT || 3002;

// ================== MSAL CLIENT ==================
const msalClient = new ConfidentialClientApplication({
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET
    }
});

// ================== GET APPLICATION TOKEN ==================
async function getAppToken() {
    const result = await msalClient.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"]
    });

    if (!result || !result.accessToken) {
        throw new Error("Failed to acquire Graph access token");
    }

    return result.accessToken;
}

// ================== HEALTH CHECK ==================
app.get("/", (req, res) => {
    res.json({
        name: "Teams Recording Fetch API",
        status: "running",
        endpoints: {
            recordings: "/api/recordings?fromDate=YYYY-MM-DD"
        }
    });
});

// ================== FETCH RECORDINGS ==================
app.get("/api/recordings", async (req, res) => {
    try {
        const token = await getAppToken();

        const userId = process.env.TEAMS_USER_ID;
        if (!userId) {
            return res.status(400).json({ error: "TEAMS_USER_ID env variable not set" });
        }

        // Optional filter by date
        const fromDate = req.query.fromDate ? new Date(req.query.fromDate) : null;

        const graphResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/search(q='.mp4')?$select=id,name,size,webUrl,@microsoft.graph.downloadUrl,lastModifiedDateTime&$top=100`,
            {
                headers: {
                    Authorization: `Bearer ${token}`
                }
            }
        );

        const files = graphResponse.data.value || [];

        // Filter only Teams recordings
        let recordings = files.filter(file =>
            file.name.toLowerCase().includes("meeting") ||
            file.name.toLowerCase().includes("recording")
        );

        // Optional date filter
        if (fromDate) {
            recordings = recordings.filter(file =>
                new Date(file.lastModifiedDateTime) >= fromDate
            );
        }

        res.json({
            total: recordings.length,
            recordings: recordings.map(file => ({
                id: file.id,
                name: file.name,
                sizeMB: (file.size / 1024 / 1024).toFixed(2),
                lastModified: file.lastModifiedDateTime,
                webUrl: file.webUrl,
                downloadUrl: file["@microsoft.graph.downloadUrl"]
            }))
        });

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
Teams Recording API Running
=========================================
GET /api/recordings
Optional: ?fromDate=2026-01-01
=========================================
`);
});
