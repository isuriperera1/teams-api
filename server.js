const express = require('express');
const axios = require('axios');
const cors = require('cors');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const app = express();
const PORT = process.env.PORT || 3002;

app.use(cors());
app.use(express.json());

// ================== MSAL CONFIG ==================
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
        throw new Error("Failed to acquire Graph token");
    }
    return result.accessToken;
}

// ================== HEALTH CHECK ==================
app.get("/", (req, res) => {
    res.json({
        status: "running",
        server: "Teams Recording Server",
        endpoints: {
            recordings: "/api/recordings",
            download: "/api/download/:itemId",
            n8n: "/api/n8n/recordings"
        }
    });
});

// ================== LIST ALL MP4 RECORDINGS (FOR FRONTEND) ==================
app.get("/api/recordings", async (req, res) => {
    try {
        const token = await getAppToken();
        const userId = process.env.TEAMS_USER_ID;

        if (!userId) {
            return res.status(500).json({
                error: "TEAMS_USER_ID environment variable missing"
            });
        }

        const response = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/search(q='.mp4')?$select=id,name,lastModifiedDateTime,webUrl,size`,
            {
                headers: { Authorization: `Bearer ${token}` }
            }
        );

        const recordings = (response.data.value || []).map(file => ({
            meetingName: file.name,
            meetingDate: file.lastModifiedDateTime,
            sizeBytes: file.size,
            sizeMB: file.size ? (file.size / 1024 / 1024).toFixed(2) + " MB" : "Unknown",
            itemId: file.id,
            webUrl: file.webUrl,
            stableDownloadUrl: `${req.protocol}://${req.get('host')}/api/download/${file.id}`
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

// ================== 🔥 N8N ENDPOINT - GET RECORDING URLS ==================
app.get("/api/n8n/recordings", async (req, res) => {
    try {
        const token = await getAppToken();
        const userId = process.env.TEAMS_USER_ID;

        if (!userId) {
            return res.status(500).json({
                error: "TEAMS_USER_ID environment variable missing"
            });
        }

        const response = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/search(q='.mp4')?$select=id,name,lastModifiedDateTime,size`,
            {
                headers: { Authorization: `Bearer ${token}` }
            }
        );

        // Clean format for n8n
        const recordings = (response.data.value || []).map(file => ({
            fileName: file.name,
            fileSize: file.size ? (file.size / 1024 / 1024).toFixed(2) + " MB" : "Unknown",
            recordingDate: file.lastModifiedDateTime,
            itemId: file.id,
            downloadUrl: `${req.protocol}://${req.get('host')}/api/download/${file.id}`
        }));

        res.json({
            success: true,
            total: recordings.length,
            recordings: recordings
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.response?.data || err.message
        });
    }
});

// ================== STREAM/DOWNLOAD RECORDING ==================
app.get("/api/download/:itemId", async (req, res) => {
    try {
        const token = await getAppToken();
        const userId = process.env.TEAMS_USER_ID;
        const itemId = req.params.itemId;

        const fileResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${itemId}/content`,
            {
                headers: { Authorization: `Bearer ${token}` },
                responseType: "stream"
            }
        );

        res.setHeader("Content-Type", "video/mp4");
        res.setHeader("Content-Disposition", "attachment"); // Forces download
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
==================================================
  Teams Recording Server
  Port: ${PORT}
==================================================
  Frontend:   /
  Recordings: /api/recordings
  Download:   /api/download/:itemId
  🔥 N8N:     /api/n8n/recordings
==================================================
`);
});
