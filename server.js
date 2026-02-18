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
        server: "Teams Stable Recording Server",
        endpoints: {
            listRecordings: "/api/recordings",
            downloadRecording: "/api/download/:itemId"
        }
    });
});

// ================== LIST ALL MP4 RECORDINGS ==================
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
            graphWebUrl: file.webUrl,

            // 🔥 THIS is your stable hardcoded backend download URL
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

// ================== STREAM RECORDING (STABLE DOWNLOAD LINK) ==================
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
        res.setHeader("Content-Disposition", "inline");

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
Teams Stable Recording Server Running
==================================================
GET  /api/recordings
GET  /api/download/:itemId
==================================================
`);
});
