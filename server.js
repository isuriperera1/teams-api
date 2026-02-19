require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

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

// ================== GOOGLE DRIVE SETUP ==================
const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, 'credentials.json'),  // ← FIXED: Absolute path
    scopes: ['https://www.googleapis.com/auth/drive.file']
});

const drive = google.drive({ version: 'v3', auth });

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

// ================== UPLOAD TO GOOGLE DRIVE ==================
async function uploadToDrive(fileStream, fileName, folderId) {
    const fileMetadata = {
        name: fileName,
        parents: [folderId]
    };

    const media = {
        mimeType: 'video/mp4',
        body: fileStream
    };

    const response = await drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id, name, webViewLink'
    });

    return response.data;
}

// ================== HEALTH CHECK ==================
app.get("/", (req, res) => {
    res.json({
        status: "running",
        server: "Teams Recording Server with Google Drive",
        endpoints: {
            uploadAll: "/api/upload-all-to-drive"
        }
    });
});

// ================== 🔥 UPLOAD ALL RECORDINGS TO GOOGLE DRIVE ==================
app.get("/api/upload-all-to-drive", async (req, res) => {
    try {
        const token = await getAppToken();
        const userId = process.env.TEAMS_USER_ID;
        const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

        if (!userId || !folderId) {
            return res.status(500).json({
                error: "Missing TEAMS_USER_ID or GOOGLE_DRIVE_FOLDER_ID"
            });
        }

        // Get all MP4 recordings
        const response = await axios.get(
            `https://graph.microsoft.com/v1.0/users/${userId}/drive/root/search(q='.mp4')?$select=id,name,lastModifiedDateTime,size`,
            { headers: { Authorization: `Bearer ${token}` } }
        );

        const recordings = response.data.value || [];
        const results = [];

        console.log(`Found ${recordings.length} recordings. Starting upload...`);

        // Upload each recording
        for (const file of recordings) {
            try {
                console.log(`Uploading: ${file.name}...`);

                // Download from OneDrive as stream
                const fileStream = await axios.get(
                    `https://graph.microsoft.com/v1.0/users/${userId}/drive/items/${file.id}/content`,
                    {
                        headers: { Authorization: `Bearer ${token}` },
                        responseType: 'stream'
                    }
                );

                // Upload to Google Drive
                const driveFile = await uploadToDrive(
                    fileStream.data,
                    file.name,
                    folderId
                );

                results.push({
                    fileName: file.name,
                    status: 'uploaded',
                    googleDriveLink: driveFile.webViewLink,
                    fileId: driveFile.id
                });

                console.log(`✅ Uploaded: ${file.name}`);
            } catch (err) {
                results.push({
                    fileName: file.name,
                    status: 'failed',
                    error: err.message
                });
                console.error(`❌ Failed: ${file.name} - ${err.message}`);
            }
        }

        res.json({
            success: true,
            total: recordings.length,
            uploaded: results.filter(r => r.status === 'uploaded').length,
            failed: results.filter(r => r.status === 'failed').length,
            results: results
        });
    } catch (err) {
        res.status(500).json({
            success: false,
            error: err.response?.data || err.message
        });
    }
});

// ================== START SERVER ==================
app.listen(PORT, () => {
    console.log(`
==================================================
  Teams Recording → Google Drive Server
  Port: ${PORT}
==================================================
  Test: http://localhost:${PORT}
  Upload All: http://localhost:${PORT}/api/upload-all-to-drive
==================================================
`);
});
