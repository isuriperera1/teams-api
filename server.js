const express = require('express');
const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const app = express();
const PORT = process.env.PORT || 3002;

const msalClient = new ConfidentialClientApplication({
    auth: {
        clientId: process.env.AZURE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
        clientSecret: process.env.AZURE_CLIENT_SECRET
    }
});

// 🔐 Get application token
async function getAppToken() {
    const result = await msalClient.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"]
    });
    return result.accessToken;
}

// 🔥 Get ALL recap transcripts across tenant
app.get('/api/transcripts', async (req, res) => {
    try {
        const token = await getAppToken();

        // 1️⃣ Get call records
        const callRecordsResponse = await axios.get(
            "https://graph.microsoft.com/beta/communications/callRecords",
            {
                headers: { Authorization: `Bearer ${token}` }
            }
        );

        const callRecords = callRecordsResponse.data.value || [];
        let allTranscripts = [];

        // 2️⃣ For each call record fetch transcripts
        for (const record of callRecords) {
            try {
                const transcriptResponse = await axios.get(
                    `https://graph.microsoft.com/beta/communications/callRecords/${record.id}/transcripts`,
                    {
                        headers: { Authorization: `Bearer ${token}` }
                    }
                );

                for (const transcript of transcriptResponse.data.value || []) {

                    // 3️⃣ Fetch transcript content
                    let content = null;
                    try {
                        const contentResponse = await axios.get(
                            `https://graph.microsoft.com/beta/communications/callRecords/${record.id}/transcripts/${transcript.id}/content`,
                            {
                                headers: { Authorization: `Bearer ${token}` }
                            }
                        );
                        content = contentResponse.data;
                    } catch {}

                    allTranscripts.push({
                        callRecordId: record.id,
                        transcriptId: transcript.id,
                        createdDateTime: transcript.createdDateTime,
                        content: content
                    });
                }
            } catch {}
        }

        res.json({
            total: allTranscripts.length,
            transcripts: allTranscripts
        });

    } catch (err) {
        res.status(500).json({
            error: err.response?.data || err.message
        });
    }
});

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
