import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { OAuth2Client } from "google-auth-library";
import { google } from "googleapis";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

// Google OAuth Configuration
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = `${process.env.APP_URL || 'http://localhost:3000'}/auth/google/callback`;

const oauth2Client = new OAuth2Client(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

app.use(express.json());

// API Routes
app.get("/api/auth/google/url", (req, res) => {
  if (!CLIENT_ID || !CLIENT_SECRET) {
    return res.status(500).json({ error: "Google OAuth credentials not configured." });
  }

  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/documents.readonly",
    ],
    prompt: "consent",
  });

  res.json({ url });
});

app.get("/auth/google/callback", async (req, res) => {
  const { code } = req.query;
  if (!code) {
    return res.status(400).send("No code provided.");
  }

  try {
    const { tokens } = await oauth2Client.getToken(code as string);
    
    // Send success message to parent window and close popup
    // We pass the tokens back to the client for this demo
    res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ 
                type: 'OAUTH_AUTH_SUCCESS', 
                tokens: ${JSON.stringify(tokens)} 
              }, '*');
              window.close();
            } else {
              window.location.href = '/';
            }
          </script>
          <p>Authentication successful. This window should close automatically.</p>
        </body>
      </html>
    `);
  } catch (error) {
    console.error("OAuth callback error:", error);
    res.status(500).send("Authentication failed.");
  }
});

import mammoth from "mammoth";

// ... (existing code)

// Proxy to fetch Google Doc content
app.post("/api/gdoc/fetch", async (req, res) => {
  const { docId, accessToken } = req.body;
  if (!docId || !accessToken) {
    return res.status(400).json({ error: "Missing docId or accessToken." });
  }

  try {
    const auth = new OAuth2Client();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: "v3", auth });
    
    // Get file metadata to check mimeType
    const fileMetadata = await drive.files.get({
      fileId: docId,
      fields: "id, name, mimeType",
    });

    const mimeType = fileMetadata.data.mimeType;
    let content = "";

    if (mimeType === "application/vnd.google-apps.document") {
      // It's a native Google Doc, use Docs API or Drive Export
      const docs = google.docs({ version: "v1", auth });
      const doc = await docs.documents.get({ documentId: docId });
      
      if (doc.data.body && doc.data.body.content) {
        doc.data.body.content.forEach(element => {
          if (element.paragraph) {
            element.paragraph.elements?.forEach(el => {
              if (el.textRun) {
                content += el.textRun.content;
              }
            });
          }
        });
      }
    } else if (mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
      // It's a Word file on Drive, download and parse with mammoth
      const response = await drive.files.get(
        { fileId: docId, alt: "media" },
        { responseType: "arraybuffer" }
      );
      
      const buffer = Buffer.from(response.data as ArrayBuffer);
      const result = await mammoth.extractRawText({ buffer });
      content = result.value;
    } else {
      throw new Error(`このファイル形式（${mimeType}）は、Googleドキュメントタブではサポートされていません。GoogleドキュメントまたはWordファイルを選択してください。`);
    }

    res.json({ content, title: fileMetadata.data.name });
  } catch (error: any) {
    console.error("Google Drive/Docs API error details:", error.response?.data || error.message);
    const message = error.response?.data?.error?.message || error.message || "Failed to fetch document content.";
    res.status(500).json({ error: message });
  }
});

// Vite middleware setup
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
