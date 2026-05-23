import express from "express";
import path from "path";
import { GoogleGenAI } from "@google/genai";
import { createServer as createViteServer } from "vite";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // API routes
  app.post("/api/analyze", async (req, res) => {
    const { prompt, config } = req.body;
    try {
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY! });
      const response = await ai.models.generateContent({
        model: "gemini-3.1-pro-preview",
        contents: prompt,
        config: config
      });
      res.json({ text: response.text });
    } catch (error) {
      console.error("API error:", error);
      res.status(500).json({ error: "Failed to analyze" });
    }
  });

  app.post("/api/sheets/export", async (req, res) => {
    const { spreadsheetId, accessToken, candidates } = req.body;
    try {
        // Simple append for now
        const values = candidates.map((c: any) => [c.name, c.email, c.jobProfile, c.location, new Date().toISOString()]);
        
        await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/Sheet1!A1:append?valueInputOption=RAW`, {
            method: 'POST',
            headers: { 
                Authorization: `Bearer ${accessToken}`, 
                'Content-Type': 'application/json' 
            },
            body: JSON.stringify({ values })
        });
        
        res.json({ message: "Exported successfully" });
    } catch (error) {
        console.error("Sheets Export error:", error);
        res.status(500).json({ error: "Failed to export" });
    }
  });

  app.get("/api/sheets/test", async (req, res) => {
      // Placeholder for Sheets API interaction
      res.json({ message: "Sheets API endpoint ready" });
  });

  app.post("/api/login", async (req, res) => {
    const { username, password } = req.body;
    if ((username === process.env.DEV_RECRUITER_USERNAME && password === process.env.DEV_RECRUITER_PASSWORD) || 
        (username === "Mariya" && password === "9442")) {
        res.json({ 
            user: { 
                name: username, 
                email: `${username}@recruiter.com`, 
                isAdmin: true 
            } 
        });
    } else {
        res.status(401).json({ error: "Invalid credentials" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
