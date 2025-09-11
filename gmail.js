import express from "express";
import fs from "fs";
import path from "path";
import multer from "multer";
import nodemailer from "nodemailer";
import cors from "cors";
import dotenv from "dotenv";
import { neon } from "@neondatabase/serverless";
import readXlsxFile from "read-excel-file/node";
import { google } from "googleapis";
import { fileURLToPath } from "url";  

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config();
const app = express();
const client = neon(process.env.db_url);

app.use(cors({ origin: "http://localhost:3000" }));
app.use(express.json());

const upload = multer({ dest: "uploads/" });

async function initDB() {
  await client`
    CREATE TABLE IF NOT EXISTS sent_emails (
      id SERIAL PRIMARY KEY,
      receiver TEXT NOT NULL,
      subject TEXT,
      message TEXT,
      sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
  `;

  await client`
    CREATE TABLE IF NOT EXISTS receive_emails (
      id SERIAL PRIMARY KEY,
      gmail_id TEXT UNIQUE,
      subject TEXT,
      sender TEXT,
      receiver TEXT,
      received_at TIMESTAMP
    )
  `;
  console.log(" Tables ready in Neon");
}
initDB();

const oAuth2Client = new google.auth.OAuth2(
  process.env.client_id,
  process.env.client_secret,
  process.env.redirect_uris,
);

oAuth2Client.setCredentials({
  refresh_token: process.env.refresh_token
});

app.post("/send-email", upload.array("attachments"), async (req, res) => {
  try {
    const { to, subject, message } = req.body;

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: { user: process.env.MAIL, pass: process.env.PASS },
    });

    await transporter.sendMail({
      to,
      subject,
      text: message,
      attachments: req.files.map((file) => ({
        filename: file.originalname,
        path: file.path,
      })),
    });

    await client`
      INSERT INTO sent_emails (receiver, subject, message)
      VALUES (${to}, ${subject}, ${message})
    `;

    res.json({ message: "Email sent successfully!" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to send email" });
  }
});

app.post("/send-excel-emails", upload.single("excel"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const rows = await readXlsxFile(req.file.path);
    if (rows.length < 2) {
      return res.status(400).json({ error: "Excel file has no data" });
    }

    const headers = rows[0];
    const emailIdx = headers.findIndex((h) => h.toLowerCase() === "email");
    const subjectIdx = headers.findIndex((h) => h.toLowerCase() === "subject");
    const messageIdx = headers.findIndex((h) => h.toLowerCase() === "message");

    if (emailIdx === -1) {
      return res.status(400).json({ error: "Excel must contain 'Email' column" });
    }

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: { user: process.env.MAIL, pass: process.env.PASS },
    });

    let sentCount = 0;
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const to = row[emailIdx];
      const subject = subjectIdx !== -1 ? row[subjectIdx] : "No Subject";
      const message = messageIdx !== -1 ? row[messageIdx] : "";
      if (!to) continue;

      await transporter.sendMail({ to, subject, text: message });
      await client`
        INSERT INTO sent_emails (receiver, subject, message)
        VALUES (${to}, ${subject}, ${message})
      `;
      sentCount++;
    }

    res.json({ message: `${sentCount} emails sent successfully!` });
  } catch (err) {
    console.error("Error sending excel emails:", err);
    res.status(500).json({ error: "Failed to process Excel emails" });
  }
});

app.get("/read-mails", async (req, res) => {
  try {
    const gmail = google.gmail({ version: "v1", auth: oAuth2Client });
    const listRes = await gmail.users.messages.list({
      userId: "me",
      q: "is:unread",
      maxResults: 10,
    });

    const messages = [];
    if (listRes.data.messages) {
      for (const msg of listRes.data.messages) {
        const msgRes = await gmail.users.messages.get({
          userId: "me",
          id: msg.id,
        });

        const headers = msgRes.data.payload.headers;
        const subject = headers.find((h) => h.name === "Subject")?.value || "";
        const from = headers.find((h) => h.name === "From")?.value || "";
        const date = headers.find((h) => h.name === "Date")?.value || "";
        const receiver = process.env.MAIL;

        messages.push({ id: msg.id, subject, from, to: receiver, date });

        await client`
          INSERT INTO receive_emails (gmail_id, subject, sender, receiver, received_at)
          VALUES (${msg.id}, ${subject}, ${from}, ${receiver}, ${new Date(date)})
          ON CONFLICT (gmail_id) DO NOTHING
        `;
      }
    }
    messages.sort((a, b) => new Date(b.date) - new Date(a.date));
    res.json({ messages });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error reading mails");
  }
});


app.get("/logout", async (req, res) => {
  try {
    if (fs.existsSync(TOKEN_PATH)) {
      const token = JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8"));
      await oAuth2Client.revokeToken(token.access_token);
      fs.unlinkSync(TOKEN_PATH);
      console.log("Token revoked and deleted!");
    }
    res.send("You have been logged out successfully.");
  } catch (error) {
    console.error(error);
    res.status(500).send("Error logging out.");
  }
});

app.listen(process.env.PORT, () => {
  console.log(` Server running on http://localhost:${process.env.PORT}`);
});