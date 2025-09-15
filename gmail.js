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

  await client`
    CREATE TABLE IF NOT EXISTS unread_emails (
      id TEXT PRIMARY KEY,
      subject TEXT,
      sender TEXT,
      receiver TEXT,
      received_at TIMESTAMP
    )
  `;

  await client`
    CREATE TABLE IF NOT EXISTS last_seen_count (
      id SERIAL PRIMARY KEY,
      last_count INT DEFAULT 0,
      updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
  `;

  console.log("Tables ready in Neon");
}
initDB();

const oAuth2Client = new google.auth.OAuth2(
  process.env.client_id,
  process.env.client_secret,
  process.env.redirect_uris
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
      attachments: req.files?.map((file) => ({
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
    if (rows.length < 2) return res.status(400).json({ error: "Excel file has no data" });

    const headers = rows[0];
    const emailIdx = headers.findIndex((h) => h.toLowerCase() === "email");
    const subjectIdx = headers.findIndex((h) => h.toLowerCase() === "subject");
    const messageIdx = headers.findIndex((h) => h.toLowerCase() === "message");

    if (emailIdx === -1) return res.status(400).json({ error: "Excel must contain 'Email' column" });

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

app.get("/sent-emails", async (req, res) => {
  try {
    const result = await client`
      SELECT id, receiver, subject, message
      FROM sent_emails
      ORDER BY id DESC
      LIMIT 21
    `;
    res.json({ sentEmails: result });
  } catch (err) {
    console.error("Error fetching sent mails:", err);
    res.status(500).json({ error: "Failed to fetch sent emails" });
  }
});

app.get("/read-mails", async (req, res) => {
  try {
    const countResult = await client`
      SELECT COUNT(*) AS count FROM unread_emails
    `;
    const totalCount = parseInt(countResult[0].count, 10);

    const lastSeenRes = await client`
      SELECT id, last_count FROM last_seen_count ORDER BY id DESC LIMIT 1
    `;

    if (lastSeenRes.length) {
      await client`
        UPDATE last_seen_count
        SET last_count = ${totalCount}, updated_at = NOW()
        WHERE id = ${lastSeenRes[0].id}
      `;
    } else {
      await client`
        INSERT INTO last_seen_count (last_count) VALUES (${totalCount})
      `;
    }

    const gmail = google.gmail({ version: "v1", auth: oAuth2Client });
    const listRes = await gmail.users.messages.list({
      userId: "me",
      q: "is:unread",
      maxResults: 10,
    });

    if (!listRes.data.messages || !listRes.data.messages.length) {
      return res.json({ messages: [] });
    }

    const messages = [];

    for (const msg of listRes.data.messages) {
      const msgRes = await gmail.users.messages.get({
        userId: "me",
        id: msg.id,
      });

      const headers = msgRes.data.payload.headers;
      const subject = headers.find(h => h.name === "Subject")?.value || "(No Subject)";
      const from = headers.find(h => h.name === "From")?.value || "(Unknown)";
      const to = headers.find(h => h.name === "To")?.value || process.env.MAIL;
      const date = headers.find(h => h.name === "Date")?.value;

      const row = {
        id: msg.id,
        subject,
        sender: from,
        receiver: to,
        received_at: new Date(date),
      };

      messages.push(row);

      await client`
        INSERT INTO receive_emails (gmail_id, subject, sender, receiver, received_at)
        VALUES (${row.id}, ${row.subject}, ${row.sender}, ${row.receiver}, ${row.received_at})
        ON CONFLICT (gmail_id) DO NOTHING
      `;

      await client`
        INSERT INTO unread_emails (id, subject, sender, receiver, received_at)
        VALUES (${row.id}, ${row.subject}, ${row.sender}, ${row.receiver}, ${row.received_at})
        ON CONFLICT (id) DO NOTHING
      `;
    }

    messages.sort((a, b) => new Date(b.received_at) - new Date(a.received_at));
    res.redirect("/getmails");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error reading mails");
  }
});

app.get("/getmails", async (req, res) => {
  try {
    const lastSeenRes = await client`
      SELECT last_count FROM last_seen_count ORDER BY id DESC LIMIT 1
    `;
    const lastCount = lastSeenRes.length ? lastSeenRes[0].last_count : 0;

    const result = await client`
      SELECT * FROM unread_emails ORDER BY received_at ASC OFFSET ${lastCount}
    `;

    res.json({
      message: "success",
      newCount: result.length,
      data: result,
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Error getting mails");
  }
});

app.get("/logout", async (req, res) => {
  try {
    if (fs.existsSync("token.json")) {
      const token = JSON.parse(fs.readFileSync("token.json", "utf8"));
      await oAuth2Client.revokeToken(token.access_token);
      fs.unlinkSync("token.json");
      console.log("Token revoked and deleted!");
    }
    res.send("You have been logged out successfully.");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error logging out.");
  }
});

// ... keep your other imports and routes above

// ---- Delta mails route ----
app.get("/delta-mails", async (req, res) => {
  try {
    const lastId = req.query.lastId || null;

    // Total unread mails count
    const [{ total }] =
      await client`SELECT COUNT(*)::int AS total FROM unread_emails`;

    let rows;
    if (lastId) {
      // next 10 newer than lastId
      rows = await client`
        SELECT * FROM unread_emails
        WHERE id > ${lastId}
        ORDER BY received_at ASC
        LIMIT 10
      `;
    } else {
      // first call: oldest 10 unread
      rows = await client`
        SELECT * FROM unread_emails
        ORDER BY received_at ASC
        LIMIT 10
      `;
    }

    res.json({
      message: "success",
      totalCount: total, // 👈 total unread count
      data: rows,
      lastDeltaId: rows.length
        ? rows[rows.length - 1].id
        : lastId || null,
    });
  } catch (err) {
    console.error("delta fetch error:", err);
    res.status(500).json({ error: "Failed to fetch delta mails" });
  }
});

// ---- start server ----
app.listen(process.env.PORT, () => {
  console.log(`Server running on http://localhost:${process.env.PORT}`);
});
