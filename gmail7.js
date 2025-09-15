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

app.use(cors({ origin: "https://mail-sender-alpha.vercel.app" }));
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
  await client`
    CREATE TABLE IF NOT EXISTS inbox_stats (
      id SERIAL PRIMARY KEY,
      total_inbox INT DEFAULT 0,
      updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
  `;
  await client`
    CREATE TABLE IF NOT EXISTS all_mails (
      id TEXT PRIMARY KEY,
      subject TEXT,
      sender TEXT,
      receiver TEXT,
      received_at TIMESTAMP,
      seen BOOLEAN DEFAULT FALSE
    )
  `;
  await client`
    CREATE TABLE IF NOT EXISTS total_mail (
      id SERIAL PRIMARY KEY,
      total_count INT DEFAULT 0,
      updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
  `;
  await client`
    CREATE TABLE IF NOT EXISTS gmail_history (
      id SERIAL PRIMARY KEY,
      history_id BIGINT,
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
oAuth2Client.setCredentials({ refresh_token: process.env.refresh_token });

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
      attachments: req.files?.map(f => ({ filename: f.originalname, path: f.path })),
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
    const emailIdx = headers.findIndex(h => h.toLowerCase() === "email");
    const subjectIdx = headers.findIndex(h => h.toLowerCase() === "subject");
    const messageIdx = headers.findIndex(h => h.toLowerCase() === "message");
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

app.get("/read-mails", async (req, res) => {
  try {
    const gmail = google.gmail({ version: "v1", auth: oAuth2Client });

    const countResult = await client`SELECT COUNT(*) AS count FROM unread_emails`;
    const totalCount = parseInt(countResult[0].count, 10);

    const lastSeenRes = await client`SELECT id, last_count FROM last_seen_count ORDER BY id DESC LIMIT 1`;
    if (lastSeenRes.length) {
      await client`
        UPDATE last_seen_count
        SET last_count = ${totalCount}, updated_at = NOW()
        WHERE id = ${lastSeenRes[0].id}
      `;
    } else {
      await client`INSERT INTO last_seen_count (last_count) VALUES (${totalCount})`;
    }

    const listRes = await gmail.users.messages.list({ userId: "me", q: "is:unread", maxResults: 50 });
    if (!listRes.data.messages || !listRes.data.messages.length) return res.redirect("/getmails");

    for (const msg of listRes.data.messages) {
      const msgRes = await gmail.users.messages.get({ userId: "me", id: msg.id });
      const headers = msgRes.data.payload.headers || [];
      const subject = headers.find(h => h.name === "Subject")?.value || "(No Subject)";
      const from = headers.find(h => h.name === "From")?.value || "(Unknown)";
      const to = headers.find(h => h.name === "To")?.value || process.env.MAIL;
      const dateStr = headers.find(h => h.name === "Date")?.value;
      const received_at = dateStr ? new Date(dateStr) : new Date();

      await client`
        INSERT INTO receive_emails (gmail_id, subject, sender, receiver, received_at)
        VALUES (${msg.id}, ${subject}, ${from}, ${to}, ${received_at})
        ON CONFLICT (gmail_id) DO NOTHING
      `;
      await client`
        INSERT INTO unread_emails (id, subject, sender, receiver, received_at)
        VALUES (${msg.id}, ${subject}, ${from}, ${to}, ${received_at})
        ON CONFLICT (id) DO NOTHING
      `;
      await client`
        INSERT INTO all_mails (id, subject, sender, receiver, received_at, seen)
        VALUES (${msg.id}, ${subject}, ${from}, ${to}, ${received_at}, ${false})
        ON CONFLICT (id) DO NOTHING
      `;
    }

    const totalRes = await client`SELECT COUNT(*) AS count FROM all_mails`;
    const totalCountAll = parseInt(totalRes[0].count, 10);

    const lastTotalRes = await client`SELECT id FROM total_mail ORDER BY id DESC LIMIT 1`;
    if (lastTotalRes.length) {
      await client`
        UPDATE total_mail
        SET total_count = ${totalCountAll}, updated_at = NOW()
        WHERE id = ${lastTotalRes[0].id}
      `;
    } else {
      await client`
        INSERT INTO total_mail (total_count)
        VALUES (${totalCountAll})
      `;
    }

    res.redirect("/getmails");
  } catch (err) {
    console.error("Error reading mails:", err);
    res.status(500).send("Error reading mails");
  }
});

app.get("/getmails", async (req, res) => {
  try {
    const lastSeenRes = await client`SELECT last_count FROM last_seen_count ORDER BY id DESC LIMIT 1`;
    const lastCount = lastSeenRes.length ? lastSeenRes[0].last_count : 0;

    const result = await client`SELECT * FROM unread_emails ORDER BY received_at ASC OFFSET ${lastCount}`;

    res.json({
      message: "success",
      newCount: result.length,
      data: result,
    });
  } catch (err) {
    console.error("Error getting mails:", err);
    res.status(500).send("Error getting mails");
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

app.get("/delta-mails", async (req, res) => {
  try {
    const gmail = google.gmail({ version: "v1", auth: oAuth2Client });

    const last = await client`SELECT history_id FROM gmail_history ORDER BY id DESC LIMIT 1`;
    const startHistoryId = last.length ? last[0].history_id : null;

    if (!startHistoryId) {
      const profile = await gmail.users.getProfile({ userId: "me" });
      await client`INSERT INTO gmail_history (history_id) VALUES (${profile.data.historyId})`;
      return res.json({
        message: "Initial history saved. Click again to fetch new mails.",
        totalUnread: 0,
        data: [],
      });
    }

    const addedIds = [];
    let nextPageToken;
    let newestHistoryId = startHistoryId;

    do {
      const histRes = await gmail.users.history.list({
        userId: "me",
        startHistoryId,
        historyTypes: ["messageAdded"],
        pageToken: nextPageToken,
      });

      if (histRes.data.historyId) newestHistoryId = histRes.data.historyId;

      (histRes.data.history || []).forEach(h =>
        (h.messagesAdded || []).forEach(m => {
          if (m.message && m.message.id) addedIds.push(m.message.id);
        })
      );

      nextPageToken = histRes.data.nextPageToken;
    } while (nextPageToken);

    for (const id of addedIds) {
      const msgRes = await gmail.users.messages.get({ userId: "me",id });
      const headers = msgRes.data.payload?.headers || [];
      const subject = headers.find(h => h.name === "Subject")?.value || "(No Subject)";
      const sender  = headers.find(h => h.name === "From")?.value || "(Unknown)";
      const receiver= headers.find(h => h.name === "To")?.value || process.env.MAIL;
      const dateStr = headers.find(h => h.name === "Date")?.value;
      const received_at = dateStr ? new Date(dateStr) : new Date();

      await client`
        INSERT INTO all_mails (id, subject, sender, receiver, received_at, seen)
        VALUES (${id}, ${subject}, ${sender}, ${receiver}, ${received_at}, false)
        ON CONFLICT (id) DO NOTHING
      `;
      await client`
        INSERT INTO unread_emails (id, subject, sender, receiver, received_at)
        VALUES (${id}, ${subject}, ${sender}, ${receiver}, ${received_at})
        ON CONFLICT (id) DO NOTHING
      `;
    }

    if (newestHistoryId && newestHistoryId !== startHistoryId) {
      await client`INSERT INTO gmail_history (history_id) VALUES (${newestHistoryId})`;
    }

    let newUnread = [];
    if (addedIds.length) {
      newUnread = await client`
        SELECT *
        FROM unread_emails
        WHERE id = ANY(${addedIds})
        ORDER BY received_at DESC
      `;
    }

    const unreadCount = await client`SELECT COUNT(*) AS c FROM unread_emails`;

    res.json({
      message: "success",
      totalUnread: parseInt(unreadCount[0].c, 10),
      data: newUnread,
    });
  } catch (err) {
    console.error("Delta sync error:", err);
    res.status(500).json({ error: "Failed to fetch new mails" });
  }
});

app.get("/all-mails", async (req, res) => {
  try {
    const page = parseInt(req.query.page || "1", 10);
    const pageSize = 10;
    const offset = (page - 1) * pageSize;

    const mails = await client`
      SELECT * FROM all_mails
      ORDER BY received_at DESC
      LIMIT ${pageSize}
      OFFSET ${offset}
    `;

    const totalRes = await client`SELECT COUNT(*) AS count FROM all_mails`;
    const totalMails = parseInt(totalRes[0].count, 10);

    res.json({
      message: "success",
      page,
      totalPages: Math.ceil(totalMails / pageSize),
      totalMails,
      data: mails,
    });
  } catch (err) {
    console.error("Error fetching all mails:", err);
    res.status(500).json({ error: "Failed to fetch all mails" });
  }
});

app.get("/mail-stats", async (req, res) => {
  try {
    const totalMailRes = await client`SELECT total_count FROM total_mail ORDER BY id DESC LIMIT 1`;
    const totalInbox = totalMailRes.length ? totalMailRes[0].total_count : 0;

    const totalUnreadRes = await client`SELECT COUNT(*) AS count FROM unread_emails`;
    const totalUnread = parseInt(totalUnreadRes[0].count, 10);

    const lastSeenRes = await client`SELECT last_count FROM last_seen_count ORDER BY id DESC LIMIT 1`;
    const currentlyLoadedUnread = lastSeenRes.length ? lastSeenRes[0].last_count : 0;

    res.json({
      totalInbox,
      totalUnread,
      currentlyLoadedUnread,
    });
  } catch (err) {
    console.error("Error fetching mail stats:", err);
    res.status(500).json({ error: "Failed to fetch mail stats" });
  }
});

app.get("/backfill-aug-sept", async (req, res) => {
  try {
    const gmail = google.gmail({ version: "v1", auth: oAuth2Client });
    const query = "after:2025/08/15 before:2025/09/16";
    let nextPageToken;
    let stored = 0;

    do {
      const listRes = await gmail.users.messages.list({
        userId: "me",
        q: query,
        maxResults: 100,
        pageToken: nextPageToken,
      });

      const msgs = listRes.data.messages || [];
      for (const msg of msgs) {
        const msgRes = await gmail.users.messages.get({ userId: "me", id: msg.id });
        const headers = msgRes.data.payload.headers || [];
        const subject = headers.find(h => h.name === "Subject")?.value || "(No Subject)";
        const from = headers.find(h => h.name === "From")?.value || "(Unknown)";
        const to = headers.find(h => h.name === "To")?.value || process.env.MAIL;
        const dateStr = headers.find(h => h.name === "Date")?.value;
        const received_at = dateStr ? new Date(dateStr) : new Date();

        await client`
          INSERT INTO all_mails (id, subject, sender, receiver, received_at, seen)
          VALUES (${msg.id}, ${subject}, ${from}, ${to}, ${received_at}, ${false})
          ON CONFLICT (id) DO NOTHING
        `;
        stored++;
      }

      nextPageToken = listRes.data.nextPageToken;
    } while (nextPageToken);

    const totalRes = await client`SELECT COUNT(*) AS count FROM all_mails`;
    const totalCount = parseInt(totalRes[0].count, 10);
    await client`INSERT INTO total_mail (total_count) VALUES (${totalCount})`;

    res.json({ message: `Backfill complete. Stored ${stored} messages`, totalCount });
  } catch (err) {
    console.error("Backfill error:", err);
    res.status(500).json({ error: "Failed to backfill emails" });
  }
});

app.get("/mail-stats-aug-sept", async (req, res) => {
  try {
    const start = "2025-08-15 00:00:00";
    const end   = "2025-09-15 23:59:59";

    const totalInboxRes = await client`
      SELECT COUNT(*) AS count
      FROM all_mails
      WHERE received_at BETWEEN ${start} AND ${end}
    `;

    const totalUnreadRes = await client`
      SELECT COUNT(*) AS count
      FROM unread_emails
      WHERE received_at BETWEEN ${start} AND ${end}
    `;

    res.json({
      totalInbox: parseInt(totalInboxRes[0].count, 10),
      totalUnread: parseInt(totalUnreadRes[0].count, 10),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch range stats" });
  }
});


app.get("/delta-sync", async (req, res) => {
  try {
    const gmail = google.gmail({ version: "v1", auth: oAuth2Client });
    const hist = await client`SELECT history_id FROM gmail_history ORDER BY id DESC LIMIT 1`;
    let startHistoryId = hist.length ? hist[0].history_id : null;

    if (!startHistoryId) {
      const profile = await gmail.users.getProfile({ userId: "me" });
      await client`INSERT INTO gmail_history (history_id) VALUES (${profile.data.historyId})`;
      return res.json({ message: "Initial sync set. No delta yet." });
    }

    const changes = [];
    let nextPageToken;
    do {
      const resHistory = await gmail.users.history.list({
        userId: "me",
        startHistoryId,
        historyTypes: ["messageAdded", "messageDeleted"],
        pageToken: nextPageToken,
      });

      (resHistory.data.history || []).forEach(h => {
        (h.messagesAdded || []).forEach(m => changes.push({ type: "added", id: m.message.id } ));
        (h.messagesDeleted || []).forEach(m => changes.push({ type: "deleted", id: m.message.id }));
      });

      nextPageToken = resHistory.data.nextPageToken;
      if (resHistory.data.historyId) {
        await client`INSERT INTO gmail_history (history_id) VALUES (${resHistory.data.historyId})`;
      }
    } while (nextPageToken);

    res.json({ message: "success", changes });
  } catch (err) {
    console.error("delta sync error:", err);
    res.status(500).json({ error: "Delta sync failed" });
  }
});

app.get("/logout", async (req, res) => {
  try {
    if (fs.existsSync("token.json")) {
      const token = JSON.parse(fs.readFileSync("token.json", "utf-8"));
      if (token.access_token) await oAuth2Client.revokeToken(token.access_token);
      fs.unlinkSync("token.json");
    }
    res.json({ message: "Logged out and token deleted" });
  } catch (err) {
    console.error("Error during logout:", err);
    res.status(500).json({ error: "Failed to logout" });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
