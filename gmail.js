import express from "express";
import multer from "multer";
import nodemailer from "nodemailer";
import cors from "cors";
import dotenv from "dotenv";
import { neon } from "@neondatabase/serverless";
import readXlsxFile from "read-excel-file/node";

dotenv.config();
const app = express();
const client = neon(process.env.db_url);

app.use(cors({ origin: "http://localhost:3000" }));
app.use(express.json());

const upload = multer({ dest: "uploads/" });

app.post("/send-email", upload.array("attachments"), async (req, res) => {
  try {
    const { to, subject, message } = req.body;

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: process.env.MAIL,
        pass: process.env.PASS,
      },
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

    res.json({ message: " Email sent successfully!" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: " Failed to send email" });
  }
});

app.post("/send-excel-emails", upload.single("excel"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded" });
    }

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
      auth: {
        user: process.env.MAIL,
        pass: process.env.PASS,
      },
    });

    let sentCount = 0;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const to = row[emailIdx];
      const subject = subjectIdx !== -1 ? row[subjectIdx] : "No Subject";
      const message = messageIdx !== -1 ? row[messageIdx] : "";

      if (!to) continue;

      await transporter.sendMail({
        to,
        subject,
        text: message,
      });

      await client`
        INSERT INTO sent_emails (receiver, subject, message)
        VALUES (${to}, ${subject}, ${message})
      `;

      sentCount++;
    }

    res.json({ message: `${sentCount} emails sent successfully!` });
  } catch (err) {
    console.error("Error sending excel emails:", err);
    res.status(500).json({ error: " Failed to process Excel emails" });
  }
});

app.listen(5000, () => console.log(" Server running on port 5000"));
