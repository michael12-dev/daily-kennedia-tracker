require("dotenv").config();
const express = require('express');
const nodemailer = require('nodemailer');
const cors = require('cors');
const path = require('path');
const { execSync } = require('child_process');
const fs = require('fs');
const os = require('os');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ─── CONFIG ──────────────────────────────────────────────────────────────────
const SMTP_CONFIG = {
  host: process.env.SMTP_HOST || 'smtp.gmail.com',
  port: parseInt(process.env.SMTP_PORT || '587'),
  secure: true,
  family: 4,
  auth: {
    user: process.env.SMTP_USER || 'kennediaconsultingtracker@gmail.com',
    pass: process.env.SMTP_PASS || 'zlup ufnk hwwm ytze',
  },
};
const SENDER_NAME = process.env.SENDER_NAME || 'Kennedia Consulting Tracker';
// ─────────────────────────────────────────────────────────────────────────────

function buildExcel(data) {
  const tmpDir = os.tmpdir();
  const dataFile = path.join(tmpDir, `report_data_${Date.now()}.json`);
  const excelFile = path.join(tmpDir, `report_${Date.now()}.xlsx`);

  fs.writeFileSync(dataFile, JSON.stringify({ ...data, excelFile }));
  execSync(`python3 ${path.join(__dirname, 'build_excel.py')} "${dataFile}"`, { timeout: 30000 });

  if (!fs.existsSync(excelFile)) throw new Error('Excel generation failed');

  const buffer = fs.readFileSync(excelFile);
  try { fs.unlinkSync(dataFile); fs.unlinkSync(excelFile); } catch (e) {}
  return buffer;
}

// ─── Send report by email ────────────────────────────────────────────────────
app.post('/api/send-report', async (req, res) => {
  try {
    const { staffName, staffUnit, staffEmail, bossEmail, date, tasks } = req.body;

    if (!bossEmail || !staffName || !staffUnit || !staffEmail) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const excelBuffer = buildExcel({ staffName, staffUnit, date, tasks });
    const safeName = staffName.replace(/[^a-zA-Z0-9]/g, '_');
    const safeDate = date.replace(/[^a-zA-Z0-9]/g, '-');
    const filename = `${safeName}_DailyReport_${safeDate}.xlsx`;

    const transporter = nodemailer.createTransport(SMTP_CONFIG);

    const taskSummary = tasks.map((t, i) =>
      `${i+1}. [${t.priority || 'M'}] ${t.client} — ${t.proposed}\n   Result: ${t.result || t.actual || 'N/A'}\n   Time: ${t.time || 'N/A'}`
    ).join('\n\n');

    const mailOptions = {
      from: `"${staffName} (Kennedia Tracker)" <${SMTP_CONFIG.auth.user}>`,
      replyTo: `"${staffName}" <${staffEmail}>`,
      to: bossEmail,
      cc: staffEmail,
      subject: `Daily Deliverable Report — ${staffName} — ${date}`,
      html: `
        <div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;">
          <div style="background:#548235;padding:20px 28px;border-radius:8px 8px 0 0;">
            <h2 style="color:white;margin:0;font-size:20px;">Kennedia Consulting</h2>
            <p style="color:#d9f0c9;margin:4px 0 0;font-size:13px;">Daily Deliverable Report</p>
          </div>
          <div style="background:#f9fdf5;padding:24px 28px;border:1px solid #c6e0a4;border-top:none;">
            <table style="width:100%;font-size:14px;margin-bottom:16px;">
              <tr><td style="color:#548235;font-weight:bold;padding:4px 0;width:80px;">Staff:</td><td>${staffName}</td></tr>
              <tr><td style="color:#548235;font-weight:bold;padding:4px 0;">Email:</td><td><a href="mailto:${staffEmail}" style="color:#548235;">${staffEmail}</a></td></tr>
              <tr><td style="color:#548235;font-weight:bold;padding:4px 0;">Unit:</td><td>${staffUnit}</td></tr>
              <tr><td style="color:#548235;font-weight:bold;padding:4px 0;">Date:</td><td>${date}</td></tr>
            </table>
            <hr style="border:none;border-top:1px solid #c6e0a4;margin:16px 0;">
            <p style="font-size:13px;color:#375623;font-weight:bold;margin-bottom:12px;">TASKS SUMMARY</p>
            <pre style="font-size:13px;color:#333;background:#fff;padding:16px;border-radius:6px;border:1px solid #e0efd4;white-space:pre-wrap;">${taskSummary}</pre>
            <p style="font-size:12px;color:#888;margin-top:20px;">The full Excel report is attached to this email.</p>
          </div>
        </div>
      `,
      attachments: [{
        filename,
        content: excelBuffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      }],
    };

    await transporter.sendMail(mailOptions);
    res.json({ success: true, filename });
  } catch (err) {
    console.error('Send error:', err);
    res.status(500).json({ error: err.message || 'Failed to send report' });
  }
});

// ─── Download report as file ─────────────────────────────────────────────────
app.post('/api/download-report', async (req, res) => {
  try {
    const { staffName, staffUnit, date, tasks } = req.body;
    const tmpDir = os.tmpdir();
    const dataFile = path.join(tmpDir, `report_data_${Date.now()}.json`);
    const excelFile = path.join(tmpDir, `report_${Date.now()}.xlsx`);

    fs.writeFileSync(dataFile, JSON.stringify({ staffName, staffUnit, date, tasks, excelFile }));
    execSync(`python3 ${path.join(__dirname, 'build_excel.py')} "${dataFile}"`, { timeout: 30000 });

    const safeName = staffName.replace(/[^a-zA-Z0-9]/g, '_');
    const safeDate = date.replace(/[^a-zA-Z0-9]/g, '-');
    const filename = `${safeName}_DailyReport_${safeDate}.xlsx`;

    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.sendFile(excelFile, {}, () => {
      try { fs.unlinkSync(dataFile); fs.unlinkSync(excelFile); } catch (e) {}
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ─── Serve frontend ──────────────────────────────────────────────────────────
app.get('/{*path}', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─── Start server ────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Kennedia Tracker running on http://localhost:${PORT}`));
