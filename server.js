require("dotenv").config();
const express = require('express');
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
const BREVO_API_KEY = process.env.BREVO_API_KEY;
const SENDER_EMAIL  = process.env.SMTP_USER   || 'kennediaconsultingtracker@gmail.com';
const SENDER_NAME   = process.env.SENDER_NAME || 'Kennedia Consulting Tracker';
const BOSS_EMAIL    = 'justice.okafor@kennediaconsulting.net';
// ─────────────────────────────────────────────────────────────────────────────

function buildExcel(data) {
  const tmpDir    = os.tmpdir();
  const dataFile  = path.join(tmpDir, `report_data_${Date.now()}.json`);
  const excelFile = path.join(tmpDir, `report_${Date.now()}.xlsx`);

  fs.writeFileSync(dataFile, JSON.stringify({ ...data, excelFile }));
  execSync(`python3 ${path.join(__dirname, 'build_excel.py')} "${dataFile}"`, { timeout: 30000 });

  if (!fs.existsSync(excelFile)) throw new Error('Excel generation failed');

  const buffer = fs.readFileSync(excelFile);
  try { fs.unlinkSync(dataFile); fs.unlinkSync(excelFile); } catch (e) {}
  return buffer;
}

async function sendViaBrevo({ to, cc, replyTo, subject, html, attachment, filename }) {
  const payload = {
    sender:  { name: SENDER_NAME, email: SENDER_EMAIL },
    to:      [{ email: to }],
    replyTo: { email: replyTo },
    subject,
    htmlContent: html,
    attachment: [{
      name:    filename,
      content: attachment.toString('base64'),
    }],
  };

  if (cc) payload.cc = [{ email: cc }];

  const response = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      'accept':       'application/json',
      'content-type': 'application/json',
      'api-key':      BREVO_API_KEY,
    },
    body: JSON.stringify(payload),
  });

  const result = await response.json();
  if (!response.ok) throw new Error(result.message || 'Brevo send failed');
  return result;
}

// ─── Send report by email ────────────────────────────────────────────────────
app.post('/api/send-report', async (req, res) => {
  try {
    const { staffName, staffUnit, staffEmail, date, tasks, session } = req.body;

    if (!staffName || !staffUnit || !staffEmail) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const sessionLabel = session === 'morning' ? '🌅 Morning' : '🌙 Evening';
    const excelBuffer  = buildExcel({ staffName, staffUnit, date, tasks, session });
    const safeName     = staffName.replace(/[^a-zA-Z0-9]/g, '_');
    const safeDate     = date.replace(/[^a-zA-Z0-9]/g, '-');
    const sessTag      = session === 'morning' ? 'Morning' : 'Evening';
    const filename     = `${safeName}_${sessTag}Report_${safeDate}.xlsx`;

    const headerBg = session === 'morning' ? '#548235' : '#4f46e5';
    const eveningCols = `
      <th style="background:${headerBg};color:white;padding:8px 10px;text-align:left;font-size:11px;">ACTUAL DELIVERABLE</th>
      <th style="background:${headerBg};color:white;padding:8px 10px;text-align:left;font-size:11px;">ACHIEVEMENT/RESULT</th>`;

    const taskRows = tasks.map((t, i) => {
      const bg       = i % 2 === 0 ? '#ffffff' : '#f0f7e9';
      const actualBg = session === 'evening' ? '#eef2ff' : bg;
      return `<tr>
        <td style="padding:8px 10px;border-bottom:1px solid #c6e0a4;background:${bg};">${i+1}</td>
        <td style="padding:8px 10px;border-bottom:1px solid #c6e0a4;background:${bg};">${t.client}</td>
        <td style="padding:8px 10px;border-bottom:1px solid #c6e0a4;background:${bg};">${t.proposed}</td>
        <td style="padding:8px 10px;border-bottom:1px solid #c6e0a4;background:${bg};text-align:center;">${t.priority||'M'}</td>
        <td style="padding:8px 10px;border-bottom:1px solid #c6e0a4;background:${bg};">${t.time||'—'}</td>
        <td style="padding:8px 10px;border-bottom:1px solid #c6e0a4;background:${actualBg};">${t.actual||'—'}</td>
        <td style="padding:8px 10px;border-bottom:1px solid #c6e0a4;background:${actualBg};">${t.result||'—'}</td>
      </tr>`;
    }).join('');

    const sessionBadge = session === 'morning'
      ? `<span style="background:#fffbeb;color:#92400e;border:1px solid #fcd34d;border-radius:20px;padding:3px 12px;font-size:12px;font-weight:700;">🌅 Morning Submission</span>`
      : `<span style="background:#eef2ff;color:#3730a3;border:1px solid #a5b4fc;border-radius:20px;padding:3px 12px;font-size:12px;font-weight:700;">🌙 Evening Submission</span>`;

    const html = `
      <div style="font-family:Calibri,Arial,sans-serif;max-width:680px;margin:0 auto;">
        <div style="background:${headerBg};padding:20px 28px;border-radius:8px 8px 0 0;">
          <h2 style="color:white;margin:0;font-size:20px;">Kennedia Consulting</h2>
          <p style="color:#d9f0c9;margin:6px 0 0;font-size:13px;">Daily Deliverable Report &nbsp;•&nbsp; ${sessionLabel}</p>
        </div>
        <div style="background:#f9fdf5;padding:20px 28px;border:1px solid #c6e0a4;border-top:none;">
          <div style="margin-bottom:14px;">${sessionBadge}</div>
          <table style="width:100%;font-size:14px;margin-bottom:16px;">
            <tr><td style="color:#548235;font-weight:bold;padding:4px 0;width:80px;">Staff:</td><td>${staffName}</td></tr>
            <tr><td style="color:#548235;font-weight:bold;padding:4px 0;">Email:</td><td><a href="mailto:${staffEmail}" style="color:#548235;">${staffEmail}</a></td></tr>
            <tr><td style="color:#548235;font-weight:bold;padding:4px 0;">Unit:</td><td>${staffUnit}</td></tr>
            <tr><td style="color:#548235;font-weight:bold;padding:4px 0;">Date:</td><td>${date}</td></tr>
          </table>
          <hr style="border:none;border-top:1px solid #c6e0a4;margin:16px 0;">
          <p style="font-size:13px;color:#375623;font-weight:bold;margin-bottom:10px;">TASKS SUMMARY</p>
          <div style="overflow-x:auto;">
            <table style="width:100%;border-collapse:collapse;font-size:12.5px;min-width:600px;">
              <thead><tr>
                <th style="background:${headerBg};color:white;padding:8px 10px;text-align:left;font-size:11px;">S/N</th>
                <th style="background:${headerBg};color:white;padding:8px 10px;text-align:left;font-size:11px;">CLIENT/FOCUS</th>
                <th style="background:${headerBg};color:white;padding:8px 10px;text-align:left;font-size:11px;">PROPOSED TASK</th>
                <th style="background:${headerBg};color:white;padding:8px 10px;text-align:left;font-size:11px;">PRI</th>
                <th style="background:${headerBg};color:white;padding:8px 10px;text-align:left;font-size:11px;">TIME</th>
                ${eveningCols}
              </tr></thead>
              <tbody>${taskRows}</tbody>
            </table>
          </div>
          <p style="font-size:12px;color:#888;margin-top:20px;">The full Excel report is attached to this email.</p>
        </div>
      </div>
    `;

    await sendViaBrevo({
      to:         BOSS_EMAIL,
      cc:         staffEmail,
      replyTo:    staffEmail,
      subject:    `${sessTag} Deliverable Report — ${staffName} — ${date}`,
      html,
      attachment: excelBuffer,
      filename,
    });

    res.json({ success: true, filename });
  } catch (err) {
    console.error('Send error:', err);
    res.status(500).json({ error: err.message || 'Failed to send report' });
  }
});

// ─── Download report as file ─────────────────────────────────────────────────
app.post('/api/download-report', async (req, res) => {
  try {
    const { staffName, staffUnit, date, tasks, session } = req.body;
    const tmpDir    = os.tmpdir();
    const dataFile  = path.join(tmpDir, `report_data_${Date.now()}.json`);
    const excelFile = path.join(tmpDir, `report_${Date.now()}.xlsx`);

    fs.writeFileSync(dataFile, JSON.stringify({ staffName, staffUnit, date, tasks, session, excelFile }));
    execSync(`python3 ${path.join(__dirname, 'build_excel.py')} "${dataFile}"`, { timeout: 30000 });

    const safeName = staffName.replace(/[^a-zA-Z0-9]/g, '_');
    const safeDate = date.replace(/[^a-zA-Z0-9]/g, '-');
    const sessTag  = session === 'morning' ? 'Morning' : 'Evening';
    const filename = `${safeName}_${sessTag}Report_${safeDate}.xlsx`;

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
