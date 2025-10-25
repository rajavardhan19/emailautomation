const express = require('express');
const multer = require('multer');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { parse } = require('csv-parse/sync');

require('dotenv').config();
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');

const app = express();
app.disable('x-powered-by');
app.use(helmet());

// Ensure uploads dir exists
const UPLOAD_DIR = process.env.UPLOAD_DIR || 'uploads';
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// Limit requests to prevent abuse
const limiter = rateLimit({ windowMs: 60 * 1000, max: 500 }); // 500 requests / minute
app.use(limiter);

// Multer with file size limit (e.g. 20MB)
const upload = multer({ dest: UPLOAD_DIR + '/', limits: { fileSize: 20 * 1024 * 1024 } });

// Load SMTP credentials from environment variables (use .env)
const SMTP_USER = process.env.SMTP_USER;
const SMTP_PASS = process.env.SMTP_PASS;
if (!SMTP_USER || !SMTP_PASS) {
  console.error('ERROR: SMTP_USER and SMTP_PASS must be set in a .env file or environment. Exiting.');
  process.exit(1);
}

const DEFAULT_SUBJECT = 'NCET NOTES - easyEngineers';
const DEFAULT_BODY = 'Hi {name},\n\nThis is an automated message sent via the email automation script By easyEngineers Formly Know As NCET NOTES 🙏🏻.';

app.use(express.static('.'));
app.use(express.urlencoded({ extended: true }));

app.get('/', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Bulk Email Sender - NCET Notes</title>
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
        }
        .container {
          background: white;
          border-radius: 20px;
          box-shadow: 0 20px 60px rgba(0,0,0,0.3);
          padding: 40px;
          max-width: 600px;
          width: 100%;
        }
        h2 {
          color: #667eea;
          text-align: center;
          margin-bottom: 10px;
          font-size: 28px;
        }
        .subtitle {
          text-align: center;
          color: #666;
          margin-bottom: 30px;
          font-size: 14px;
        }
        .form-section {
          margin-bottom: 25px;
        }
        .form-section label {
          display: block;
          color: #667eea;
          font-weight: 600;
          margin-bottom: 10px;
          font-size: 14px;
        }
        .upload-area {
          border: 2px dashed #667eea;
          border-radius: 10px;
          padding: 30px;
          text-align: center;
          background: #f8f9ff;
          transition: all 0.3s ease;
        }
        .upload-area:hover {
          border-color: #764ba2;
          background: #f0f2ff;
        }
        input[type="file"] {
          display: none;
        }
        .file-label {
          cursor: pointer;
          color: #667eea;
          font-weight: 600;
          font-size: 16px;
        }
        .file-label:hover {
          color: #764ba2;
        }
        .file-name {
          margin-top: 10px;
          color: #666;
          font-size: 13px;
        }
        textarea {
          width: 100%;
          min-height: 150px;
          padding: 15px;
          border: 2px solid #e0e0e0;
          border-radius: 10px;
          font-family: 'Courier New', monospace;
          font-size: 13px;
          resize: vertical;
          transition: border-color 0.3s;
        }
        textarea:focus {
          outline: none;
          border-color: #667eea;
          box-shadow: 0 0 5px rgba(102, 126, 234, 0.3);
        }
        button {
          width: 100%;
          padding: 15px;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          border: none;
          border-radius: 10px;
          font-size: 18px;
          font-weight: 600;
          cursor: pointer;
          transition: transform 0.2s, box-shadow 0.2s;
          margin-top: 20px;
        }
        button:hover {
          transform: translateY(-2px);
          box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }
        button:active {
          transform: translateY(0);
        }
        .emoji {
          font-size: 40px;
          margin-bottom: 10px;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📤 Bulk Email Sender</h2>
        <div class="subtitle">NCET Notes</div>
        <form method="post" enctype="multipart/form-data" action="/send" id="uploadForm">
          <div class="form-section">
            <label>📁 Choose Excel/CSV File (Name & Email columns required)</label>
            <div class="upload-area">
              <div class="emoji">�</div>
              <label for="fileInput" class="file-label">
                Click to choose file
              </label>
              <input type="file" id="fileInput" name="file" accept=".csv,.xlsx,.xls" required>
              <div class="file-name" id="fileName"></div>
            </div>
          </div>

          <div class="form-section">
            <label>📝 Paste HTML Code (Optional - for email body)</label>
            <textarea name="htmlContent" id="htmlContent" placeholder="Paste your HTML code here... (if not provided, default message will be used)"></textarea>
          </div>

          <button type="submit">🚀 Send Emails</button>
        </form>
      </div>
      <script>
        document.getElementById('fileInput').addEventListener('change', function(e) {
          const fileName = e.target.files[0]?.name || '';
          document.getElementById('fileName').textContent = fileName ? '✓ ' + fileName : '';
        });
            </script>
    </body>
    </html>
  `);
});

app.post('/send', upload.single('file'), async (req, res) => {
  let log = '';
  if (!req.file) {
    const errMsg = 'No file uploaded.';
    if (req.xhr) return res.status(400).json({ error: errMsg });
    return res.status(400).send(errMsg);
  }
  let rows = [];
  const ext = path.extname(req.file.originalname).toLowerCase();
  const htmlContent = req.body.htmlContent || '';  // Get pasted HTML from form
  try {
    if (ext === '.csv') {
      const content = fs.readFileSync(req.file.path, 'utf8');
      rows = parse(content, { columns: true, skip_empty_lines: true });
    } else {
      const wb = xlsx.readFile(req.file.path);
      const ws = wb.Sheets[wb.SheetNames[0]];
      rows = xlsx.utils.sheet_to_json(ws);
    }
  } catch (e) {
    fs.unlinkSync(req.file.path);
    return res.send('Failed to read file: ' + e.message);
  }
  // Detect columns
  function findCol(cols, candidates) {
    for (const cand of candidates) {
      const found = cols.find(c => c.toLowerCase().trim() === cand.toLowerCase().trim());
      if (found) return found;
    }
    return null;
  }
  const columns = rows.length ? Object.keys(rows[0]) : [];
  const emailCol = findCol(columns, ['Email', 'To']);
  const nameCol = findCol(columns, ['Name', 'Full Name', 'FullName']);
  const attachCol = findCol(columns, ['Attachment', 'File', 'FilePath', 'Path', 'File Path']);

  // Setup nodemailer
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });

  for (const row of rows) {
    const to = row[emailCol] || '';
    const name = row[nameCol] || '';
    const subject = DEFAULT_SUBJECT.replace('{name}', name);
    const body = DEFAULT_BODY.replace('{name}', name);
    const attachmentPath = row[attachCol] || '';
    let mailOptions = {
      from: SMTP_USER,
      to,
      subject,
      text: body,
    };
    let htmlBodySet = false;

    // Priority 1: Use pasted HTML content
    if (htmlContent && htmlContent.trim()) {
      // Replace template variables like {{to_name}}, {{email}}, {{name}}, etc.
      let processedHtml = htmlContent;
      processedHtml = processedHtml.replace(/{{to_name}}/g, name);
      processedHtml = processedHtml.replace(/{{email}}/g, to);
      processedHtml = processedHtml.replace(/{{name}}/g, name);
      // Replace any column name from Excel
      for (const col in row) {
        const placeholder = new RegExp(`{{${col}}}`, 'g');
        processedHtml = processedHtml.replace(placeholder, row[col] || '');
      }
      mailOptions.html = processedHtml;
      htmlBodySet = true;
    }
    // Priority 2: Use HTML file from attachment column
    else if (attachmentPath && fs.existsSync(attachmentPath)) {
      const ext = path.extname(attachmentPath).toLowerCase();
      if (ext === '.html' || ext === '.htm') {
        let fileHtmlContent = fs.readFileSync(attachmentPath, 'utf8');
        // Replace template variables
        fileHtmlContent = fileHtmlContent.replace(/{{to_name}}/g, name);
        fileHtmlContent = fileHtmlContent.replace(/{{email}}/g, to);
        fileHtmlContent = fileHtmlContent.replace(/{{name}}/g, name);
        for (const col in row) {
          const placeholder = new RegExp(`{{${col}}}`, 'g');
          fileHtmlContent = fileHtmlContent.replace(placeholder, row[col] || '');
        }
        mailOptions.html = fileHtmlContent;
        htmlBodySet = true;
      } else {
        mailOptions.attachments = [{ filename: path.basename(attachmentPath), path: attachmentPath }];
      }
    }

    // Fallback: Use default body
    if (!htmlBodySet) {
      mailOptions.html = `<pre>${body}</pre>`;
    }

    try {
      await transporter.sendMail(mailOptions);
      log += `✅ Sent to ${to}\n`;
    } catch (e) {
      log += `❌ Failed to send to ${to}: ${e.message}\n`;
    }
    await new Promise(r => setTimeout(r, 1500));
  }
  fs.unlinkSync(req.file.path);
  // If XHR request, return JSON log
  if (req.xhr) return res.json({ log });
  // Otherwise render the HTML result page (kept for backward compatibility)
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Bulk Email Sender - NCET Notes</title>
    </head>
    <body>
      <pre>${log}</pre>
      <p><a href="/">Back</a></p>
    </body>
    </html>
  `);
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log('Server running on http://localhost:' + PORT));
