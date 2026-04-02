'use strict';

const express = require('express');
const session = require('express-session');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const os = require('os');
const net = require('net');

const { parseExcel } = require('./src/excel.js');
const { generateForms } = require('./src/generate.js');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

const PASSWORD = 'Torechudi0';
const SESSION_SECRET = 'folio-keystone-2026';

// ─── Middleware ────────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true }
}));

// ─── Auth middleware ───────────────────────────────────────────────────────────
function requireAuth(req, res, next) {
  if (req.session && req.session.authenticated) return next();
  res.redirect('/auth');
}

// ─── Auth routes ───────────────────────────────────────────────────────────────
app.get('/auth', (req, res) => {
  res.sendFile(path.join(__dirname, 'renderer', 'password.html'));
});

app.post('/auth/login', (req, res) => {
  const { password } = req.body;
  if (password === PASSWORD) {
    req.session.authenticated = true;
    res.status(200).json({ ok: true });
  } else {
    res.status(401).json({ error: 'Unauthorized' });
  }
});

app.get('/auth/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/auth');
  });
});

// ─── App route ─────────────────────────────────────────────────────────────────
app.get('/', (req, res) => res.redirect('/auth'));

app.get('/app', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, 'renderer', 'index.html'));
});

// ─── API: Universities ─────────────────────────────────────────────────────────
app.get('/universities', requireAuth, (req, res) => {
  res.json({
    universities: [
      { id: 'mokwon', displayName: 'Mokwon University' }
    ]
  });
});

// ─── API: Metrics ──────────────────────────────────────────────────────────────
app.get('/metrics', requireAuth, (req, res) => {
  const memBytes = process.memoryUsage().rss;
  const memMB = (memBytes / 1024 / 1024).toFixed(1);
  res.json({
    version: 'v1.0.0',
    memory: `${memMB} MB`
  });
});

// ─── API: Excel template download ──────────────────────────────────────────────
app.get('/excel-template/:university', requireAuth, (req, res) => {
  const { university } = req.params;
  let templatePath;
  try {
    const electron = require('electron');
    if (electron && electron.app && electron.app.isPackaged) {
      templatePath = path.join(process.resourcesPath, 'assets', `${university}_template.xlsx`);
    } else {
      templatePath = path.join(__dirname, 'assets', `${university}_template.xlsx`);
    }
  } catch (_) {
    templatePath = path.join(__dirname, 'assets', `${university}_template.xlsx`);
  }

  if (!fs.existsSync(templatePath)) {
    return res.status(404).json({ error: 'Template not found' });
  }

  res.download(templatePath, `${university}_template.xlsx`);
});

// ─── API: Generate forms ───────────────────────────────────────────────────────
app.post('/generate', requireAuth, upload.single('file'), async (req, res) => {
  const startTime = Date.now();

  try {
    const { university } = req.body;
    if (!university) return res.status(400).json({ error: 'Missing university' });
    if (!req.file) return res.status(400).json({ error: 'Missing Excel file' });

    // Parse Excel
    let students;
    try {
      students = parseExcel(req.file.buffer);
    } catch (err) {
      return res.status(400).json({ error: `Excel parse error: ${err.message}` });
    }

    if (!students.length) {
      return res.status(400).json({ error: 'No student data found in Excel file' });
    }

    // Determine template path
    let templatePath;
    try {
      const electron = require('electron');
      if (electron && electron.app && electron.app.isPackaged) {
        templatePath = path.join(process.resourcesPath, 'templates', `${university}.docx`);
      } else {
        templatePath = path.join(__dirname, 'templates', `${university}.docx`);
      }
    } catch (_) {
      templatePath = path.join(__dirname, 'templates', `${university}.docx`);
    }

    if (!fs.existsSync(templatePath)) {
      return res.status(500).json({ error: `Template not found: ${university}.docx` });
    }

    // Determine photos directory (next to exe or cwd)
    let photosDir;
    try {
      const electron = require('electron');
      if (electron && electron.app && electron.app.isPackaged) {
        photosDir = path.join(path.dirname(process.execPath), 'photos');
      } else {
        photosDir = path.join(__dirname, 'photos');
      }
    } catch (_) {
      photosDir = path.join(__dirname, 'photos');
    }

    // Output directory
    const outputDir = path.join(os.homedir(), 'Documents', 'Mokwon University');
    fs.mkdirSync(outputDir, { recursive: true });

    // Generate
    const result = await generateForms({
      students,
      templatePath,
      photosDir,
      outputDir,
    });

    const elapsedMs = Date.now() - startTime;

    res.json({
      success: true,
      files: result.files.map(f => ({
        ...f,
        path: `/download/${encodeURIComponent(f.filename)}`
      })),
      generated: result.generated,
      totalRows: students.length,
      elapsedMs,
      warnings: result.warnings,
    });
  } catch (err) {
    console.error('Generate error:', err);
    res.status(500).json({ error: err.message || 'Internal server error' });
  }
});

// ─── API: Download generated file ──────────────────────────────────────────────
app.get('/download/:filename', requireAuth, (req, res) => {
  const { filename } = req.params;
  // Sanitize: no path traversal
  const safeName = path.basename(filename);
  const filePath = path.join(os.homedir(), 'Documents', 'Mokwon University', safeName);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found' });
  }

  res.download(filePath, safeName);
});

// ─── API: Open outputs folder ──────────────────────────────────────────────────
app.post('/open-outputs', requireAuth, (req, res) => {
  const outputDir = path.join(os.homedir(), 'Documents', 'Mokwon University');
  fs.mkdirSync(outputDir, { recursive: true });

  try {
    const { shell } = require('electron');
    shell.openPath(outputDir);
  } catch (_) {
    // Not in Electron context - ignore
  }

  res.json({ ok: true, path: outputDir });
});

// ─── API: Open templates folder ────────────────────────────────────────────────
app.post('/open-templates', requireAuth, (req, res) => {
  let templatesDir;
  try {
    const electron = require('electron');
    if (electron && electron.app && electron.app.isPackaged) {
      templatesDir = path.join(process.resourcesPath, 'templates');
    } else {
      templatesDir = path.join(__dirname, 'templates');
    }
  } catch (_) {
    templatesDir = path.join(__dirname, 'templates');
  }

  try {
    const { shell } = require('electron');
    shell.openPath(templatesDir);
  } catch (_) {}

  res.json({ ok: true, path: templatesDir });
});

// ─── Port finder ───────────────────────────────────────────────────────────────
function findFreePort(start) {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.listen(start, '127.0.0.1', () => {
      const port = server.address().port;
      server.close(() => resolve(port));
    });
    server.on('error', () => {
      findFreePort(start + 1).then(resolve).catch(reject);
    });
  });
}

// ─── Start function ────────────────────────────────────────────────────────────
function start(callback) {
  findFreePort(3721).then((port) => {
    app.listen(port, '127.0.0.1', () => {
      console.log(`Folio server running on http://127.0.0.1:${port}`);
      if (callback) callback(port);
    });
  }).catch((err) => {
    console.error('Failed to find free port:', err);
    process.exit(1);
  });
}

module.exports = { start, app };
