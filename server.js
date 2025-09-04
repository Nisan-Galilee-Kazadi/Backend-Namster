import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import sharp from 'sharp';
import xlsx from 'xlsx';
import mammoth from 'mammoth';
import archiver from 'archiver';
import crypto from 'crypto';
import cors from 'cors';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = process.env.PORT || 3000;

// CORS for frontend hosted on another domain (e.g., Netlify)
// If you know your exact frontend origin, replace '*' with that origin for tighter security.
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  maxAge: 86400,
}));
// Explicitly handle preflight for all API routes
app.options('/api/*', cors());

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

const publicDir = path.join(__dirname, 'public');
const uploadsDir = path.join(__dirname, 'uploads');
const workDir = path.join(__dirname, 'work');

for (const d of [publicDir, uploadsDir, workDir]) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
}

app.use(express.static(publicDir));

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadsDir);
  },
  filename: function (req, file, cb) {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + '-' + file.originalname.replace(/\s+/g, '_'));
  },
});
const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
});

// In-memory session store
const sessions = new Map();

function newSession() {
  const id = crypto.randomUUID();
  sessions.set(id, {
    id,
    createdAt: Date.now(),
    modelPath: null,
    listPath: null,
    names: [],
    cleanup: [],
  });
  return id;
}

function cleanupSession(id) {
  const s = sessions.get(id);
  if (!s) return;
  for (const p of s.cleanup) {
    try {
      if (p && fs.existsSync(p)) fs.rmSync(p, { recursive: true, force: true });
    } catch {}
  }
  sessions.delete(id);
}

function parseNamesFromText(text) {
  // Split by newlines/semicolons/commas, trim, filter empties
  const raw = text
    .split(/\r?\n|;|,/)
    .map((t) => t.trim())
    .filter(Boolean);
  // Filter headers/parasites like "Liste"
  const cleaned = raw.filter((t) => {
    const norm = t.toLowerCase().replace(/\s+/g, '');
    if (norm === 'liste') return false;
    return true;
  });
  return cleaned;
}

async function extractNamesFromFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.csv') {
    const content = fs.readFileSync(filePath, 'utf8');
    return parseNamesFromText(content);
  }
  if (ext === '.xlsx' || ext === '.xls') {
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1 });
    // Flatten cells row-wise
    const cells = rows.flat().filter((v) => v !== null && v !== undefined);
    return cells.map((c) => String(c)).map((s) => s.trim()).filter(Boolean);
  }
  if (ext === '.docx') {
    const { value } = await mammoth.extractRawText({ path: filePath });
    return parseNamesFromText(value || '');
  }
  if (ext === '.pdf') {
    try {
      const dataBuffer = fs.readFileSync(filePath);
      // Lazy import to avoid initialization issues with certain distributions
      const { default: pdfParse } = await import('pdf-parse');
      const data = await pdfParse(dataBuffer);
      return parseNamesFromText(data.text || '');
    } catch (e) {
      console.warn('PDF parsing failed, continuing without PDF support:', e?.message || e);
      return [];
    }
  }
  // Fallback: try reading as text
  try {
    const content = fs.readFileSync(filePath, 'utf8');
    return parseNamesFromText(content);
  } catch {
    return [];
  }
}

function buildSVGText({ text, x, y, fontFamily, fontSize, color, width, height }) {
  const safeText = (text || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const ff = fontFamily || 'Arial';
  const fsPx = Number(fontSize) || 48;
  const fill = color || '#000000';
  return `<?xml version="1.0" encoding="UTF-8"?>\n<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">\n  <style>\n    .t { font-family: '${ff}'; font-size: ${fsPx}px; fill: ${fill}; dominant-baseline: hanging; }\n  </style>\n  <text x="${x}" y="${y}" class="t">${safeText}</text>\n</svg>`;
}

async function composeImageWithText(modelPath, outPath, options) {
  const { x, y, fontFamily, fontSize, color, text } = options;
  const img = sharp(modelPath);
  const meta = await img.metadata();
  const width = meta.width || 2000;
  const height = meta.height || 1000;
  const svg = buildSVGText({ text, x, y, fontFamily, fontSize, color, width, height });
  const buffer = await img.composite([{ input: Buffer.from(svg), top: 0, left: 0 }]).png().toBuffer();
  await fs.promises.writeFile(outPath, buffer);
}

app.post('/api/upload', upload.fields([
  { name: 'model', maxCount: 1 },
  { name: 'list', maxCount: 1 },
]), async (req, res) => {
  try {
    const sid = newSession();
    const s = sessions.get(sid);
    const model = req.files['model']?.[0];
    const list = req.files['list']?.[0];
    if (!model || !list) {
      return res.status(400).json({ error: 'Model image and list file are required.' });
    }
    s.modelPath = model.path;
    s.listPath = list.path;
    s.cleanup.push(model.path, list.path);

    const names = await extractNamesFromFile(s.listPath);
    s.names = names;
    console.log(`[upload] session=${sid} names_total=${names.length}`);

    res.json({ sessionId: sid, namesPreview: names.slice(0, 50), namesTotal: names.length });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Upload failed.' });
  }
});

app.post('/api/test', express.json(), async (req, res) => {
  try {
    const { sessionId, x, y, fontFamily, fontSize, color } = req.body || {};
    const s = sessions.get(sessionId);
    if (!s) return res.status(400).json({ error: 'Invalid session' });
    if (!s.modelPath || !s.names?.length) return res.status(400).json({ error: 'Missing model or names' });

    const firstName = s.names[0];
    const outDir = path.join(workDir, sessionId);
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    const outPath = path.join(outDir, 'test.png');

    await composeImageWithText(s.modelPath, outPath, { text: firstName, x, y, fontFamily, fontSize, color });
    s.cleanup.push(outDir);

    const data = fs.readFileSync(outPath);
    const base64 = 'data:image/png;base64,' + data.toString('base64');
    res.json({ preview: base64 });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Test render failed.' });
  }
});

app.post('/api/generate', express.json(), async (req, res) => {
  try {
    const { sessionId, x, y, fontFamily, fontSize, color, format, offset, limit } = req.body || {};
    const s = sessions.get(sessionId);
    if (!s) return res.status(400).json({ error: 'Invalid session' });
    if (!s.modelPath || !s.names?.length) return res.status(400).json({ error: 'Missing model or names' });

    const outDir = path.join(workDir, sessionId, 'all');
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

    const ext = (format === 'jpg' || format === 'jpeg') ? 'jpg' : 'png';

    // Optional batching
    const start = Math.max(0, Number.isFinite(Number(offset)) ? Number(offset) : 0);
    const maxBatch = 50;
    const requested = Number.isFinite(Number(limit)) ? Number(limit) : undefined;
    const batchSize = requested ? Math.max(1, Math.min(maxBatch, requested)) : undefined;
    const endExclusive = batchSize ? Math.min(s.names.length, start + batchSize) : s.names.length;
    console.log(`[generate] session=${sessionId} total=${s.names.length} start=${start} limit_req=${requested ?? 'all'} will_process=${endExclusive - start}`);

    for (let idx = start; idx < endExclusive; idx++) {
      const name = s.names[idx];
      const filename = `${String(idx + 1).padStart(3, '0')}-${name.replace(/[^a-z0-9_-]+/gi, '_')}.${ext}`;
      const outPath = path.join(outDir, filename);

      if (ext === 'png') {
        await composeImageWithText(s.modelPath, outPath, { text: name, x, y, fontFamily, fontSize, color });
      } else {
        // compose to PNG then convert to JPG
        const tmpPng = outPath.replace(/\.jpg$/, '.png');
        await composeImageWithText(s.modelPath, tmpPng, { text: name, x, y, fontFamily, fontSize, color });
        const jpgBuf = await sharp(tmpPng).jpeg({ quality: 90 }).toBuffer();
        await fs.promises.writeFile(outPath, jpgBuf);
        fs.unlinkSync(tmpPng);
      }
    }

    // Zip
    const zipPath = path.join(workDir, sessionId, 'invitations.zip');
    await new Promise((resolve, reject) => {
      const output = fs.createWriteStream(zipPath);
      const archive = archiver('zip', { zlib: { level: 9 } });
      output.on('close', resolve);
      archive.on('error', reject);
      archive.pipe(output);
      archive.directory(outDir, false);
      archive.finalize();
    });
    s.cleanup.push(path.join(workDir, sessionId));

    res.json({ downloadUrl: `/api/download/${sessionId}`, processed: endExclusive - start, offset: start, total: s.names.length });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Generation failed.' });
  }
});

app.get('/api/download/:sid', (req, res) => {
  const sid = req.params.sid;
  const s = sessions.get(sid);
  if (!s) return res.status(400).json({ error: 'Invalid session' });
  const zipPath = path.join(workDir, sid, 'invitations.zip');
  if (!fs.existsSync(zipPath)) return res.status(404).json({ error: 'ZIP not found' });
  res.download(zipPath, 'invitations.zip', (err) => {
    // Cleanup after download attempt
    cleanupSession(sid);
  });
});

app.get('/health', (req, res) => res.json({ ok: true }));

app.listen(port, () => {
  console.log(`Namster server running at http://localhost:${port}`);
});
