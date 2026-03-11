const express   = require("express");
const multer    = require("multer");
const cors      = require("cors");
const path      = require("path");
const fs        = require("fs");
const { spawn } = require("child_process");

const app  = express();
const PORT = 3001;

app.use(cors());
app.use(express.json({ limit: "500mb" }));
app.use(express.urlencoded({ limit: "500mb", extended: true }));

const UPLOAD_DIR = path.join(__dirname, "uploads_temp");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename:    (req, file, cb) => {
    const nombre = Buffer.from(file.originalname, "latin1").toString("utf8");
    cb(null, nombre);
  },
});

// Limite de 500MB por archivo
const upload = multer({
  storage,
  limits: { fileSize: 500 * 1024 * 1024 }
});

app.post("/api/upload", (req, res) => {
  upload.array("files")(req, res, (err) => {
    if (err) {
      console.error("Error subiendo archivos:", err.message);
      return res.status(400).json({ ok: false, error: err.message });
    }
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ ok: false, error: "No se recibieron archivos" });
    }
    console.log(`Archivos subidos: ${req.files.length}`);
    res.json({ ok: true, rutas: req.files.map(f => f.path), carpeta: UPLOAD_DIR });
  });
});

app.post("/api/run", (req, res) => {
  const { scriptName, params } = req.body;

  const scriptPath = path.join(__dirname, scriptName);
  if (!fs.existsSync(scriptPath)) {
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.write(`ERROR: No se encontro el script: ${scriptName}\n`);
    res.write(`Asegurate de que el archivo .py este en: ${__dirname}\n`);
    res.end();
    return;
  }

  // Solo borrar archivos _run_ temporales anteriores, NO los xlsx de entrada
  fs.readdirSync(UPLOAD_DIR)
    .filter(f => f.startsWith("_run_") && f.endsWith(".py"))
    .forEach(f => { try { fs.unlinkSync(path.join(UPLOAD_DIR, f)); } catch(_){} });

  let script = fs.readFileSync(scriptPath, "utf8");

  const paramsAbs = { ...params };
  ["CARPETA_PDF", "CARPETA_EXCEL", "ruta_carpeta"].forEach(k => {
    if (paramsAbs[k] !== undefined) paramsAbs[k] = UPLOAD_DIR;
  });
  ["ruta_entrada", "ruta_pagos", "ruta_crp", "ruta_pac"].forEach(k => {
    if (paramsAbs[k] && !path.isAbsolute(paramsAbs[k]))
      paramsAbs[k] = path.join(UPLOAD_DIR, path.basename(paramsAbs[k]));
  });
  ["NOMBRE_EXCEL", "archivo_salida", "ruta_destino", "ruta_salida"].forEach(k => {
    if (paramsAbs[k])
      paramsAbs[k] = path.join(UPLOAD_DIR, path.basename(paramsAbs[k]));
  });

  Object.entries(paramsAbs).forEach(([k, v]) => {
    const regex = new RegExp(`^(${k}\\s*=\\s*).*$`, "m");
    const escaped = String(v).replace(/\\/g, "\\\\");
    if (regex.test(script)) script = script.replace(regex, `${k} = r"${escaped}"`);
  });

  const tempScript = path.join(UPLOAD_DIR, `_run_${Date.now()}.py`);
  fs.writeFileSync(tempScript, script, "utf8");

  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.setHeader("Transfer-Encoding", "chunked");

  // Timeout 10 minutos
  req.setTimeout(600000);
  res.setTimeout(600000);

  let excelDetectado = null;
  let stdoutAcum = "";

  const proc = spawn("python", [tempScript], {
    env: process.env,
    cwd: __dirname,
    windowsHide: true
  });

  proc.stdout.on("data", d => {
    const t = d.toString("utf8");
    stdoutAcum += t;
    console.log(t);
    res.write(t);

    const mx = stdoutAcum.match(/__EXCEL__:([^\r\n]+\.xlsx)/i);
    if (mx) excelDetectado = path.basename(mx[1].trim());

    if (!excelDetectado) {
      const m = stdoutAcum.match(/LISTO[^\r\n]*[>]\s*(.+\.xlsx)/i);
      if (m) excelDetectado = path.basename(m[1].trim());
    }
  });

  proc.stderr.on("data", d => {
    const t = d.toString("utf8");
    console.error(t);
    res.write("⚠️ " + t);
  });

  proc.on("close", code => {
    if (!excelDetectado) {
      const nuevos = fs.readdirSync(UPLOAD_DIR)
        .filter(f => f.endsWith(".xlsx") && !f.startsWith("~$") && !f.startsWith("_"));
      if (nuevos.length > 0) excelDetectado = nuevos[0];
    }

    if (excelDetectado) {
      res.write(`\n${"=".repeat(50)}\n`);
      res.write(`Excel generado: ${excelDetectado}\n`);
      res.write(`\n__EXCEL__:${excelDetectado}`);
    } else {
      res.write(`\n✅ Proceso terminado con codigo: ${code}`);
    }
    res.end();
    try { fs.unlinkSync(tempScript); } catch (_) {}
  });

  proc.on("error", err => {
    res.write(`\nError al ejecutar Python: ${err.message}\n`);
    res.end();
  });
});

app.get("/api/download/:filename", (req, res) => {
  const filename = decodeURIComponent(req.params.filename);
  const fp = path.join(UPLOAD_DIR, filename);
  if (!fs.existsSync(fp)) {
    res.status(404).json({ error: "Archivo no encontrado: " + filename });
    return;
  }
  res.download(fp, filename);
});

app.get("/api/excels", (req, res) => {
  try {
    const files = fs.readdirSync(UPLOAD_DIR)
      .filter(f => f.endsWith(".xlsx") && !f.startsWith("~$") && !f.startsWith("_"));
    res.json(files);
  } catch(_) { res.json([]); }
});

app.delete("/api/files", (req, res) => {
  try {
    fs.readdirSync(UPLOAD_DIR)
      .filter(f => !f.startsWith("_"))
      .forEach(f => { try { fs.unlinkSync(path.join(UPLOAD_DIR, f)); } catch(_){} });
    res.json({ ok: true });
  } catch (e) { res.json({ ok: false, error: e.message }); }
});

app.listen(PORT, () => {
  console.log(`\nServidor FDL USME 2026 -> http://localhost:${PORT}`);
  console.log(`Archivos en: ${UPLOAD_DIR}\n`);
});
