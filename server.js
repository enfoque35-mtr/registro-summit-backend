/**
 * Backend de registro — I Summit Comunicación Política
 * ─────────────────────────────────────────────────────
 * Recibe el POST del formulario, genera un token único + QR,
 * lo guarda en Postgres y envía el QR automáticamente por correo.
 *
 * Instalar dependencias:
 *   npm install express qrcode pg resend dotenv cors
 *
 * Variables de entorno (Railway las inyecta / tú las agregas):
 *   DATABASE_URL=postgresql://...        -> la crea Railway al añadir el addon de Postgres
 *   RESEND_API_KEY=re_xxxxxxxxxxxx        -> cuenta gratis en resend.com
 *   FROM_EMAIL=registro@enfoque35.co      -> debe ser un dominio verificado en Resend
 *   PORT=3000
 *
 * Ejecutar:
 *   node server.js
 *
 * El formulario HTML ya apunta a este endpoint vía REGISTRATION_ENDPOINT:
 *   POST /api/registro   body: { nombre, correo, ocupacion, ciudad, telefono, comentarios, evento }
 */

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const QRCode = require('qrcode');
const { Pool } = require('pg');
const { Resend } = require('resend');
const crypto = require('crypto');
const ExcelJS = require('exceljs');
const { generarCertificado } = require('./certificado');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' })); // más grande de lo normal: aquí llegan las imágenes en base64

const resend = new Resend(process.env.RESEND_API_KEY);
const FROM_EMAIL = process.env.FROM_EMAIL || 'registro@enfoque35.co';
const EVENT_NAME = 'I Summit Comunicación Política';
const EVENT_DATE = 'Viernes 23 de octubre de 2026 · Montería';
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || 'https://registro-summit-backend-production.up.railway.app';
const ADMIN_KEY = process.env.ADMIN_KEY || 'cambia-esta-clave';
// Página pública de la evaluación (el archivo evaluacion-summit.html subido al hosting)
const EVALUACION_URL = process.env.EVALUACION_URL || 'https://www.enfoque35.co/evaluacion-summit/';

// ── Base de datos (Postgres) ─────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes('railway') ? { rejectUnauthorized: false } : false,
});

async function initDb() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS registros (
      id              SERIAL PRIMARY KEY,
      token           TEXT UNIQUE NOT NULL,
      nombre          TEXT NOT NULL,
      correo          TEXT NOT NULL,
      ocupacion       TEXT NOT NULL,
      ciudad          TEXT NOT NULL,
      telefono        TEXT NOT NULL,
      comentarios     TEXT,
      evento          TEXT NOT NULL,
      correo_enviado  BOOLEAN DEFAULT FALSE,
      check_in        BOOLEAN DEFAULT FALSE,
      check_in_at     TIMESTAMP,
      created_at      TIMESTAMP DEFAULT NOW()
    )
  `);
  // Un correo solo puede registrarse una vez por evento
  await pool.query(`
    CREATE UNIQUE INDEX IF NOT EXISTS correo_evento_unico
    ON registros (correo, evento)
  `);
  // Control de pagos (se marca manualmente desde el panel)
  await pool.query(`ALTER TABLE registros ADD COLUMN IF NOT EXISTS pagado BOOLEAN DEFAULT FALSE`);
  await pool.query(`ALTER TABLE registros ADD COLUMN IF NOT EXISTS pagado_at TIMESTAMP`);

  // Evaluación del evento: una respuesta por persona (por token)
  await pool.query(`
    CREATE TABLE IF NOT EXISTS evaluaciones (
      token         TEXT PRIMARY KEY,
      relevancia    TEXT NOT NULL,
      ponentes      TEXT NOT NULL,
      asistiria     TEXT NOT NULL,
      tema          TEXT NOT NULL,
      herramientas  TEXT NOT NULL,
      calificacion  INTEGER NOT NULL,
      created_at    TIMESTAMP DEFAULT NOW()
    )
  `);
}

// ── Validación mínima ────────────────────────────────────────────
function validate(body) {
  const required = ['nombre', 'correo', 'ocupacion', 'ciudad', 'telefono'];
  for (const field of required) {
    if (!body[field] || String(body[field]).trim().length < 2) {
      return `Campo inválido o vacío: ${field}`;
    }
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(body.correo)) {
    return 'Correo electrónico inválido';
  }
  return null;
}

// ── Email con el QR (imagen visible por URL + también adjunto) ───
async function enviarCorreoConQR({ nombre, correo, token }) {
  const qrImageUrl = `${PUBLIC_BASE_URL}/api/qr/${token}`;
  const qrPngBuffer = await QRCode.toBuffer(token, { width: 500, margin: 2 });

  await resend.emails.send({
    from: `${EVENT_NAME} <${FROM_EMAIL}>`,
    to: correo,
    subject: `Tu entrada — ${EVENT_NAME}`,
    html: `
      <div style="font-family: Arial, sans-serif; max-width:480px; margin:0 auto; padding:24px;">
        <p style="color:#0E2A54; font-weight:700; font-size:13px; letter-spacing:.03em; text-transform:uppercase;">
          ${EVENT_NAME}
        </p>
        <h1 style="color:#0E2A54; font-size:22px; margin:8px 0 4px;">¡Registro confirmado, ${nombre}!</h1>
        <p style="color:#4C6270; font-size:14px; line-height:1.6;">${EVENT_DATE}</p>
        <p style="color:#0E2A54; font-weight:700; font-size:16px; line-height:1.6; margin:16px 0;">
          Ten en cuenta que este código solo será válido una vez se realice y valide el pago.
          Para cualquier información o duda contacte al número 310 7089040.
        </p>
        <p style="color:#4C6270; font-size:14px; line-height:1.6;">
          Presenta este código QR el día del evento en el punto de ingreso. Es personal e intransferible.
        </p>
        <div style="text-align:center; margin:28px 0;">
          <img src="${qrImageUrl}" alt="Código QR de acceso" width="220" height="220" style="width:220px; height:220px; display:block; margin:0 auto;" />
        </div>
        <p style="color:#9AA3AC; font-size:12px;">Código: ${token}</p>
      </div>
    `,
    attachments: [
      {
        filename: 'qr-entrada.png',
        content: qrPngBuffer.toString('base64'),
      },
    ],
  });
}

// ── Endpoint principal: registro ─────────────────────────────────
app.post('/api/registro', async (req, res) => {
  const errorMsg = validate(req.body);
  if (errorMsg) return res.status(400).json({ ok: false, error: errorMsg });

  const { nombre, correo, ocupacion, ciudad, telefono, comentarios, evento } = req.body;
  const token = crypto.randomUUID();
  const eventoId = evento || 'summit-comunicacion-politica-2026';

  try {
    await pool.query(
      `INSERT INTO registros (token, nombre, correo, ocupacion, ciudad, telefono, comentarios, evento)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8)`,
      [
        token,
        nombre.trim(),
        correo.trim().toLowerCase(),
        ocupacion.trim(),
        ciudad.trim(),
        telefono.trim(),
        (comentarios || '').trim(),
        eventoId,
      ]
    );
  } catch (e) {
    if (e.code === '23505') {
      return res.status(409).json({ ok: false, error: 'Este correo ya está registrado' });
    }
    console.error(e);
    return res.status(500).json({ ok: false, error: 'Error guardando el registro' });
  }

  try {
    await enviarCorreoConQR({ nombre, correo, token });
    await pool.query(`UPDATE registros SET correo_enviado = TRUE WHERE token = $1`, [token]);
  } catch (e) {
    console.error('Error enviando correo:', e);
    return res.status(200).json({
      ok: true,
      warning: 'Registro guardado, pero el correo no pudo enviarse. Revisa RESEND_API_KEY / FROM_EMAIL.',
    });
  }

  res.status(200).json({ ok: true });
});

// ── Imagen del QR, servida públicamente por token ─────────────────
app.get('/api/qr/:token', async (req, res) => {
  try {
    const buffer = await QRCode.toBuffer(req.params.token, { width: 500, margin: 2 });
    res.setHeader('Content-Type', 'image/png');
    res.setHeader('Cache-Control', 'public, max-age=31536000, immutable');
    res.send(buffer);
  } catch (e) {
    res.status(400).send('Token inválido');
  }
});

// ── Endpoint de verificación el día del evento ───────────────────
app.get('/api/checkin/:token', async (req, res) => {
  const { rows } = await pool.query(`SELECT * FROM registros WHERE token = $1`, [req.params.token]);
  const registro = rows[0];

  if (!registro) {
    return res.status(404).json({ ok: false, status: 'no_encontrado' });
  }
  if (registro.check_in) {
    return res.status(200).json({
      ok: false,
      status: 'ya_usado',
      nombre: registro.nombre,
      check_in_at: registro.check_in_at,
    });
  }

  await pool.query(
    `UPDATE registros SET check_in = TRUE, check_in_at = NOW() WHERE token = $1`,
    [req.params.token]
  );
  res.status(200).json({ ok: true, status: 'valido', nombre: registro.nombre, ocupacion: registro.ocupacion });
});

// ── Panel simple: cuántos registrados / cuántos ya entraron ──────
app.get('/api/resumen', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT evento, COUNT(*) AS registrados, COUNT(*) FILTER (WHERE check_in) AS ingresados
    FROM registros GROUP BY evento
  `);
  res.json(rows);
});

// ── Exportar todos los registros como archivo Excel (.xlsx) ──────
async function exportarExcel(req, res) {
  const { rows } = await pool.query(`
    SELECT r.token, r.nombre, r.correo, r.ocupacion, r.ciudad, r.telefono, r.comentarios,
           CASE WHEN r.pagado THEN 'SI' ELSE 'NO' END AS pagado, r.pagado_at,
           r.correo_enviado, r.check_in, r.check_in_at, r.created_at,
           CASE WHEN e.token IS NULL THEN 'NO' ELSE 'SI' END AS evaluo,
           e.relevancia, e.ponentes, e.asistiria, e.tema, e.herramientas, e.calificacion,
           e.created_at AS evaluacion_at
    FROM registros r
    LEFT JOIN evaluaciones e ON e.token = r.token
    ORDER BY r.created_at
  `);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Registros');

  sheet.columns = [
    { header: 'token', key: 'token', width: 38 },
    { header: 'nombre', key: 'nombre', width: 28 },
    { header: 'correo', key: 'correo', width: 28 },
    { header: 'ocupacion', key: 'ocupacion', width: 22 },
    { header: 'ciudad', key: 'ciudad', width: 18 },
    { header: 'telefono', key: 'telefono', width: 16 },
    { header: 'comentarios', key: 'comentarios', width: 30 },
    { header: 'pagado', key: 'pagado', width: 10 },
    { header: 'pagado_at', key: 'pagado_at', width: 22 },
    { header: 'correo_enviado', key: 'correo_enviado', width: 16 },
    { header: 'check_in', key: 'check_in', width: 12 },
    { header: 'check_in_at', key: 'check_in_at', width: 22 },
    { header: 'created_at', key: 'created_at', width: 22 },
    { header: 'evaluo', key: 'evaluo', width: 10 },
    { header: 'relevancia_temas', key: 'relevancia', width: 24 },
    { header: 'nivel_ponentes', key: 'ponentes', width: 16 },
    { header: 'asistiria_marzo', key: 'asistiria', width: 20 },
    { header: 'tema_sugerido', key: 'tema', width: 36 },
    { header: 'herramientas', key: 'herramientas', width: 16 },
    { header: 'calificacion_summit', key: 'calificacion', width: 20 },
    { header: 'evaluacion_at', key: 'evaluacion_at', width: 22 },
  ];
  sheet.getRow(1).font = { bold: true };

  rows.forEach((row) => sheet.addRow(row));

  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.setHeader('Content-Disposition', 'attachment; filename="registros-summit.xlsx"');

  await workbook.xlsx.write(res);
  res.end();
}

// ── Protección simple para los endpoints de administrador ────────
function requireAdmin(req, res, next) {
  const key = req.headers['x-admin-key'];
  if (!key || key !== ADMIN_KEY) {
    return res.status(401).json({ ok: false, error: 'No autorizado' });
  }
  next();
}

// El Excel contiene datos personales: solo con clave de administrador
// Permite abrir el enlace directo en el navegador: /api/exportar?key=TU_CLAVE
app.get('/api/exportar', (req, res, next) => {
  if (req.query.key && req.query.key === ADMIN_KEY) return exportarExcel(req, res);
  return requireAdmin(req, res, () => exportarExcel(req, res));
});
app.get('/api/admin/exportar', requireAdmin, exportarExcel);

function personalizar(texto, nombre) {
  return (texto || '').replaceAll('{{nombre}}', nombre);
}

// Para el cuerpo: el nombre del registrado se escapa para que no pueda inyectar HTML
function escaparHtml(s) {
  return String(s || '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
// {{enlace}} = enlace personal a la evaluación (con el token de cada registrado)
function personalizarCuerpo(texto, nombre, token) {
  return (texto || '')
    .replaceAll('{{nombre}}', escaparHtml(nombre))
    .replaceAll('{{enlace}}', `${EVALUACION_URL}?t=${encodeURIComponent(token || '')}`);
}

function construirHtmlMensaje(cuerpoPersonalizado, imagenUrl, modoHtml = false) {
  const bloqueImagen = imagenUrl
    ? `<div style="text-align:center; margin:20px 0;">
         <img src="${imagenUrl}" alt="" style="max-width:100%; border-radius:4px;" />
       </div>`
    : '';

  return `
    <div style="font-family: Arial, sans-serif; max-width:480px; margin:0 auto; padding:24px;">
      <p style="color:#0E2A54; font-weight:700; font-size:13px; letter-spacing:.03em; text-transform:uppercase;">
        ${EVENT_NAME}
      </p>
      ${bloqueImagen}
      <div style="color:#10222F; font-size:15px; line-height:1.7; ${modoHtml ? '' : 'white-space:pre-line;'} margin-top:12px;">
        ${cuerpoPersonalizado}
      </div>
    </div>
  `;
}

// Estado del envío masivo en curso (un solo envío a la vez es suficiente aquí)
let broadcastJob = {
  running: false,
  total: 0,
  enviados: 0,
  fallidos: 0,
  iniciado_en: null,
  terminado_en: null,
};

// ── Imágenes para los mensajes masivos (guardadas en memoria) ─────
// Se guardan aquí en vez de en la base de datos porque son temporales:
// solo hacen falta mientras se redactan y envían los mensajes del evento.
const imagenesMensajes = new Map(); // id -> { buffer, mimeType }

app.post('/api/admin/subir-imagen', requireAdmin, (req, res) => {
  const { contenidoBase64, mimeType } = req.body;
  if (!contenidoBase64 || !mimeType) {
    return res.status(400).json({ ok: false, error: 'Falta la imagen o su tipo' });
  }
  if (!mimeType.startsWith('image/')) {
    return res.status(400).json({ ok: false, error: 'El archivo debe ser una imagen' });
  }

  const id = crypto.randomUUID();
  const buffer = Buffer.from(contenidoBase64, 'base64');

  if (buffer.length > 5 * 1024 * 1024) {
    return res.status(400).json({ ok: false, error: 'La imagen no debe superar 5 MB' });
  }

  imagenesMensajes.set(id, { buffer, mimeType });
  res.json({ ok: true, id, url: `${PUBLIC_BASE_URL}/api/admin/imagen/${id}` });
});

// Servida SIN clave de administrador: los clientes de correo de cada
// registrado necesitan poder cargarla al abrir el mensaje.
app.get('/api/admin/imagen/:id', (req, res) => {
  const imagen = imagenesMensajes.get(req.params.id);
  if (!imagen) return res.status(404).send('Imagen no encontrada');
  res.setHeader('Content-Type', imagen.mimeType);
  res.setHeader('Cache-Control', 'public, max-age=604800');
  res.send(imagen.buffer);
});

// ── Lista de registrados para el panel de administrador ───────────
app.get('/api/admin/registros', requireAdmin, async (req, res) => {
  const { rows } = await pool.query(`
    SELECT id, nombre, correo, ocupacion, ciudad, telefono, check_in, pagado
    FROM registros ORDER BY created_at
  `);
  res.json(rows);
});

// ── Marcar / desmarcar pagos ────────────────────────────────────────
app.post('/api/admin/marcar-pago', requireAdmin, async (req, res) => {
  const { ids, pagado } = req.body || {};
  if (!Array.isArray(ids) || ids.length === 0 || typeof pagado !== 'boolean') {
    return res.status(400).json({ ok: false, error: 'Datos inválidos' });
  }
  const idsNum = ids.map(Number).filter(Number.isInteger);
  const r = await pool.query(
    `UPDATE registros
     SET pagado = $1::boolean, pagado_at = CASE WHEN $1::boolean THEN NOW() ELSE NULL END
     WHERE id = ANY($2::int[])`,
    [pagado, idsNum]
  );
  res.json({ ok: true, actualizados: r.rowCount });
});

// ── Agregar asistentes que no se pre-registraron (para poder enviarles su enlace) ──
app.post('/api/admin/importar-asistentes', requireAdmin, async (req, res) => {
  const lista = Array.isArray(req.body?.asistentes) ? req.body.asistentes : [];
  if (lista.length === 0 || lista.length > 500) {
    return res.status(400).json({ ok: false, error: 'Lista vacía o demasiado grande' });
  }
  let creados = 0;
  for (const a of lista) {
    const correo = String(a.correo || '').trim().toLowerCase();
    const nombre = String(a.nombre || '').replace(/\s+/g, ' ').trim();
    if (!nombre || !/^\S+@\S+\.\S+$/.test(correo)) continue;
    const r = await pool.query(
      `INSERT INTO registros (token, nombre, correo, ocupacion, ciudad, telefono, evento, correo_enviado, check_in, check_in_at)
       VALUES ($1,$2,$3,$4,$5,$6,$7, FALSE, TRUE, NOW())
       ON CONFLICT (correo, evento) DO NOTHING`,
      [crypto.randomUUID(), nombre, correo, 'Asistente (registro en sitio)', '', String(a.telefono || '').trim(), 'summit-comunicacion-politica-2026']
    );
    creados += r.rowCount;
  }
  res.json({ ok: true, creados, omitidos: lista.length - creados });
});

// ── Estado del envío masivo (para la barra de progreso) ────────────
app.get('/api/admin/estado-envio', requireAdmin, (req, res) => {
  res.json(broadcastJob);
});

// ── Enviar mensaje personalizado a los registrados seleccionados ──
app.post('/api/admin/enviar-mensaje', requireAdmin, async (req, res) => {
  const { asunto, mensaje, ids, imagenUrl, modoHtml } = req.body;
  if (!asunto || !mensaje) {
    return res.status(400).json({ ok: false, error: 'Falta el asunto o el mensaje' });
  }
  if (broadcastJob.running) {
    return res.status(409).json({ ok: false, error: 'Ya hay un envío en curso, espera a que termine' });
  }

  // Si "ids" viene con una lista de IDs, solo se envía a esos. Si no viene o viene vacía, se envía a todos.
  const { rows } =
    Array.isArray(ids) && ids.length > 0
      ? await pool.query(`SELECT nombre, correo, token FROM registros WHERE id = ANY($1)`, [ids])
      : await pool.query(`SELECT nombre, correo, token FROM registros`);

  if (rows.length === 0) {
    return res.status(400).json({ ok: false, error: 'No hay destinatarios para enviar' });
  }

  broadcastJob = {
    running: true,
    total: rows.length,
    enviados: 0,
    fallidos: 0,
    iniciado_en: new Date().toISOString(),
    terminado_en: null,
  };

  res.status(200).json({ ok: true, mensaje: `Envío iniciado a ${rows.length} personas` });

  // El envío corre en segundo plano; el panel consulta el progreso con /api/admin/estado-envio
  (async () => {
    const TAMANO_LOTE = 90; // el envío por lotes de Resend acepta hasta 100 por llamada
    for (let i = 0; i < rows.length; i += TAMANO_LOTE) {
      const lote = rows.slice(i, i + TAMANO_LOTE);
      const correos = lote.map((r) => ({
        from: `${EVENT_NAME} <${FROM_EMAIL}>`,
        to: r.correo,
        subject: personalizar(asunto, r.nombre),
        html: construirHtmlMensaje(personalizarCuerpo(mensaje, r.nombre, r.token), imagenUrl, modoHtml === true),
      }));

      try {
        await resend.batch.send(correos);
        broadcastJob.enviados += lote.length;
      } catch (e) {
        console.error('Error enviando lote:', e);
        broadcastJob.fallidos += lote.length;
      }
    }
    broadcastJob.running = false;
    broadcastJob.terminado_en = new Date().toISOString();
  })();
});

// ── Evaluación del evento (pública, protegida por el token de cada persona) ──
const OPCIONES = {
  relevancia: ['Nada relevantes', 'Poco relevantes', 'Medianamente relevantes', 'Muy relevantes', 'Totalmente relevantes'],
  ponentes: ['Deficiente', 'Regular', 'Bueno', 'Muy bueno', 'Excelente'],
  asistiria: ['Sí, definitivamente', 'Probablemente sí', 'No estoy seguro', 'Probablemente no', 'No'],
  herramientas: ['Ninguna', 'Pocas', 'Algunas', 'Varias', 'Muchas'],
};
const TEMAS = [
  'Campañas electorales y estrategia política',
  'Marketing político digital y redes sociales',
  'Comunicación de gobierno y gestión pública',
  'Manejo de crisis y reputación',
  'Oratoria y vocería',
  'Inteligencia artificial y datos en la política',
];

// ¿Quién es y ya evaluó? (la página lo usa al abrirse)
app.get('/api/evaluacion/:token', async (req, res) => {
  const { rows } = await pool.query(
    `SELECT r.nombre, (e.token IS NOT NULL) AS respondio
     FROM registros r LEFT JOIN evaluaciones e ON e.token = r.token
     WHERE r.token = $1`,
    [req.params.token]
  );
  if (rows.length === 0) return res.status(404).json({ ok: false, error: 'Enlace no válido' });
  res.json({ ok: true, nombre: rows[0].nombre, respondio: rows[0].respondio });
});

app.post('/api/evaluacion/:token', async (req, res) => {
  const { rows } = await pool.query(`SELECT 1 FROM registros WHERE token = $1`, [req.params.token]);
  if (rows.length === 0) return res.status(404).json({ ok: false, error: 'Enlace no válido' });

  const b = req.body || {};
  const tema = String(b.tema || '').trim();
  const temaOk = TEMAS.includes(tema) || (tema.startsWith('Otro: ') && tema.length > 6 && tema.length <= 160);
  const cal = Number(b.calificacion);
  const valido =
    Object.keys(OPCIONES).every((k) => OPCIONES[k].includes(b[k])) &&
    temaOk && Number.isInteger(cal) && cal >= 1 && cal <= 5;
  if (!valido) return res.status(400).json({ ok: false, error: 'Respuestas incompletas o no válidas' });

  // Si ya había respondido, se conserva la primera respuesta
  await pool.query(
    `INSERT INTO evaluaciones (token, relevancia, ponentes, asistiria, tema, herramientas, calificacion)
     VALUES ($1,$2,$3,$4,$5,$6,$7) ON CONFLICT (token) DO NOTHING`,
    [req.params.token, b.relevancia, b.ponentes, b.asistiria, tema, b.herramientas, cal]
  );
  res.json({ ok: true });
});

// Certificado en PDF: solo si la persona ya respondió la evaluación
app.get('/api/certificado/:token', async (req, res) => {
  const { rows } = await pool.query(
    `SELECT r.nombre FROM registros r
     JOIN evaluaciones e ON e.token = r.token
     WHERE r.token = $1`,
    [req.params.token]
  );
  if (rows.length === 0) {
    return res.status(403).send('Primero debes responder la evaluación del evento para descargar tu certificado.');
  }
  try {
    const pdf = await generarCertificado(rows[0].nombre);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="certificado-i-summit.pdf"');
    res.send(pdf);
  } catch (e) {
    console.error('Error generando certificado:', e);
    res.status(500).send('No se pudo generar el certificado. Intenta de nuevo.');
  }
});

const PORT = process.env.PORT || 3000;
initDb()
  .then(() => app.listen(PORT, () => console.log(`Backend de registro escuchando en puerto ${PORT}`)))
  .catch((e) => {
    console.error('No se pudo conectar a la base de datos:', e);
    process.exit(1);
  });
