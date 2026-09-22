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

const app = express();
app.use(cors());
app.use(express.json());

const resend = new Resend(process.env.RESEND_API_KEY);
const FROM_EMAIL = process.env.FROM_EMAIL || 'registro@enfoque35.co';
const EVENT_NAME = 'I Summit Comunicación Política';
const EVENT_DATE = 'Viernes 23 de octubre de 2026 · Montería';
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || 'https://registro-summit-backend-production.up.railway.app';

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

// ── Email con el QR (como imagen servida desde una URL pública) ──
async function enviarCorreoConQR({ nombre, correo, token }) {
  const qrImageUrl = `${PUBLIC_BASE_URL}/api/qr/${token}`;

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
app.get('/api/exportar', async (req, res) => {
  const { rows } = await pool.query(`
    SELECT token, nombre, correo, ocupacion, ciudad, telefono, comentarios,
           correo_enviado, check_in, check_in_at, created_at
    FROM registros ORDER BY created_at
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
    { header: 'correo_enviado', key: 'correo_enviado', width: 16 },
    { header: 'check_in', key: 'check_in', width: 12 },
    { header: 'check_in_at', key: 'check_in_at', width: 22 },
    { header: 'created_at', key: 'created_at', width: 22 },
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
});

const PORT = process.env.PORT || 3000;
initDb()
  .then(() => app.listen(PORT, () => console.log(`Backend de registro escuchando en puerto ${PORT}`)))
  .catch((e) => {
    console.error('No se pudo conectar a la base de datos:', e);
    process.exit(1);
  });
