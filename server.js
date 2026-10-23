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
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' })); // más grande de lo normal: aquí llegan las imágenes en base64

const resend = new Resend(process.env.RESEND_API_KEY);
const FROM_EMAIL = process.env.FROM_EMAIL || 'registro@enfoque35.co';
const EVENT_NAME = 'I Summit Comunicación Política';
const EVENT_DATE = 'Viernes 23 de octubre de 2026 · Montería';
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || 'https://registro-summit-backend-production.up.railway.app';
const ADMIN_KEY = process.env.ADMIN_KEY || 'cambia-esta-clave';

// ── Base de datos (Postgres) ─────────────────────────────────────
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes('railway') ? { rejectUnauthorized: false } : false,
});


// ── Plantillas de mensajes (guardadas en la base de datos) ───────
const QR_URL = `${PUBLIC_BASE_URL}/api/qr/{{token}}`;

const PLANTILLAS_INICIALES = [
  {
    nombre: '1. Último llamado de pago (no pagados)',
    audiencia: 'nopagados',
    asunto: '⏰ Últimas horas para tu cupo en el I Summit de Comunicación Política',
    mensaje: `<div style="background-color:#d62828;padding:10px;text-align:center;font-size:15px;font-weight:bold;color:#ffffff;border-radius:4px;">
  🚨 ¡ARRANCA MAÑANA! 🚨
</div>

<p>Hola <strong>{{nombre}}</strong> 👋,</p>

<p>⏰ El <strong>I SUMMIT DE COMUNICACIÓN POLÍTICA</strong> arranca <strong>mañana</strong> y estamos <strong>cerrando la lista de asistentes</strong>.</p>

<p>Tu nombre sigue en nuestro radar y no queremos que te quedes por fuera. Si todavía no has pagado, 🔥 <strong>hoy es el último día para asegurar tu cupo</strong>.</p>

<div style="background-color:#f0f5fb;border-left:4px solid #1f3a5f;padding:14px 18px;margin:20px 0;border-radius:4px;">
  <p style="margin:0 0 8px 0;">💳 <strong>Datos para tu pago</strong></p>
  <strong>Bancolombia, cuenta de ahorros</strong><br>
  <strong>Número: 56922771541</strong><br>
  <strong>Titular: Katia Milena Polo Herrera</strong><br>
  <strong>C.C. 1067843531</strong>
</div>

<p>📲 Cuando lo hagas, <strong>envía el comprobante al WhatsApp +57 310 7089040</strong> y te confirmamos tu entrada de inmediato. Si tienes cualquier duda, <strong>escríbenos a ese mismo número</strong>.</p>

<p>Será una jornada de <strong>ideas, estrategia y contactos de primer nivel</strong> con quienes hacen <strong>comunicación política</strong> en serio. ⚠️ <strong>Esas conversaciones no se repiten</strong>, y queremos que estés en ellas.</p>

<p>✅ Si ya realizaste tu pago, gracias, <strong>nos vemos mañana</strong>.</p>

<p><strong>¡Te esperamos!</strong> 🙌<br>
Equipo I Summit de Comunicación Política</p>`,
  },
  {
    nombre: '2. Recordatorio: mañana es el día (pagados, con QR)',
    audiencia: 'pagados',
    asunto: '🎉 ¡Mañana es el día! Tu cupo en el I Summit de Comunicación Política está confirmado',
    mensaje: `<div style="background-color:#1e7b4a;padding:10px;text-align:center;font-size:15px;font-weight:bold;color:#ffffff;border-radius:4px;">
  🎉 ¡MAÑANA ES EL DÍA! 🎉
</div>

<p>Hola <strong>{{nombre}}</strong> 👋,</p>

<p>✅ <strong>Tu pago está confirmado</strong> y tu cupo en el <strong>I SUMMIT DE COMUNICACIÓN POLÍTICA</strong> está asegurado. ¡Gracias por hacer parte!</p>

<p>⏰ Te recordamos que el evento es <strong>mañana</strong>. Estos son los datos para que no te falte nada:</p>

<div style="background-color:#eaf5ee;border-left:4px solid #1e7b4a;padding:14px 18px;margin:20px 0;border-radius:4px;">
  📅 <strong>Fecha:</strong> Mañana, viernes 23 de octubre<br>
  🕐 <strong>Hora:</strong> 8:00 a. m.<br>
  📍 <strong>Lugar:</strong> Auditorio Cultural de la Universidad de Córdoba, Montería
</div>

<div style="text-align:center;border:2px dashed #1e7b4a;border-radius:8px;padding:20px;margin:24px 0;">
  <p style="margin:0 0 12px 0;">🎟️ <strong>Tu código QR de ingreso</strong></p>
  <img src="${QR_URL}" width="220" height="220" alt="Código QR de ingreso" style="width:220px;height:220px;display:block;margin:0 auto;">
  <p style="margin:12px 0 0 0;font-size:14px;color:#444;"><strong>Preséntalo en la entrada</strong> desde tu celular o impreso. Es personal e intransferible.</p>
</div>

<p>Te recomendamos <strong>llegar con tiempo</strong> para el registro y para que no te pierdas nada del arranque.</p>

<p>Será una jornada de <strong>ideas, estrategia y contactos de primer nivel</strong> con quienes hacen <strong>comunicación política</strong> en serio. 🚀</p>

<p>📲 Si tienes cualquier duda, <strong>escríbenos al WhatsApp +57 310 7089040</strong>.</p>

<p><strong>¡Te esperamos mañana!</strong> 🙌<br>
Equipo I Summit de Comunicación Política</p>`,
  },
  {
    nombre: '3. Llegó el gran día (pagados, con QR)',
    audiencia: 'pagados',
    asunto: '🎉 ¡Llegó el gran día! Hoy es el I Summit de Comunicación Política',
    mensaje: `<div style="background-color:#1e7b4a;padding:12px;text-align:center;font-size:16px;font-weight:bold;color:#ffffff;border-radius:4px;">
  🎉 ¡LLEGÓ EL GRAN DÍA! 🎉
</div>

<p>Buenos días, <strong>{{nombre}}</strong> ☀️,</p>

<p>¡Hoy es el día! 🙌 El <strong>I SUMMIT DE COMUNICACIÓN POLÍTICA</strong> comienza en pocos minutos y tu lugar ya está reservado. Gracias por ser parte de este primer encuentro.</p>

<div style="background-color:#eaf5ee;border-left:4px solid #1e7b4a;padding:14px 18px;margin:20px 0;border-radius:4px;">
  📅 <strong>Fecha:</strong> Hoy, viernes 23 de octubre<br>
  🕐 <strong>Hora:</strong> 8:00 a. m.<br>
  📍 <strong>Lugar:</strong> Auditorio Cultural de la Universidad de Córdoba, Montería
</div>

<div style="text-align:center;border:2px dashed #1e7b4a;border-radius:8px;padding:20px;margin:24px 0;">
  <p style="margin:0 0 12px 0;">🎟️ <strong>Tu código QR de ingreso</strong></p>
  <img src="${QR_URL}" width="220" height="220" alt="Código QR de ingreso" style="width:220px;height:220px;display:block;margin:0 auto;">
  <p style="margin:12px 0 0 0;font-size:14px;color:#444;"><strong>Tenlo listo en tu celular</strong> para mostrarlo en la entrada. Es personal e intransferible.</p>
</div>

<p><strong>Al llegar:</strong></p>
<p>
✅ Ten tu QR a la mano (sube el brillo de la pantalla)<br>
✅ Lleva tu documento de identidad<br>
✅ Pasa primero por el registro 🚶‍♂️
</p>

<p>Será una jornada de <strong>ideas, estrategia y contactos de primer nivel</strong> con quienes hacen <strong>comunicación política</strong> en serio. 🚀</p>

<p>📲 ¿Alguna duda de último momento? <strong>Escríbenos al WhatsApp +57 310 7089040</strong>.</p>

<p><strong>¡Nos vemos muy pronto!</strong> 🙌<br>
Equipo I Summit de Comunicación Política</p>`,
  },
];

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

  // Plantillas de mensajes
  await pool.query(`
    CREATE TABLE IF NOT EXISTS plantillas (
      id          SERIAL PRIMARY KEY,
      nombre      TEXT UNIQUE NOT NULL,
      audiencia   TEXT NOT NULL DEFAULT 'todos',
      asunto      TEXT NOT NULL,
      mensaje     TEXT NOT NULL,
      updated_at  TIMESTAMP DEFAULT NOW()
    )
  `);
  // Solo se cargan las plantillas iniciales si la tabla está vacía
  const { rows: cnt } = await pool.query(`SELECT COUNT(*)::int AS n FROM plantillas`);
  if (cnt[0].n === 0) {
    for (const t of PLANTILLAS_INICIALES) {
      await pool.query(
        `INSERT INTO plantillas (nombre, audiencia, asunto, mensaje) VALUES ($1, $2, $3, $4)`,
        [t.nombre, t.audiencia, t.asunto, t.mensaje]
      );
    }
  }
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
    SELECT token, nombre, correo, ocupacion, ciudad, telefono, comentarios,
           CASE WHEN pagado THEN 'SI' ELSE 'NO' END AS pagado, pagado_at,
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
    { header: 'pagado', key: 'pagado', width: 10 },
    { header: 'pagado_at', key: 'pagado_at', width: 22 },
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
function personalizarCuerpo(texto, nombre, token) {
  return (texto || '')
    .replaceAll('{{nombre}}', escaparHtml(nombre))
    .replaceAll('{{token}}', token || '');
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
  ultimo_error: null,
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
    ultimo_error: null,
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
        // El SDK de Resend NO lanza excepción cuando falla: devuelve { data, error }
        const { error } = await resend.batch.send(correos);
        if (error) throw new Error(`${error.name || 'error'}: ${error.message || JSON.stringify(error)}`);
        broadcastJob.enviados += lote.length;
      } catch (e) {
        console.error('Error enviando lote:', e);
        broadcastJob.fallidos += lote.length;
        broadcastJob.ultimo_error = String(e.message || e);
      }
    }
    broadcastJob.running = false;
    broadcastJob.terminado_en = new Date().toISOString();
  })();
});

// ── Página de plantillas (se entra con la clave de administrador) ──
app.get('/plantillas', (req, res) => {
  res.sendFile(path.join(__dirname, 'plantillas.html'));
});

app.get('/api/admin/plantillas', requireAdmin, async (req, res) => {
  const { rows } = await pool.query(
    `SELECT id, nombre, audiencia, asunto, mensaje, updated_at FROM plantillas ORDER BY id`
  );
  res.json(rows);
});

// Crea una plantilla nueva (sin id) o actualiza una existente (con id)
app.post('/api/admin/plantillas', requireAdmin, async (req, res) => {
  const { id, nombre, audiencia, asunto, mensaje } = req.body || {};
  const audiencias = ['todos', 'pagados', 'nopagados'];
  if (!nombre || !asunto || !mensaje || !audiencias.includes(audiencia)) {
    return res.status(400).json({ ok: false, error: 'Faltan datos de la plantilla' });
  }
  try {
    if (id) {
      const r = await pool.query(
        `UPDATE plantillas SET nombre=$1, audiencia=$2, asunto=$3, mensaje=$4, updated_at=NOW()
         WHERE id=$5 RETURNING id`,
        [nombre.trim(), audiencia, asunto, mensaje, Number(id)]
      );
      if (r.rowCount === 0) return res.status(404).json({ ok: false, error: 'Plantilla no encontrada' });
      return res.json({ ok: true, id: r.rows[0].id });
    }
    const r = await pool.query(
      `INSERT INTO plantillas (nombre, audiencia, asunto, mensaje) VALUES ($1,$2,$3,$4) RETURNING id`,
      [nombre.trim(), audiencia, asunto, mensaje]
    );
    res.json({ ok: true, id: r.rows[0].id });
  } catch (e) {
    if (e.code === '23505') {
      return res.status(409).json({ ok: false, error: 'Ya existe una plantilla con ese nombre' });
    }
    console.error(e);
    res.status(500).json({ ok: false, error: 'No se pudo guardar la plantilla' });
  }
});

app.delete('/api/admin/plantillas/:id', requireAdmin, async (req, res) => {
  const r = await pool.query(`DELETE FROM plantillas WHERE id = $1`, [Number(req.params.id)]);
  res.json({ ok: true, eliminadas: r.rowCount });
});

// ── Enviar a una lista cargada (nombre + correo), sin depender de la tabla ──
// Valida cada correo, descarta los malos y, si Resend rechaza un lote,
// reintenta uno por uno para que un solo correo malo no frene a los demás.
app.post('/api/admin/enviar-lista', requireAdmin, async (req, res) => {
  const { asunto, mensaje, destinatarios, modoHtml } = req.body || {};
  if (!asunto || !mensaje) {
    return res.status(400).json({ ok: false, error: 'Falta el asunto o el mensaje' });
  }
  if (!Array.isArray(destinatarios) || destinatarios.length === 0) {
    return res.status(400).json({ ok: false, error: 'La lista está vacía' });
  }
  if (broadcastJob.running) {
    return res.status(409).json({ ok: false, error: 'Ya hay un envío en curso, espera a que termine' });
  }

  const correoValido = /^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)+$/;
  const vistos = new Set();
  const lista = [];
  const descartados = [];
  for (const d of destinatarios) {
    const correo = String((d && d.correo) || '').trim().toLowerCase();
    const nombre = String((d && d.nombre) || '').trim() || 'amigo/a';
    if (!correoValido.test(correo)) { descartados.push({ nombre, correo, motivo: 'Correo inválido' }); continue; }
    if (vistos.has(correo)) { descartados.push({ nombre, correo, motivo: 'Repetido' }); continue; }
    vistos.add(correo);
    lista.push({ nombre, correo });
  }
  if (lista.length === 0) {
    return res.status(400).json({ ok: false, error: 'Ningún correo de la lista es válido', descartados });
  }

  broadcastJob = {
    running: true,
    total: lista.length,
    enviados: 0,
    fallidos: 0,
    iniciado_en: new Date().toISOString(),
    terminado_en: null,
    ultimo_error: null,
  };
  res.status(200).json({ ok: true, mensaje: `Envío iniciado a ${lista.length} personas`, descartados });

  (async () => {
    const TAMANO_LOTE = 90;
    const armar = (r) => ({
      from: `${EVENT_NAME} <${FROM_EMAIL}>`,
      to: r.correo,
      subject: personalizar(asunto, r.nombre),
      html: construirHtmlMensaje(personalizarCuerpo(mensaje, r.nombre, ''), null, modoHtml === true),
    });
    for (let i = 0; i < lista.length; i += TAMANO_LOTE) {
      const lote = lista.slice(i, i + TAMANO_LOTE);
      try {
        const { error } = await resend.batch.send(lote.map(armar));
        if (error) throw new Error(`${error.name || 'error'}: ${error.message || JSON.stringify(error)}`);
        broadcastJob.enviados += lote.length;
      } catch (e) {
        console.error('Lote rechazado, se reintenta uno por uno:', e);
        for (const r of lote) {
          try {
            const { error } = await resend.emails.send(armar(r));
            if (error) throw new Error(`${error.name || 'error'}: ${error.message || JSON.stringify(error)}`);
            broadcastJob.enviados += 1;
          } catch (e2) {
            broadcastJob.fallidos += 1;
            broadcastJob.ultimo_error = `${r.correo}: ${String(e2.message || e2)}`;
          }
        }
      }
    }
    broadcastJob.running = false;
    broadcastJob.terminado_en = new Date().toISOString();
  })().catch((e) => {
    console.error('Error inesperado en envío de lista:', e);
    broadcastJob.running = false;
    broadcastJob.ultimo_error = String(e.message || e);
  });
});

const PORT = process.env.PORT || 3000;
initDb()
  .then(() => app.listen(PORT, () => console.log(`Backend de registro escuchando en puerto ${PORT}`)))
  .catch((e) => {
    console.error('No se pudo conectar a la base de datos:', e);
    process.exit(1);
  });
