// Genera el certificado en PDF escribiendo el nombre sobre la plantilla
const fs = require('fs');
const path = require('path');
const { PDFDocument, rgb } = require('pdf-lib');
const fontkit = require('@pdf-lib/fontkit');

const FONDO = fs.readFileSync(path.join(__dirname, 'fondo-certificado.jpg'));
const FUENTE = fs.readFileSync(path.join(__dirname, 'Montserrat-Bold.woff'));

// Posición del nombre en la plantilla (carta horizontal 792 x 612 pt)
const CENTRO_X = 396;
const BASE_Y = 612 - 275;      // línea base, entre "Certifica que" y "Asistió y participó del:"
const ANCHO_MAX = 560;
const TAM_MAX = 30;
const TAM_MIN = 16;
const COLOR = rgb(9 / 255, 27 / 255, 78 / 255);

// "JOSE PEREZ" o "jose perez" -> "Jose Perez"; si ya viene mezclado se respeta
function formatearNombre(n) {
  let s = String(n || '').replace(/\s+/g, ' ').trim();
  if (s === s.toUpperCase() || s === s.toLowerCase()) {
    s = s.toLowerCase().replace(/(^|\s|-)(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
    s = s.replace(/\b(De|Del|La|Las|Los|Y)\b/g, (m, w, off) => (off === 0 ? w : w.toLowerCase()));
  }
  return s;
}

async function generarCertificado(nombre) {
  const pdf = await PDFDocument.create();
  pdf.registerFontkit(fontkit);
  const fuente = await pdf.embedFont(FUENTE, { subset: true });
  const page = pdf.addPage([792, 612]);
  const fondo = await pdf.embedJpg(FONDO);
  page.drawImage(fondo, { x: 0, y: 0, width: 792, height: 612 });

  const texto = formatearNombre(nombre);
  let tam = TAM_MAX;
  while (tam > TAM_MIN && fuente.widthOfTextAtSize(texto, tam) > ANCHO_MAX) tam -= 1;
  const ancho = fuente.widthOfTextAtSize(texto, tam);

  page.drawText(texto, { x: CENTRO_X - ancho / 2, y: BASE_Y, size: tam, font: fuente, color: COLOR });
  return Buffer.from(await pdf.save());
}

module.exports = { generarCertificado, formatearNombre };
