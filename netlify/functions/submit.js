'use strict';
const Busboy = require('busboy');

// ── Constantes ───────────────────────────────────────────────────────────────
const GRAPH      = 'https://graph.microsoft.com/v1.0';
const MAX_BYTES  = 8 * 1024 * 1024; // 8 MB por imagen
const IMG_FOLDER = 'Formulario_Fotos'; // carpeta raíz en OneDrive

// Firmas de archivos de imagen válidos (magic bytes)
const MAGIC = [
  { sig: [0xff, 0xd8, 0xff],                           mime: 'image/jpeg' },
  { sig: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], mime: 'image/png'  },
  { sig: [0x52, 0x49, 0x46, 0x46],  extra: { offset: 8, bytes: [0x57,0x45,0x42,0x50] }, mime: 'image/webp' },
];

function detectImageMime(buf) {
  for (const { sig, extra, mime } of MAGIC) {
    if (sig.every((b, i) => buf[i] === b)) {
      if (!extra) return mime;
      const eb = extra.bytes;
      if (eb.every((b, i) => buf[extra.offset + i] === b)) return mime;
    }
  }
  // HEIC / HEIF: bytes 4-7 == 'ftyp'
  if (buf.length >= 8) {
    const ftyp = [0x66, 0x74, 0x79, 0x70];
    if (ftyp.every((b, i) => buf[4 + i] === b)) return 'image/heic';
  }
  return null;
}

// ── Microsoft Graph helpers ──────────────────────────────────────────────────
async function getToken() {
  const { GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET } = process.env;
  const resp = await fetch(
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     GRAPH_CLIENT_ID,
        client_secret: GRAPH_CLIENT_SECRET,
        scope:         'https://graph.microsoft.com/.default',
      }).toString(),
    }
  );
  const d = await resp.json();
  if (!d.access_token) throw new Error(`Auth error: ${d.error_description || JSON.stringify(d)}`);
  return d.access_token;
}

function shareId(url) {
  // Base64url del URL (sin padding)
  return 'u!' + Buffer.from(url).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function resolveDrive(token, shareUrl) {
  const r = await fetch(
    `${GRAPH}/shares/${shareId(shareUrl)}/driveItem?$select=id,parentReference`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!r.ok) throw new Error(`Resolve share failed ${r.status}: ${await r.text()}`);
  const item = await r.json();
  return { driveId: item.parentReference.driveId, itemId: item.id };
}

async function readSheet(token, driveId, itemId, sheet) {
  const enc = encodeURIComponent(sheet);
  const r = await fetch(
    `${GRAPH}/drives/${driveId}/items/${itemId}/workbook/worksheets/${enc}/usedRange(valuesOnly=true)`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!r.ok) return { headers: [], rows: [] };
  const { values = [] } = await r.json();
  if (!values.length) return { headers: [], rows: [] };
  return { headers: values[0].map(h => String(h ?? '').trim()), rows: values };
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function serialize(v) {
  if (v === null || v === undefined || v === '') return '';
  return String(v);
}

// Normaliza DNI a solo dígitos (réplica de _clean_dni en Python)
function cleanDni(v) {
  return String(v ?? '').split('.')[0].replace(/\D/g, '');
}

// Normaliza patente a mayúsculas sin espacios
function cleanPat(v) {
  return String(v ?? '').toUpperCase().replace(/\s/g, '');
}

async function patchRange(token, driveId, itemId, sheet, addr, rowValues) {
  const enc = encodeURIComponent(sheet);
  const r = await fetch(
    `${GRAPH}/drives/${driveId}/items/${itemId}/workbook/worksheets/${enc}/range(address='${addr}')`,
    {
      method:  'PATCH',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body:    JSON.stringify({ values: [rowValues] }),
    }
  );
  if (!r.ok) throw new Error(`Patch ${addr} failed ${r.status}: ${await r.text()}`);
}

// Actualiza fila existente (lectura + escritura en bloque = 2 llamadas a Graph)
async function updateRow(token, driveId, itemId, sheet, rowNum, updates, headers) {
  const enc    = encodeURIComponent(sheet);
  const nCols  = headers.length;
  const addr   = `A${rowNum}:${colLetter(nCols)}${rowNum}`;

  // 1 — leer fila actual
  const r = await fetch(
    `${GRAPH}/drives/${driveId}/items/${itemId}/workbook/worksheets/${enc}/range(address='${addr}')`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  let currentVals = r.ok ? (await r.json()).values?.[0] ?? [] : [];
  while (currentVals.length < nCols) currentVals.push('');

  // 2 — aplicar actualizaciones
  for (const [col, val] of Object.entries(updates)) {
    const idx = headers.indexOf(col);
    if (idx !== -1) currentVals[idx] = serialize(val);
  }

  // 3 — escribir de vuelta en un solo PATCH
  await patchRange(token, driveId, itemId, sheet, addr, currentVals);
}

// Agrega fila nueva al final
async function appendRow(token, driveId, itemId, sheet, headers, rowData) {
  const { rows } = await readSheet(token, driveId, itemId, sheet);
  const newRowNum = rows.length + 1; // siguiente fila libre (1-indexed, row 1 = header)
  const values = headers.map(h => serialize(rowData[h]));
  const addr = `A${newRowNum}:${colLetter(headers.length)}${newRowNum}`;
  await patchRange(token, driveId, itemId, sheet, addr, values);
}

// Sube imagen a OneDrive bajo IMG_FOLDER/{subfolder}/{filename}
async function uploadImage(token, driveId, subfolder, filename, buffer, mime) {
  const path = `${IMG_FOLDER}/${subfolder}/${filename}`;
  const r = await fetch(
    `${GRAPH}/drives/${driveId}/root:/${encodeURIComponent(path).replace(/%2F/g, '/')}:/content`,
    {
      method:  'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': mime },
      body:    buffer,
    }
  );
  if (!r.ok) throw new Error(`Image upload failed ${r.status}: ${await r.text()}`);
}

// ── Parseo multipart ─────────────────────────────────────────────────────────
function parseMultipart(event) {
  return new Promise((resolve, reject) => {
    const fields = {};
    const files  = [];

    const ct = event.headers['content-type'] || event.headers['Content-Type'] || '';
    let bb;
    try {
      bb = Busboy({ headers: { 'content-type': ct }, limits: { fileSize: MAX_BYTES + 1 } });
    } catch (e) { return reject(new Error('Content-Type inválido para multipart')); }

    bb.on('field', (name, val) => { fields[name] = val; });

    bb.on('file', (name, file, info) => {
      const chunks = [];
      let truncated = false;
      file.on('data',     d => chunks.push(d));
      file.on('limit',    () => { truncated = true; });
      file.on('end', () => {
        if (chunks.length) {
          files.push({ field: name, filename: info.filename, mime: info.mimeType,
                       buffer: Buffer.concat(chunks), truncated });
        }
      });
    });

    bb.on('finish', () => resolve({ fields, files }));
    bb.on('error',  reject);

    const body = event.isBase64Encoded
      ? Buffer.from(event.body, 'base64')
      : Buffer.from(event.body || '');
    bb.write(body);
    bb.end();
  });
}

// ── Handler principal ────────────────────────────────────────────────────────
exports.handler = async (event) => {
  const corsHeaders = {
    'Access-Control-Allow-Origin':  '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type':                 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 204, headers: corsHeaders, body: '' };
  }
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers: corsHeaders, body: JSON.stringify({ error: 'Método no permitido' }) };
  }

  // Verificar variables de entorno
  const { GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, BD_SHARE_URL } = process.env;
  if (!GRAPH_CLIENT_SECRET || !BD_SHARE_URL) {
    return { statusCode: 500, headers: corsHeaders,
      body: JSON.stringify({ error: 'Servidor no configurado. Contacte al administrador.' }) };
  }

  // Parsear formulario
  let fields, files;
  try {
    ({ fields, files } = await parseMultipart(event));
  } catch (e) {
    return { statusCode: 400, headers: corsHeaders,
      body: JSON.stringify({ error: `Error al leer formulario: ${e.message}` }) };
  }

  // ── Validar imágenes (magic bytes + tamaño) ──────────────────────────────
  for (const f of files) {
    if (f.truncated) {
      return { statusCode: 400, headers: corsHeaders,
        body: JSON.stringify({ error: `La imagen "${f.filename}" supera el límite de 8 MB.` }) };
    }
    const detected = detectImageMime(f.buffer);
    if (!detected) {
      return { statusCode: 400, headers: corsHeaders,
        body: JSON.stringify({ error: `"${f.filename}" no es una imagen válida. Solo se permiten JPG, PNG, HEIC y WebP.` }) };
    }
    f.mime = detected; // siempre usar el tipo detectado, no el declarado
  }

  const tipo = (fields.tipo || '').trim();
  if (!['chofer', 'unidad', 'ambos'].includes(tipo)) {
    return { statusCode: 400, headers: corsHeaders,
      body: JSON.stringify({ error: 'Tipo de formulario inválido.' }) };
  }

  const modo = (fields.modo || '').trim();
  if (!['alta', 'actualizar'].includes(modo)) {
    return { statusCode: 400, headers: corsHeaders,
      body: JSON.stringify({ error: 'Modo de formulario inválido.' }) };
  }

  // ── Validación extra para modo Alta ─────────────────────────────────────
  if (modo === 'alta') {
    if (tipo === 'chofer' || tipo === 'ambos') {
      const reqFields   = ['nombre','empresa','petrolera','patente_ch_asig',
                           'venc_carnet','venc_psico','venc_cargas'];
      const reqImgNames = ['img_carnet','img_psico','img_cargas'];
      const fieldLabels = { nombre:'Apellido y Nombre', empresa:'Empresa Transporte',
        petrolera:'Petrolera', patente_ch_asig:'Patente Chasis asignada',
        venc_carnet:'Fecha venc. Carnet', venc_psico:'Fecha venc. Psicofísico',
        venc_cargas:'Fecha venc. Cargas Peligrosas' };

      for (const f of reqFields) {
        if (!(fields[f] || '').trim()) {
          return { statusCode: 400, headers: corsHeaders,
            body: JSON.stringify({ error: `Alta — campo obligatorio: ${fieldLabels[f]}` }) };
        }
      }
      for (const f of reqImgNames) {
        if (!files.find(img => img.field === f)) {
          return { statusCode: 400, headers: corsHeaders,
            body: JSON.stringify({ error: `Alta — foto obligatoria: ${f.replace('img_','')}` }) };
        }
      }
    }
    if (tipo === 'unidad' || tipo === 'ambos') {
      const reqFields   = ['venc_itv_chas','venc_itv_cis'];
      const reqImgNames = ['img_itv_chas','img_itv_cis'];
      const fieldLabels = { venc_itv_chas:'Fecha venc. ITV Chasis',
                            venc_itv_cis:'Fecha venc. ITV Cisterna' };

      for (const f of reqFields) {
        if (!(fields[f] || '').trim()) {
          return { statusCode: 400, headers: corsHeaders,
            body: JSON.stringify({ error: `Alta — campo obligatorio: ${fieldLabels[f]}` }) };
        }
      }
      for (const f of reqImgNames) {
        if (!files.find(img => img.field === f)) {
          return { statusCode: 400, headers: corsHeaders,
            body: JSON.stringify({ error: `Alta — foto obligatoria: ${f.replace('img_','')}` }) };
        }
      }
    }
  }

  const ts = new Date().toISOString().replace(/[T:]/g, '-').slice(0, 19);
  const results = [];

  try {
    const token              = await getToken();
    const { driveId, itemId } = await resolveDrive(token, BD_SHARE_URL);

    // ── CHOFER ─────────────────────────────────────────────────────────────
    if (tipo === 'chofer' || tipo === 'ambos') {
      const dni = cleanDni(fields.dni || '');
      if (!dni) {
        return { statusCode: 400, headers: corsHeaders,
          body: JSON.stringify({ error: 'DNI es obligatorio para datos de chofer.' }) };
      }

      const { headers: shCols, rows } = await readSheet(token, driveId, itemId, 'Choferes');
      const dniIdx = shCols.indexOf('Dni');

      let targetRow = null;
      for (let i = 1; i < rows.length; i++) {
        if (cleanDni(rows[i][dniIdx]) === dni) { targetRow = i + 1; break; }
      }

      // Solo incluir campos que el usuario completó
      const upd = {};
      if (fields.nombre)        upd['Apellido y Nombre']          = fields.nombre.trim();
      if (fields.empresa)       upd['Empresa Transporte']         = fields.empresa.trim();
      if (fields.petrolera)     upd['Petrolera']                  = fields.petrolera.trim();
      if (fields.patente_ch_asig)  upd['Patente Chasis']          = cleanPat(fields.patente_ch_asig);
      if (fields.patente_cis_asig) upd['Patente Cisterna']        = cleanPat(fields.patente_cis_asig);
      if (fields.venc_carnet)   upd['Fecha Venc. Car.Con']        = fields.venc_carnet;
      if (fields.venc_psico)    upd['Fecha Venc. Psicofisico']    = fields.venc_psico;
      if (fields.venc_cargas)   upd['Fecha Venc Cargas Peligrosas'] = fields.venc_cargas;
      if (fields.obs_chofer != null) upd['Observacion']           = fields.obs_chofer.trim();

      if (targetRow) {
        await updateRow(token, driveId, itemId, 'Choferes', targetRow, upd, shCols);
        results.push(`Chofer DNI ${dni} actualizado`);
      } else {
        upd['Dni'] = dni;
        await appendRow(token, driveId, itemId, 'Choferes', shCols, upd);
        results.push(`Chofer DNI ${dni} registrado como nuevo`);
      }

      // Subir fotos del chofer
      const choferImgs = { img_carnet: 'carnet', img_psico: 'psicofisico', img_cargas: 'cargas_peligrosas' };
      for (const [field, label] of Object.entries(choferImgs)) {
        const img = files.find(f => f.field === field);
        if (img) {
          const ext = img.mime === 'image/png' ? 'png' : 'jpg';
          await uploadImage(token, driveId, `chofer_${dni}`, `${label}_${ts}.${ext}`, img.buffer, img.mime);
        }
      }
    }

    // ── UNIDAD ─────────────────────────────────────────────────────────────
    if (tipo === 'unidad' || tipo === 'ambos') {
      const patCh = cleanPat(fields.patente_ch || '');
      if (!patCh) {
        return { statusCode: 400, headers: corsHeaders,
          body: JSON.stringify({ error: 'Patente Chasis es obligatoria para datos de unidad.' }) };
      }

      const { headers: shCols, rows } = await readSheet(token, driveId, itemId, 'Patentes');
      const pchIdx = shCols.indexOf('Patente Chasis');

      let targetRow = null;
      for (let i = 1; i < rows.length; i++) {
        if (cleanPat(rows[i][pchIdx]) === patCh) { targetRow = i + 1; break; }
      }

      const upd = {};
      if (fields.patente_cis)  upd['Patente Cisterna']       = cleanPat(fields.patente_cis);
      if (fields.cod_inv)      upd['Cod INV']                = fields.cod_inv.trim();
      if (fields.venc_itv_chas) upd['Fecha Venc. ITV. Chas'] = fields.venc_itv_chas;
      if (fields.venc_itv_cis)  upd['Fecha Venc. ITV. Cis']  = fields.venc_itv_cis;
      if (fields.obs_unidad != null) upd['Observacion']      = fields.obs_unidad.trim();

      if (targetRow) {
        await updateRow(token, driveId, itemId, 'Patentes', targetRow, upd, shCols);
        results.push(`Unidad ${patCh} actualizada`);
      } else {
        upd['Patente Chasis'] = patCh;
        await appendRow(token, driveId, itemId, 'Patentes', shCols, upd);
        results.push(`Unidad ${patCh} registrada como nueva`);
      }

      // Subir fotos de la unidad
      const unidadImgs = { img_itv_chas: 'itv_chasis', img_itv_cis: 'itv_cisterna' };
      for (const [field, label] of Object.entries(unidadImgs)) {
        const img = files.find(f => f.field === field);
        if (img) {
          const ext = img.mime === 'image/png' ? 'png' : 'jpg';
          await uploadImage(token, driveId, `unidad_${patCh}`, `${label}_${ts}.${ext}`, img.buffer, img.mime);
        }
      }
    }

    return {
      statusCode: 200,
      headers: corsHeaders,
      body: JSON.stringify({ ok: true, message: results.join('. ') + '.' }),
    };

  } catch (err) {
    console.error('[submit]', err.message);
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: `Error del servidor: ${err.message}` }),
    };
  }
};
