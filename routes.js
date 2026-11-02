// src/routes.js
// Todas las acciones del sistema — equivalente al WebApp.gs + Portal médico

const express = require('express');
const router = express.Router();
const { leerRango, escribirCelda, escribirFila, agregarFila, col, formatFecha, getSheets, SPREADSHEET_ID } = require('./sheets');

// Escribe varias filas en UNA sola llamada a Google Sheets (para importaciones)
async function escribirFilas(rango, filas) {
  const sheets = await getSheets();
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: rango,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: filas }
  });
}
const { crearConcepto, construirFactura, timbrar, cancelar, getSaldo, buscarPorUUID } = require('./facturama');
const { guardarFacturaEnDrive, guardarComplementoEnDrive, getLinkCarpetaFactura } = require('./drive');
const { descargarArchivo, listarCfdisEmitidos } = require('./facturama');
const auth = require('./portalAuth');

// ─── Login (sistema interno) ─────────────────────────────────
// La contraseña en USUARIOS puede estar en texto o cifrada (scrypt, igual que la de los médicos).
// 8 intentos fallidos en 15 min bloquean ese usuario 15 min.
async function wa_login(usuario, contrasena) {
  usuario = String(usuario || '').trim(); contrasena = String(contrasena || '');
  const clave = 'admin:' + usuario.toLowerCase();
  if (!usuario || !contrasena) return { ok: false };
  if (auth.estaBloqueado(clave)) return { ok: false, error: 'Demasiados intentos. Espera 15 minutos.' };
  const datos = await leerRango('USUARIOS!A:C');
  for (let i = 1; i < datos.length; i++) {
    const u = (datos[i][0] || '').toString().trim();
    const p = (datos[i][1] || '').toString().trim();
    const n = (datos[i][2] || '').toString().trim();
    if (!u || u !== usuario || !p) continue;
    const ok = auth.esHash(p) ? auth.verificarPassword(contrasena, p) : p === contrasena;
    if (ok) { auth.limpiarFallos(clave); return { ok: true, nombre: n, usuario: u }; }
  }
  auth.registrarFallo(clave);
  return { ok: false };
}

// Sesión del sistema interno: cada acción manda usuario + contraseña (solo en memoria del navegador).
// Para no leer la hoja USUARIOS en cada clic, una sesión válida se recuerda 5 minutos.
const sesionesAdmin = new Map();
async function verificarAdmin(p) {
  const usuario = String((p && p.usuario) || '').trim(), contrasena = String((p && p.contrasena) || '');
  if (!usuario || !contrasena) return false;
  const k = require('crypto').createHash('sha256').update(usuario + '\u0000' + contrasena).digest('hex');
  const c = sesionesAdmin.get(k);
  if (c && Date.now() - c < 5 * 60 * 1000) return true;
  const r = await wa_login(usuario, contrasena);
  if (r.ok) sesionesAdmin.set(k, Date.now()); else sesionesAdmin.delete(k);
  return r.ok;
}

// ─── Obtener emisores ─────────────────────────────────────────
async function wa_getEmisores() {
  const datos = await leerRango('EMISORES!A:K');
  const result = [];
  for (let i = 1; i < datos.length; i++) {
    const nombre = (datos[i][2] || '').toString().trim();
    const rfc = (datos[i][1] || '').toString().trim();
    const regimen = (datos[i][4] || '').toString().trim();
    if (nombre && rfc) result.push({ nombre, rfc, regimen, tasaIsr: tasaIsrPorRegimen(regimen) });
  }
  return result;
}

// ─── Obtener aseguradoras ─────────────────────────────────────
async function wa_getAseguradoras() {
  const datos = await leerRango('Datos aseguradoras!A:F');
  const result = [];
  for (let i = 1; i < datos.length; i++) {
    const nombre = (datos[i][0] || '').toString().trim();
    if (!nombre) continue;
    result.push({
      nombre,
      rfc: (datos[i][1] || '').toString().trim(),
      nombreFiscal: (datos[i][2] || '').toString().trim(),
      cp: (datos[i][3] || '').toString().trim(),
      usoCfdi: (datos[i][5] || 'G03').toString().trim()
    });
  }
  return result;
}

// ─── Retención de ISR según el régimen del médico ─────────────
// 626 RESICO: 1.25% · 612 Actividades empresariales y profesionales: 10%
// (honorarios médicos: IVA exento en ambos casos)
const TASA_ISR = { '626': 0.0125, '612': 0.10 };
function tasaIsrPorRegimen(regimen) {
  const m = String(regimen || '').match(/\d{3}/);
  return (m && TASA_ISR[m[0]]) || 0.0125;
}
// Tasa con la que realmente se emitió una factura (para el complemento), a partir de bruto y neto
function tasaIsrDeFactura(bruto, neto, porDefecto) {
  if (!(bruto > 0) || !(neto > 0) || neto > bruto) return porDefecto;
  const r = (bruto - neto) / bruto;
  return Object.values(TASA_ISR).reduce((mejor, t) => Math.abs(t - r) < Math.abs(mejor - r) ? t : mejor, porDefecto);
}

// ─── Obtener datos emisor por nombre ──────────────────────────
async function obtenerDatosEmisor(nombreBuscado) {
  const datos = await leerRango('EMISORES!A:K');
  const nombre = nombreBuscado.toString().trim().toUpperCase();
  for (let i = 1; i < datos.length; i++) {
    if ((datos[i][2] || '').toString().trim().toUpperCase() === nombre) {
      return {
        rfc: (datos[i][1] || '').toString().trim(),
        nombre: (datos[i][2] || '').toString().trim(),
        cp: (datos[i][3] || '').toString().trim(),
        regimen: (datos[i][4] || '').toString().trim(),
        serie: (datos[i][5] || '').toString().trim(),
        tasaIsr: tasaIsrPorRegimen(datos[i][4])
      };
    }
  }
  throw new Error('Emisor no encontrado: ' + nombreBuscado);
}

// ─── Obtener facturas (sistema interno, todas) ────────────────
async function wa_getFacturas() {
  const data = await leerRango('Registro!A:AZ');
  if (!data.length) return [];
  const headers = data[0];

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const medico = (row[col(headers, 'EMISOR (MEDICO)')] || '').toString().trim();
    if (!medico) continue;

    const monto1 = Number(row[col(headers, 'MONTO CONCEPTO 1')]) || 0;
    const monto2 = Number(row[col(headers, 'MONTO CONCEPTO 2')]) || 0;

    result.push({
      fila: i + 1,
      medico,
      aseguradora: (row[col(headers, 'ASEGURADORA')] || '').toString().trim(),
      monto: monto1 + monto2,
      montoFinal: Number(row[col(headers, 'MONTO FINAL A PAGAR')]) || 0,
      estado: (row[col(headers, 'ESTADO')] || '').toString().trim().toUpperCase(),
      estadoComp: (row[col(headers, 'ESTADO COMPLEMENTO')] || '').toString().trim().toUpperCase(),
      fechaSubida: formatFecha(row[col(headers, 'Fecha de subida al portal')]),
      fechaPago: formatFecha(row[col(headers, 'FECHA DE TRANSFERENCIA')]),
      uuid: (row[col(headers, 'Folio Factura (UUID)')] || '').toString().trim(),
      uuidComp: (row[col(headers, 'Folio Complemento(UUID)')] || '').toString().trim(),
      idPago: (row[col(headers, 'ID PAGO')] || '').toString().trim(),
      folioPortal: (row[col(headers, 'FOLIO PORTAL')] || '').toString().trim(),
      observaciones: (row[col(headers, 'OBSERVACIONES')] || '').toString().trim(),
      ultimaModificacion: (row[col(headers, 'ULTIMA MODIFICACION')] || '').toString().trim(),
      // Para el buscador por palabra clave (paciente, procedimiento, receptor)
      concepto1: (row[col(headers, 'CONCEPTO DE FACTURA 1')] || '').toString().trim(),
      concepto2: (row[col(headers, 'CONCEPTO DE FACTURA 2')] || '').toString().trim(),
      nombreFiscal: (row[col(headers, 'NOMBRE FISCAL')] || '').toString().trim(),
      rfcReceptor: (row[col(headers, 'RFC ASEGURADORA')] || '').toString().trim()
    });
  }
  return result;
}

// ─── Conceptos múltiples ──────────────────────────────────────
// En la hoja: K/L = concepto 1; M/N = conceptos 2..n juntos (" || ") con su monto sumado,
// así MONTO CONCEPTO 1 + 2 sigue siendo el total (lo usan el complemento y los reportes).
// El detalle exacto se guarda como JSON en la columna "CONCEPTOS DETALLE" (se crea sola).
const COL_DETALLE = 'CONCEPTOS DETALLE';
const SEP_CONCEPTOS = ' || ';

function normalizarConceptos(p) {
  let lista = Array.isArray(p.conceptos) ? p.conceptos : [
    { concepto: p.concepto1, monto: p.monto1 },
    { concepto: p.concepto2, monto: p.monto2 }
  ];
  return lista
    .map(c => ({ concepto: String((c && c.concepto) || '').trim(), monto: Number(c && c.monto) || 0 }))
    .filter(c => c.concepto && c.monto > 0)
    .map(c => ({ concepto: c.concepto, monto: Number(c.monto.toFixed(2)) }));
}

function conceptosParaHoja(conceptos) {
  const primero = conceptos[0];
  const resto = conceptos.slice(1);
  return {
    concepto1: primero.concepto,
    monto1: primero.monto,
    concepto2: resto.map(c => c.concepto).join(SEP_CONCEPTOS),
    monto2: Number(resto.reduce((s, c) => s + c.monto, 0).toFixed(2))
  };
}

async function columnaDetalle() {
  const enc = await leerRango('Registro!1:1');
  const headers = (enc[0] || []).map(h => String(h || '').trim().toUpperCase());
  let i = headers.indexOf(COL_DETALLE);
  if (i < 0) {
    i = Math.max(headers.length, 30); // después de AD (ULTIMA MODIFICACION)
    await escribirCelda(`Registro!${columnLetter(i + 1)}1`, COL_DETALLE);
  }
  return columnLetter(i + 1);
}

// ─── Timbrar fila existente ───────────────────────────────────
async function wa_timbrarFila(fila) {
  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];
  const row = data[fila - 1];

  const medico = (row[col(headers, 'EMISOR (MEDICO)')] || '').toString().trim();
  const emisor = await obtenerDatosEmisor(medico);

  const rfcRec = (row[col(headers, 'RFC ASEGURADORA')] || '').toString();
  const nombreRec = (row[col(headers, 'NOMBRE FISCAL')] || '').toString();
  const cpRec = (row[col(headers, 'CODIGO POSTAL ASEGURADORA')] || '').toString();
  const usoCfdi = (row[col(headers, 'USO CFDI')] || 'G03').toString();
  const concepto1 = (row[col(headers, 'CONCEPTO DE FACTURA 1')] || '').toString();
  const monto1 = Number(row[col(headers, 'MONTO CONCEPTO 1')]) || 0;
  const concepto2 = (row[col(headers, 'CONCEPTO DE FACTURA 2')] || '').toString();
  const monto2 = Number(row[col(headers, 'MONTO CONCEPTO 2')]) || 0;

  const items = [];
  let detalle = null;
  try {
    const iDet = headers.map(h => String(h || '').trim().toUpperCase()).indexOf(COL_DETALLE);
    if (iDet >= 0 && row[iDet]) detalle = normalizarConceptos({ conceptos: JSON.parse(row[iDet]) });
  } catch (e) { detalle = null; }
  if (detalle && detalle.length) {
    detalle.forEach(c => items.push(crearConcepto(c.concepto, c.monto, emisor.tasaIsr)));
  } else {
    if (monto1 > 0) items.push(crearConcepto(concepto1, monto1, emisor.tasaIsr));
    if (monto2 > 0) items.push(crearConcepto(concepto2, monto2, emisor.tasaIsr));
  }
  if (!items.length) return { ok: false, error: 'Sin conceptos con monto > 0' };

  const factura = construirFactura(emisor, rfcRec, nombreRec, cpRec, usoCfdi, items, fila);
  const result = await timbrar(factura);

  if (result.uuid) {
    const colUUID = col(headers, 'Folio Factura (UUID)') + 1;
    const colEstado = col(headers, 'ESTADO') + 1;
    const colID = col(headers, 'ID FACTURAMA') + 1;

    const letraUUID = columnLetter(colUUID);
    const letraEstado = columnLetter(colEstado);
    const letraID = columnLetter(colID);

    await escribirCelda(`Registro!${letraUUID}${fila}`, result.uuid);
    await escribirCelda(`Registro!${letraEstado}${fila}`, 'TIMBRADO');
    await escribirCelda(`Registro!${letraID}${fila}`, result.id);
    // MONTO FINAL A PAGAR con la tasa de ISR del régimen del médico (1.25% RESICO / 10% 612)
    try {
      const neto = Number(items.reduce((t, it) => t + Number(it.Total || 0), 0).toFixed(2));
      await escribirCelda(`Registro!${columnLetter(col(headers, 'MONTO FINAL A PAGAR') + 1)}${fila}`, neto);
    } catch (e) { console.error('No se pudo escribir MONTO FINAL A PAGAR:', e.message); }

    guardarFacturaEnDrive(result.id, medico, nombreRec, emisor.serie, fila, concepto1).catch(console.error);
    return { ok: true, uuid: result.uuid };
  }
  return { ok: false, error: result.error };
}

// ─── Timbrar nueva factura ────────────────────────────────────
 async function wa_timbrarNueva(p) {
 const data = await leerRango('Registro!A:T');
// Buscar la primera fila vacía después del encabezado
let nuevaFila = 2;
for (let i = 1; i < data.length; i++) {
  const medico = (data[i][0] || '').toString().trim();
  const uuid = (data[i][19] || '').toString().trim(); // columna T = UUID
  if (medico || uuid) {
    nuevaFila = i + 2; // siguiente fila después de esta
  }
}

  const emisor = await obtenerDatosEmisor(p.medico);
  const esFisica = p.receptorTipo === 'FISICA';

  // Factura a persona física (paciente): validar datos del receptor
  if (esFisica) {
    const rfcRec = String(p.rfcAseguradora || '').trim().toUpperCase();
    if (!/^[A-ZÑ&]{4}[0-9]{6}[A-Z0-9]{3}$/.test(rfcRec)) {
      return { ok: false, error: 'El RFC del receptor debe ser de persona física (13 caracteres).' };
    }
    if (!String(p.nombreFiscal || '').trim()) return { ok: false, error: 'Falta el nombre fiscal del receptor.' };
    if (!/^[0-9]{5}$/.test(String(p.cpAseguradora || '').trim())) return { ok: false, error: 'El código postal fiscal debe tener 5 dígitos.' };
    if (!String(p.regimenReceptor || '').trim()) return { ok: false, error: 'Falta el régimen fiscal del receptor.' };
  }

  const conceptos = normalizarConceptos(p);
  if (!conceptos.length) return { ok: false, error: 'Sin conceptos con monto > 0' };
  if (conceptos.length > 30) return { ok: false, error: 'Máximo 30 conceptos por factura' };
  const items = conceptos.map(c => crearConcepto(c.concepto, c.monto, emisor.tasaIsr));
  const hoja = conceptosParaHoja(conceptos);

  // Persona física: sí objeto de impuesto (02), IVA exento y SIN retención de ISR
  if (esFisica) {
    items.forEach(function(it) {
      it.TaxObject = '02';
      it.Taxes = (it.Taxes || []).filter(function(t) { return t.Name !== 'ISR'; });
      it.Total = Number(it.Subtotal != null ? it.Subtotal : it.UnitPrice);
    });
  }

  const factura = construirFactura(emisor, p.rfcAseguradora, p.nombreFiscal, p.cpAseguradora, p.usoCfdi, items, nuevaFila);

  // Persona física: pago en una sola exhibición y régimen propio del receptor
  if (esFisica) {
    factura.PaymentMethod = 'PUE';
    factura.PaymentForm = String(p.formaPago || '03');
    factura.Receiver.FiscalRegime = String(p.regimenReceptor).trim();
  }

  const result = await timbrar(factura);

  if (result.uuid) {
    // Neto = suma de los totales de cada concepto, con el mismo redondeo que la factura timbrada
    const montoFinal = Number(items.reduce((s, it) => s + Number(it.Total || 0), 0).toFixed(2));

    // Escribir toda la fila de una sola vez — 28 columnas A hasta AB
    const fila = new Array(28).fill('');
    fila[0]  = p.medico;
fila[1]  = emisor.rfc;        // B - RFC EMISOR
fila[2]  = emisor.nombre;     // C - NOMBRE EMISOR
fila[3]  = emisor.cp;         // D - CP EMISOR
fila[4]  = emisor.regimen;    // E - REGIMEN FISCAL EMISOR
fila[5]  = p.aseguradora;
fila[6]  = p.rfcAseguradora;
fila[7]  = p.nombreFiscal;
fila[8]  = p.cpAseguradora;
fila[9]  = p.usoCfdi;
fila[10] = hoja.concepto1;
fila[11] = hoja.monto1;
fila[12] = hoja.concepto2;
fila[13] = hoja.monto2;
fila[15] = montoFinal;
fila[17] = esFisica ? 'PAGADA (PUE)' : 'TIMBRADO';
if (esFisica) fila[20] = "'" + new Date().toLocaleDateString('es-MX', { timeZone: 'America/Mexico_City', day: '2-digit', month: '2-digit', year: 'numeric' }); // U - pagada al emitir (PUE)
fila[19] = result.uuid;
fila[27] = result.id;

    await escribirFila(`Registro!A${nuevaFila}:AB${nuevaFila}`, fila);
    if (conceptos.length > 2) {
      try {
        const letraDet = await columnaDetalle();
        await escribirCelda(`Registro!${letraDet}${nuevaFila}`, JSON.stringify(conceptos));
      } catch (e) { console.error('No se pudo guardar el detalle de conceptos:', e.message); }
    }

    guardarFacturaEnDrive(result.id, p.medico, p.nombreFiscal, emisor.serie, nuevaFila, hoja.concepto1).catch(console.error);
    return { ok: true, uuid: result.uuid };
  }
  return { ok: false, error: result.error };
}
// ─── Timbrar complemento de pago ─────────────────────────────
async function wa_timbrarComplemento(idPago, fechaPagoStr, filaEspecifica) {
  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];

  const fechaPago = new Date(fechaPagoStr).toISOString().replace('.000Z', '');

const filasGrupo = [];
if (filaEspecifica) {
  filasGrupo.push(filaEspecifica - 1);
} else {
  for (let i = 1; i < data.length; i++) {
    const idFila = (data[i][col(headers, 'ID PAGO')] || '').toString().trim();
    if (idFila === idPago) filasGrupo.push(i);
  }
}
if (!filasGrupo.length) return { ok: false, error: 'No se encontró la factura' };
  // ── VALIDACIÓN OPCIÓN A: si el ID de pago está en varias filas, devolver opciones ──
  if (filasGrupo.length > 1) {
    const opciones = filasGrupo.map(function(idx) {
      return {
        fila: idx + 1,
        aseguradora: (data[idx][col(headers, 'ASEGURADORA')] || '').toString().trim(),
        nombreFiscal: (data[idx][col(headers, 'NOMBRE FISCAL')] || '').toString().trim(),
        monto: Number(data[idx][col(headers, 'MONTO FINAL A PAGAR')]) || 0,
        uuid: (data[idx][col(headers, 'Folio Factura (UUID)')] || '').toString().trim(),
        estadoComp: (data[idx][col(headers, 'ESTADO COMPLEMENTO')] || '').toString().trim()
      };
    });
    return { ok: false, requiereSeleccion: true, idPago: idPago, opciones: opciones };
  }

  const doctoRelacionados = [];
  for (const idx of filasGrupo) {
    const uuid = (data[idx][col(headers, 'Folio Factura (UUID)')] || '').toString().trim();
    const monto = Number(data[idx][col(headers, 'MONTO FINAL A PAGAR')]) || 0;
    const estado = (data[idx][col(headers, 'ESTADO COMPLEMENTO')] || '').toString().trim().toUpperCase();
    const montoConceptos = (Number(data[idx][col(headers, 'MONTO CONCEPTO 1')]) || 0) +
                           (Number(data[idx][col(headers, 'MONTO CONCEPTO 2')]) || 0);
    if (!uuid) return { ok: false, error: `Fila ${idx + 1} sin UUID` };
    if (estado === 'COMPLEMENTADO') return { ok: false, error: `Fila ${idx + 1} ya complementada` };
    doctoRelacionados.push({ fila: idx, uuid, monto, montoConceptos });
  }

  const primeraFila = data[filasGrupo[0]];
  const medicoNombre = (primeraFila[col(headers, 'EMISOR (MEDICO)')] || '').toString().trim();
  const emisor = await obtenerDatosEmisor(medicoNombre);
  const rfcReceptor = (primeraFila[col(headers, 'RFC ASEGURADORA')] || '').toString().trim();
  const nombreReceptor = (primeraFila[col(headers, 'NOMBRE FISCAL')] || '').toString().trim();
  const cpReceptor = (primeraFila[col(headers, 'CODIGO POSTAL ASEGURADORA')] || '').toString().trim();

  const tasaIsr = emisor.tasaIsr || 0.0125;
 doctoRelacionados.forEach(d => { d.tasa = tasaIsrDeFactura(d.montoConceptos, d.monto, tasaIsr); });
 const documentosRel = doctoRelacionados.map(d => ({
  TaxObject: '02',
  Uuid: d.uuid,
  Serie: String(emisor.serie).trim(),
  Currency: 'MXN',
  PaymentMethod: 'PPD',
  PartialityNumber: 1,
  PreviousBalanceAmount: Number(d.monto.toFixed(2)),  // saldo anterior = lo que se debía (monto neto)
  AmountPaid: Number(d.monto.toFixed(2)),             // se paga completo
  RemainingBalance: 0,                                 // saldo insoluto = 0
  Taxes: [{
    Name: 'ISR',
    Base: Number(d.montoConceptos.toFixed(2)),
    Rate: d.tasa,
    Total: Number((d.montoConceptos * d.tasa).toFixed(2)),
    IsRetention: true
  }]
}));

  const totalPago = Number(doctoRelacionados.reduce((s, d) => s + d.monto, 0).toFixed(2));

  const rep = {
    CfdiType: 'P',
    Serie: String(emisor.serie).trim(),
    Folio: String(idPago || filaEspecifica || '').trim(),
    ExpeditionPlace: String(emisor.cp).padStart(5, '0'),
    Issuer: { Rfc: emisor.rfc, Name: emisor.nombre, FiscalRegime: emisor.regimen },
    Receiver: { Rfc: rfcReceptor, Name: nombreReceptor, CfdiUse: 'CP01', FiscalRegime: '601', TaxZipCode: cpReceptor },
    Complemento: {
      Payments: [{ Date: fechaPago, PaymentForm: '03', Currency: 'MXN', Amount: totalPago, RelatedDocuments: documentosRel }]
    }
  };

  const fetch = require('node-fetch');
  const { getAuth } = require('./facturama');
  const BASE = process.env.FACTURAMA_BASE || 'https://api.facturama.mx';

  const res = await fetch(`${BASE}/api-lite/3/cfdis`, {
    method: 'POST',
    headers: { 'Authorization': getAuth(), 'Content-Type': 'application/json' },
    body: JSON.stringify(rep)
  });
  const raw = await res.text();
  if (res.status !== 200 && res.status !== 201) return { ok: false, error: raw };

  const result = JSON.parse(raw);
  const uuidRep = result?.Complement?.TaxStamp?.Uuid || null;
  const idRep = result?.Id || null;

  if (uuidRep) {
for (const idx of filasGrupo) {
  const filaNum = idx + 1;
  const colUUIDComp = columnLetter(col(headers, 'Folio Complemento(UUID)') + 1);
  const colIDComp = columnLetter(col(headers, 'ID FACTURAMA COMPLEMENTO') + 1);
  const colEstComp = columnLetter(col(headers, 'ESTADO COMPLEMENTO') + 1);
  const colEstado = columnLetter(col(headers, 'ESTADO') + 1);  // ← nueva
  await escribirCelda(`Registro!${colUUIDComp}${filaNum}`, uuidRep);
  await escribirCelda(`Registro!${colIDComp}${filaNum}`, idRep);
  await escribirCelda(`Registro!${colEstComp}${filaNum}`, 'COMPLEMENTADO');
  await escribirCelda(`Registro!${colEstado}${filaNum}`, 'COMPLEMENTO POR SUBIR');  // ← nueva
}
    const uuidOriginal = doctoRelacionados[0].uuid;
    guardarComplementoEnDrive(idRep, medicoNombre, nombreReceptor, uuidOriginal).catch(console.error);
    return { ok: true, uuid: uuidRep };
  }
  return { ok: false, error: raw };
}

// ─── Guardar datos (edición) ──────────────────────────────────
async function wa_guardarDatos(params) {
  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];
  const fila = params.fila;
  const setVal = async (colName, val) => {
    try {
      const idx = col(headers, colName);
      const letra = columnLetter(idx + 1);
      await escribirCelda(`Registro!${letra}${fila}`, val);
    } catch (e) {}
  };
  if (params.folioPortal !== undefined) await setVal(' FOLIO PORTAL', params.folioPortal);
  if (params.fechaSubida !== undefined && params.fechaSubida !== '') await setVal('Fecha de subida al portal', params.fechaSubida);
  if (params.idPago !== undefined && params.idPago !== '') await setVal('ID PAGO', params.idPago);
  if (params.fechaPago !== undefined && params.fechaPago !== '') await setVal('FECHA DE TRANSFERENCIA', params.fechaPago);
  if (params.estado !== undefined && params.estado !== '') await setVal('ESTADO', params.estado);
  if (params.observaciones !== undefined) await setVal('OBSERVACIONES', params.observaciones);
  if (params.ultimaModificacion !== undefined) await setVal('ULTIMA MODIFICACION', params.ultimaModificacion);
  return { ok: true };
}

// ─── Cancelar factura ─────────────────────────────────────────
async function wa_cancelarFactura(fila, motivo = '02') {
  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];
  const row = data[fila - 1];

  let id = (row[col(headers, 'ID FACTURAMA')] || '').toString().trim();
  const uuid = (row[col(headers, 'Folio Factura (UUID)')] || '').toString().trim();

  if (!id && uuid) {
    id = await buscarPorUUID(uuid);
    if (id) {
      const letraID = columnLetter(col(headers, 'ID FACTURAMA') + 1);
      await escribirCelda(`Registro!${letraID}${fila}`, id);
    }
  }
  if (!id) return { ok: false, error: 'Sin ID FACTURAMA' };

  const res = await cancelar(id, motivo);
  if (res.ok) {
    const letraEstado = columnLetter(col(headers, 'ESTADO') + 1);
    const letraObs = columnLetter(col(headers, 'OBSERVACIONES') + 1);
    await escribirCelda(`Registro!${letraEstado}${fila}`, 'CANCELADO');
    await escribirCelda(`Registro!${letraObs}${fila}`, 'Cancelado via app');
    return { ok: true };
  }
  return { ok: false, error: res.raw };
}

async function wa_cancelarComplemento(p) {
  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];
  const fila = p.fila;
  const idComp = (data[fila - 1][col(headers, 'ID FACTURAMA COMPLEMENTO')] || '').toString().trim();
  if (!idComp) return { ok: false, error: 'Esta fila no tiene complemento emitido' };

  const result = await cancelar(idComp, p.motive || '02');
  if (!result.ok) return result;

  const colEstComp = columnLetter(col(headers, 'ESTADO COMPLEMENTO') + 1);
  const colUUIDComp = columnLetter(col(headers, 'Folio Complemento(UUID)') + 1);
  const colEstado = columnLetter(col(headers, 'ESTADO') + 1);
  await escribirCelda(`Registro!${colEstComp}${fila}`, 'COMPLEMENTO CANCELADO');
  await escribirCelda(`Registro!${colUUIDComp}${fila}`, '');
  await escribirCelda(`Registro!${colEstado}${fila}`, 'EN ESPERA DE TRANSFERENCIA');
  return { ok: true };
}

// ─── Obtener timbres ──────────────────────────────────────────
async function wa_getTimbres() {
  const saldo = await getSaldo();
  return { saldo };
}

// ════════════════════════════════════════════════════════════
// PORTAL MÉDICO
// ════════════════════════════════════════════════════════════

// ─── Hoja EMISORES con columnas del portal ────────────────────
// Columna K = contraseña (cifrada). Las demás se buscan por nombre de encabezado
// y, si no existen, se crean solas a partir de la columna L.
const COLS_PORTAL = {
  temp:        'CONTRASEÑA TEMPORAL',
  vence:       'TEMPORAL VENCE',
  activado:    'PORTAL ACTIVADO',
  fechaCambio: 'FECHA CAMBIO CONTRASEÑA',
  ultimoEnvio: 'ULTIMO ENVIO DE ACCESO',
  token:       'LIGA CAPTURA (TOKEN)',
  gestoria:        'GESTORIA',
  gestoriaFecha:   'GESTORIA ACEPTADA',
  gestoriaVersion: 'GESTORIA VERSION',
  plan:            'PLAN'
};
const IDX_PASS = 10; // columna K

function normHeader(h) {
  return String(h || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
}

async function leerEmisoresPortal() {
  const data = await leerRango('EMISORES!A:Z');
  const headers = (data[0] || []).slice();
  const norm = headers.map(normHeader);
  const idx = { pass: IDX_PASS };
  let siguiente = Math.max(headers.length, IDX_PASS + 1);

  async function asegurar(clave, nombre, buscar) {
    let i = buscar ? norm.findIndex(buscar) : norm.indexOf(normHeader(nombre));
    if (i < 0) {
      i = siguiente++;
      await escribirCelda(`EMISORES!${columnLetter(i + 1)}1`, nombre);
      headers[i] = nombre; norm[i] = normHeader(nombre);
    }
    idx[clave] = i;
  }

  await asegurar('correo', 'CORREO', h => /CORREO|E-?MAIL/.test(h));
  for (const [clave, nombre] of Object.entries(COLS_PORTAL)) await asegurar(clave, nombre);
  return { data, idx };
}

function celda(row, i) { return ((row || [])[i] || '').toString().trim(); }

async function buscarMedicoPortal(rfc) {
  const { data, idx } = await leerEmisoresPortal();
  const rfcBuscado = (rfc || '').toString().trim().toUpperCase();
  if (!rfcBuscado) return { idx, medico: null };
  for (let i = 1; i < data.length; i++) {
    if (celda(data[i], 1).toUpperCase() === rfcBuscado) {
      return { idx, medico: { fila: i + 1, row: data[i], rfc: rfcBuscado, nombre: celda(data[i], 2), correo: celda(data[i], idx.correo) } };
    }
  }
  return { idx, medico: null };
}

async function setCeldaEmisor(fila, i, valor) {
  await escribirCelda(`EMISORES!${columnLetter(i + 1)}${fila}`, valor);
}

// ─── Cuentas gratis (hoja "CUENTAS") ─────────────────────────
// Médicos que se registran solos para llevar el control de sus pacientes.
// No son emisores: no aparecen en "Nueva factura". Al contratar, Olga los da
// de alta en EMISORES con el MISMO RFC y conservan pacientes, contraseña y liga.
const HOJA_CUENTAS = 'CUENTAS';
// El médico gratis entra con su CORREO. La columna RFC queda vacía: Olga la llena cuando contrata.
const ENC_CUENTAS = ['RFC', 'NOMBRE', 'CORREO', 'CELULAR', 'FECHA NACIMIENTO', 'CONTRASEÑA', 'FECHA REGISTRO', 'ACEPTO AVISO',
  'LIGA CAPTURA (TOKEN)', 'CONTRASEÑA TEMPORAL', 'TEMPORAL VENCE', 'PIDIO SEGUIMIENTO', 'NOTAS OGM',
  'GESTORIA', 'GESTORIA ACEPTADA', 'GESTORIA VERSION'];
const IDX_CUENTA = { rfc: 0, nombre: 1, correo: 2, celular: 3, nacimiento: 4, pass: 5, fecha: 6, aviso: 7, token: 8, temp: 9, vence: 10, seguimiento: 11, gestoria: 13, gestoriaFecha: 14, gestoriaVersion: 15,
  activado: -1, fechaCambio: -1, ultimoEnvio: -1 };

// Usuario del portal: RFC (clientes) o correo (cuentas gratis)
function normUsuario(v) {
  const t = String(v || '').trim();
  return t.includes('@') ? t.toLowerCase() : t.toUpperCase();
}
function mismaClave(a, b) { return String(a || '').trim().toUpperCase() === String(b || '').trim().toUpperCase(); }
const ULT_COL_CUENTAS = columnLetter(ENC_CUENTAS.length);

// Devuelve null si la pestaña CUENTAS no existe todavía
async function leerCuentas() {
  let data;
  try { data = await leerRango(`${HOJA_CUENTAS}!A:${ULT_COL_CUENTAS}`); } catch (e) { return null; }
  data = data || [];
  const enc = data[0] || [];
  if (ENC_CUENTAS.some((h, i) => normHeader(enc[i]) !== normHeader(h))) {
    await escribirFila(`${HOJA_CUENTAS}!A1:${ULT_COL_CUENTAS}1`, ENC_CUENTAS);
    if (!data.length) data.push(ENC_CUENTAS.slice()); else data[0] = ENC_CUENTAS.slice();
  }
  return data;
}

function cuentaGratisDeFila(data, i) {
  const row = data[i], fila = i + 1;
  const rfcReal = celda(row, 0).toUpperCase(), correo = celda(row, 2).toLowerCase();
  return {
    plan: 'gratis', fila, row, idx: IDX_CUENTA,
    rfc: rfcReal || correo,       // clave con la que se guardan sus pacientes y con la que entra
    rfcReal, nombre: celda(row, 1), correo, celular: celda(row, 3),
    set: (k, v) => k < 0 ? null : escribirCelda(`${HOJA_CUENTAS}!${columnLetter(k + 1)}${fila}`, v)
  };
}

async function buscarCuentaGratis(rfc) {
  const data = await leerCuentas();
  if (!data) return null;
  for (let i = 1; i < data.length; i++) {
    if (mismaClave(celda(data[i], 0), rfc) || mismaClave(celda(data[i], 2), rfc)) return cuentaGratisDeFila(data, i);
  }
  return null;
}

// Cliente (EMISORES) o cuenta gratis (CUENTAS), con la misma forma
async function buscarCuenta(rfc) {
  const rfcB = normUsuario(rfc);
  if (!rfcB) return null;
  const comoCliente = async r => {
    const { idx, medico } = await buscarMedicoPortal(r);
    return medico ? Object.assign(medico, { plan: 'cliente', idx, set: (k, v) => k < 0 ? null : setCeldaEmisor(medico.fila, k, v) }) : null;
  };
  if (!rfcB.includes('@')) { const c = await comoCliente(rfcB); if (c) return c; }
  else {
    // cliente que escribe su correo en lugar del RFC (columna CORREO de EMISORES)
    const { data, idx } = await leerEmisoresPortal();
    for (let i = 1; i < data.length; i++) {
      if (celda(data[i], 1) && mismaClave(celda(data[i], idx.correo), rfcB)) return await comoCliente(celda(data[i], 1).toUpperCase());
    }
  }
  const gratis = await buscarCuentaGratis(rfcB);
  // cuenta gratis que ya contrató (Olga le puso su RFC y lo dio de alta en EMISORES)
  if (gratis && gratis.rfcReal) { const c = await comoCliente(gratis.rfcReal); if (c) return c; }
  return gratis;
}

// ─── Login del médico (por RFC) ───────────────────────────────
// Acepta su contraseña normal o una temporal vigente.
// requiereCambioPassword = true si entró con temporal o si su contraseña aún no está cifrada.
async function wa_loginMedico(rfc, contrasena) {
  const rfcB = normUsuario(rfc);
  if (auth.estaBloqueado(rfcB)) {
    return { ok: false, error: 'Demasiados intentos. Espera 15 minutos o usa "¿Olvidaste tu contraseña?".' };
  }
  const medico = await buscarCuenta(rfcB);
  if (!medico) { auth.registrarFallo(rfcB); return { ok: false, error: 'Usuario o contraseña incorrectos' }; }
  const idx = medico.idx;

  let guardada = celda(medico.row, idx.pass);
  const temp = celda(medico.row, idx.temp);
  const vence = Date.parse(celda(medico.row, idx.vence));

  let okNormal = auth.verificarPassword(contrasena, guardada);
  const okTemp = !okNormal && temp && !isNaN(vence) && vence > Date.now() && auth.verificarPassword(contrasena, temp);

  // Cuenta gratis que ya contrató: aún no tiene contraseña en EMISORES → usa la de su cuenta gratis
  if (!okNormal && !okTemp && medico.plan === 'cliente' && !guardada) {
    const gratis = await buscarCuentaGratis(medico.rfc);
    const passGratis = gratis ? celda(gratis.row, IDX_CUENTA.pass) : '';
    if (passGratis && auth.verificarPassword(contrasena, passGratis)) {
      okNormal = true; guardada = passGratis;
      await medico.set(idx.pass, passGratis);
      await medico.set(idx.activado, 'SÍ');
    }
  }

  if (!okNormal && !okTemp) {
    auth.registrarFallo(rfcB);
    return { ok: false, error: 'Usuario o contraseña incorrectos' };
  }
  auth.limpiarFallos(rfcB);
  return {
    ok: true,
    nombre: medico.nombre,
    rfc: medico.rfc,
    plan: medico.plan,
    gestoria: tieneGestoria(medico),
    requiereCambioPassword: Boolean(okTemp || !auth.esHash(guardada))
  };
}

// ─── Cambiar contraseña del médico ────────────────────────────
// Se guarda cifrada en K. En el Sheet solo queda visible: PORTAL ACTIVADO = SÍ y la fecha del cambio.
async function wa_cambiarPasswordMedico(rfc, passwordActual, passwordNueva) {
  const login = await wa_loginMedico(rfc, passwordActual);
  if (!login.ok) return { ok: false, error: login.error && login.error.startsWith('Demasiados') ? login.error : 'La contraseña actual no coincide' };

  const err = auth.validarNuevaPassword(passwordNueva, login.rfc);
  if (err) return { ok: false, error: err };
  if (String(passwordNueva) === String(passwordActual)) return { ok: false, error: 'La nueva contraseña debe ser distinta a la actual.' };

  const medico = await buscarCuenta(login.rfc);
  const idx = medico.idx;
  await medico.set(idx.pass, auth.hashPassword(passwordNueva));
  await medico.set(idx.temp, '');
  await medico.set(idx.vence, '');
  await medico.set(idx.activado, 'SÍ');
  await medico.set(idx.fechaCambio, auth.ahoraMx());
  return { ok: true };
}

// ─── Enviar acceso (temporal) por correo ──────────────────────
// tipo 'bienvenida' (desde el sistema, vence en 7 días) o 'recuperacion' (desde el portal, vence en 24 h).
// La contraseña actual del médico NO se borra: la temporal solo es una llave extra.
async function enviarAccesoMedico(medico, idx, tipo) {
  if (!medico.correo || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(medico.correo)) {
    return { ok: false, error: `${medico.nombre} no tiene un correo válido en la hoja EMISORES (columna CORREO).` };
  }
  const esBienvenida = tipo === 'bienvenida';
  const horas = esBienvenida ? 24 * 7 : 24;
  const temporal = auth.generarTemporal();

  await medico.set(idx.temp, auth.hashPassword(temporal));
  await medico.set(idx.vence, new Date(Date.now() + horas * 3600 * 1000).toISOString());
  if (idx.activado >= 0 && !celda(medico.row, idx.activado)) await medico.set(idx.activado, 'NO');

  const html = esBienvenida
    ? auth.plantillaBienvenida({ nombre: medico.nombre, rfc: medico.rfc, temporal })
    : auth.plantillaCorreo({
    titulo: esBienvenida ? 'Tu acceso al portal de OGMpay' : 'Recupera tu acceso a OGMpay',
    saludo: `Hola, ${medico.nombre}:`,
    cuerpo: esBienvenida
      ? 'Ya puedes consultar en línea el estado de tus facturas y descargar tu reporte mensual (PDF, XML y Excel). Estos son tus datos de acceso:'
      : 'Recibimos una solicitud para recuperar tu acceso al portal. Usa esta contraseña temporal para entrar:',
    rfc: medico.rfc,
    temporal,
    vigencia: esBienvenida ? '7 días' : '24 horas'
  });
  await auth.enviarCorreo(medico.correo, esBienvenida ? 'Bienvenido a OGMpay — tu acceso al Portal médico' : 'Contraseña temporal — OGMpay', html);
  await medico.set(idx.ultimoEnvio, auth.ahoraMx() + (esBienvenida ? ' (acceso)' : ' (recuperación)'));
  return { ok: true, correo: medico.correo };
}

// Desde el sistema interno: botón "Enviar acceso"
async function wa_enviarAccesoPortal(rfc) {
  const { idx, medico } = await buscarMedicoPortal(rfc);
  if (!medico) return { ok: false, error: 'RFC no encontrado en EMISORES' };
  medico.set = (k, v) => k < 0 ? null : setCeldaEmisor(medico.fila, k, v);
  return await enviarAccesoMedico(medico, idx, 'bienvenida');
}

// Desde el portal: "¿Olvidaste tu contraseña?"
// Siempre responde lo mismo, exista o no el RFC (para no revelar quién es cliente).
const ultimaRecuperacion = new Map();
async function wa_recuperarPasswordMedico(rfc) {
  const rfcB = normUsuario(rfc);
  const respuesta = { ok: true, mensaje: 'Si tu usuario está registrado, en unos minutos te llegará un correo con una contraseña temporal. Revisa también tu carpeta de spam.' };
  if (!rfcB) return respuesta;
  const ultima = ultimaRecuperacion.get(rfcB) || 0;
  if (Date.now() - ultima < 10 * 60 * 1000) { console.log('Recuperación: se pidió hace menos de 10 min para', rfcB); return respuesta; } // máximo 1 correo cada 10 min
  ultimaRecuperacion.set(rfcB, Date.now());
  try {
    const medico = await buscarCuenta(rfcB);
    if (medico) {
      const r = await enviarAccesoMedico(medico, medico.idx, 'recuperacion');
      if (!r.ok) console.error('Recuperación sin envío:', r.error);
      else console.log('Recuperación enviada a', r.correo);
    } else {
      console.log('Recuperación: no se encontró ninguna cuenta con', rfcB);
    }
  } catch (e) {
    console.error('Error en recuperación de contraseña:', e.message);
  }
  return respuesta;
}

// Para la pestaña "Acceso portal" del sistema interno
async function wa_getEstadoPortalMedicos() {
  const { data, idx } = await leerEmisoresPortal();
  const lista = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rfc = celda(row, 1).toUpperCase();
    const nombre = celda(row, 2);
    if (!rfc || !nombre) continue;
    const vence = Date.parse(celda(row, idx.vence));
    const encE = (data[0] || []).map(normHeader);
    const iCel = encE.findIndex(h => /CELULAR|TELEFONO|WHATSAPP/.test(h));
    lista.push({
      rfc, nombre,
      cp: celda(row, 3), regimen: celda(row, 4), serie: celda(row, 5), tasaIsr: tasaIsrPorRegimen(celda(row, 4)),
      plan: (celda(row, idx.plan) || 'ESENCIAL').toUpperCase(),
      celular: iCel >= 0 ? celda(row, iCel) : '',
      correo: celda(row, idx.correo),
      activado: celda(row, idx.activado).toUpperCase().startsWith('S'),
      fechaCambio: celda(row, idx.fechaCambio),
      ultimoEnvio: celda(row, idx.ultimoEnvio),
      temporalVigente: !isNaN(vence) && vence > Date.now(),
      gestoria: celda(row, idx.gestoria).toUpperCase().startsWith('S'),
      gestoriaFecha: celda(row, idx.gestoriaFecha)
    });
  }
  return { ok: true, medicos: lista };
}

// ─── Facturas filtradas por RFC del médico ────────────────────
async function wa_getFacturasMedico(rfc, contrasena) {
  const login = await wa_loginMedico(rfc, contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const rfcBuscado = login.rfc;

  const emisores = await leerRango('EMISORES!A:K');
  let nombreMedico = null;
  for (let i = 1; i < emisores.length; i++) {
    const rfcFila = (emisores[i][1] || '').toString().trim().toUpperCase();
    if (rfcFila === rfcBuscado) {
      nombreMedico = (emisores[i][2] || '').toString().trim();
      break;
    }
  }
  if (!nombreMedico) return [];

  const todas = await wa_getFacturas();
  return todas.filter(f => f.medico.toUpperCase() === nombreMedico.toUpperCase());
}

// ─── Enviar aclaración ─────────────────────────────────────────
async function wa_enviarAclaracion(params) {
  // Solo el médico dueño de la factura puede pedir aclaración; los datos del correo salen de la hoja, no del navegador
  const login = await wa_loginMedico(params.rfc, params.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const fila = Number(params.fila);
  const mensaje = String(params.mensaje || '').trim().slice(0, 1500);
  if (!mensaje) return { ok: false, error: 'Escribe el motivo de la aclaración.' };
  const factura = (await wa_getFacturas()).find(f => f.fila === fila && f.medico.toUpperCase() === String(login.nombre || '').toUpperCase());
  if (!factura) return { ok: false, error: 'No encontré esa factura en tu cuenta.' };
  const medico = login.nombre, rfc = login.rfc, aseguradora = factura.aseguradora, monto = factura.montoFinal, estado = factura.estado;

  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];

  try {
    const colObs = col(headers, 'OBSERVACIONES');
    const letraObs = columnLetter(colObs + 1);
    const obsActual = (data[fila - 1][colObs] || '').toString().trim();
    const nuevaObs = (obsActual ? obsActual + ' | ' : '') +
      `⚠ ACLARACIÓN (${new Date().toLocaleDateString('es-MX')}): ${mensaje}`;
    await escribirCelda(`Registro!${letraObs}${fila}`, nuevaObs);
  } catch (e) {
    console.error('No se pudo marcar observación:', e.message);
  }

  try {
    await enviarCorreoAclaracion({ fila, medico, rfc, aseguradora, monto, estado, mensaje });
  } catch (e) {
    console.error('Error enviando correo de aclaración:', e.message);
  }

  return { ok: true };
}

async function enviarCorreoAclaracion({ fila, medico, rfc, aseguradora, monto, estado, mensaje }) {
  const RESEND_API_KEY = process.env.RESEND_API_KEY;
  if (!RESEND_API_KEY) {
    console.log('⚠️ RESEND_API_KEY no configurado — correo no enviado, solo quedó marcado en el Sheet');
    return;
  }

  const fetch = require('node-fetch');
  const esc = v => String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const html = `
    <h2>⚠️ Nueva aclaración — OGMpay</h2>
    <p><strong>Médico:</strong> ${esc(medico)} (${esc(rfc)})</p>
    <p><strong>Fila en Sheet:</strong> ${fila}</p>
    <p><strong>Aseguradora:</strong> ${esc(aseguradora)}</p>
    <p><strong>Monto:</strong> $${Number(monto).toLocaleString('es-MX', {minimumFractionDigits:2})}</p>
    <p><strong>Estado actual:</strong> ${esc(estado)}</p>
    <hr>
    <p><strong>Mensaje del médico:</strong></p>
    <p>${esc(mensaje)}</p>
  `;

  await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${RESEND_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      from: 'OGMpay <notificaciones@ogmpay.com>',
      to: ['facturacion@ogmpay.com'],
      subject: `⚠️ Aclaración: ${medico} — ${aseguradora}`,
      html
    })
  });
}

// ─── Obtener archivos de una factura (base64) para descarga local ──
async function wa_getArchivosFactura(fila, tipo) {
  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];
  const row = data[fila - 1];
  if (!row) return { ok: false, error: 'Fila no encontrada' };
  const esComp = tipo === 'complemento';

  let id = (row[col(headers, esComp ? 'ID FACTURAMA COMPLEMENTO' : 'ID FACTURAMA')] || '').toString().trim();
  const uuid = (row[col(headers, esComp ? 'Folio Complemento(UUID)' : 'Folio Factura (UUID)')] || '').toString().trim();
  if (!id && uuid) {
    id = await buscarPorUUID(uuid);
  }
  if (!id) return { ok: false, error: esComp ? 'Esta factura no tiene complemento emitido' : 'Esta factura no tiene ID de Facturama' };

  const medico = (row[col(headers, 'EMISOR (MEDICO)')] || '').toString().trim();
  const receptor = (row[col(headers, 'NOMBRE FISCAL')] || '').toString().trim();
  const serie = (row[col(headers, 'ID FACTURAMA')] ? '' : '');
  const concepto = (row[col(headers, 'CONCEPTO DE FACTURA 1')] || '').toString();
  const partes = concepto.split('_');
  const paciente = partes.length > 1 ? partes[partes.length - 1].trim().replace(/[^\w\s]/gi, '') : 'FACTURA';
  const nombreBase = (esComp ? 'COMPLEMENTO_' : '') + (uuid ? uuid.substring(0, 8) : 'factura') + '_' + (paciente || 'archivo');

  const xmlBuf = await descargarArchivo(id, 'xml');
  const pdfBuf = await descargarArchivo(id, 'pdf');

  return {
    ok: true,
    nombreBase: nombreBase,
    xml: xmlBuf ? xmlBuf.toString('base64') : null,
    pdf: pdfBuf ? pdfBuf.toString('base64') : null
  };
}

// ─── REPORTE MENSUAL (médico y administración) ────────────────
// Núcleo compartido: lista los CFDIs (facturas + complementos) emitidos por un RFC del 1 al último día del mes.
async function listarCfdisDelMes(rfc, mes, anio) {
  mes = Number(mes); anio = Number(anio);
  if (!(mes >= 1 && mes <= 12) || !(anio >= 2020 && anio <= 2100)) return { ok: false, error: 'Mes o año inválido' };

  const mm = String(mes).padStart(2, '0');
  const ultimoDia = new Date(anio, mes, 0).getDate();
  const inicio = `${anio}-${mm}-01`;
  const fin = `${anio}-${mm}-${String(ultimoDia).padStart(2, '0')}`;

  const lista = await listarCfdisEmitidos(rfc, inicio, fin);

  const cfdis = lista
    .map(c => {
      const fecha = String(c.Date || c.CreatedAt || '');
      const estado = String(c.Status || '').toLowerCase();
      return {
        id: c.Id || c.id,
        uuid: c.Uuid || c.UUID || '',
        fecha: fecha,
        receptor: c.TaxName || c.Name || '',
        rfcReceptor: c.Rfc || '',
        total: Number(c.Total) || 0,
        cancelado: estado === 'canceled' || estado === 'cancelled' || c.IsActive === false
      };
    })
    // Seguro extra: solo lo que de verdad cae en el mes pedido (si Facturama trae la fecha)
    .filter(c => !c.fecha || c.fecha.slice(0, 7) === `${anio}-${mm}`);

  return { ok: true, cfdis };
}

// Núcleo compartido: PDF + XML de un CFDI, solo si el emisor del XML es el RFC indicado.
async function archivoCfdiDeRfc(rfc, idCfdi) {
  const id = String(idCfdi || '').trim();
  if (!id) return { ok: false, error: 'Falta el id del comprobante' };

  const xmlBuf = await descargarArchivo(id, 'xml');
  if (!xmlBuf) return { ok: false, error: 'No se pudo descargar el XML' };

  const xmlTxt = xmlBuf.toString('utf8');
  const m = xmlTxt.match(/<cfdi:Emisor\b[^>]*\bRfc="([^"]+)"/i);
  const rfcEmisor = m ? m[1].trim().toUpperCase() : '';
  if (rfcEmisor !== String(rfc || '').trim().toUpperCase()) return { ok: false, error: 'Este comprobante no pertenece a ese RFC' };

  const pdfBuf = await descargarArchivo(id, 'pdf');
  return { ok: true, xml: xmlBuf.toString('base64'), pdf: pdfBuf ? pdfBuf.toString('base64') : null };
}

// Portal médico: requiere RFC + contraseña del médico
async function wa_getReporteMensualMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  if (login.plan !== 'cliente') return { ok: false, error: 'El reporte mensual es para clientes de OGMpay.' };
  return await listarCfdisDelMes(login.rfc, p.mes, p.anio);
}
async function wa_getArchivoCfdiMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  return await archivoCfdiDeRfc(login.rfc, p.id);
}

// Sistema interno (Resumen Fiscal): requiere usuario + contraseña del sistema
async function wa_getReporteMensualAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const rfc = String(p.rfc || '').trim().toUpperCase();
  if (!rfc) return { ok: false, error: 'Falta el RFC del médico' };
  return await listarCfdisDelMes(rfc, p.mes, p.anio);
}
async function wa_getArchivoCfdiAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  return await archivoCfdiDeRfc(p.rfc, p.id);
}

// ─── Obtener link de la carpeta de Drive de una factura ──
async function wa_getLinkDrive(fila) {
  const data = await leerRango('Registro!A:AZ');
  const headers = data[0];
  const row = data[fila - 1];
  if (!row) return { ok: false, error: 'Fila no encontrada' };
  const medico = (row[col(headers, 'EMISOR (MEDICO)')] || '').toString().trim();
  const receptor = (row[col(headers, 'NOMBRE FISCAL')] || '').toString().trim();
  return await getLinkCarpetaFactura(medico, receptor);
}

// ═════════════════════════════════════════════════════════════
// PACIENTES POR COBRAR (cartas de honorarios por solicitar)
// Hoja "SOLICITUDES". El médico registra al paciente egresado
// con su liga personal (sin contraseña, solo puede AGREGAR) o desde
// su portal; OGMpay pide la carta al hospital y la cobra.
// ═════════════════════════════════════════════════════════════
const crypto = require('crypto');
const HOJA_SOL = 'SOLICITUDES';
const ENC_SOL = ['ID', 'FECHA CAPTURA', 'RFC MEDICO', 'MEDICO', 'PACIENTE', 'HOSPITAL', 'FECHA EGRESO',
  'ASEGURADORA', 'NOTAS DEL MEDICO', 'ESTADO', 'FECHA ACTUALIZACION', 'FACTURA / FOLIO',
  'COMENTARIO PARA EL MEDICO', 'ORIGEN', 'PIDE SEGUIMIENTO OGM', 'TIPO DE ATENCION', 'PROCEDIMIENTO', 'PARTICIPACION', 'FECHA PROCEDIMIENTO'];
const TIPOS_ATENCION = ['Cirugía / Procedimiento', 'Atención médica'];
const PARTICIPACIONES = ['Cirujano', 'Anestesiólogo', 'Primer ayudante', 'Segundo ayudante', 'Tercer ayudante', 'Visita médica'];
// Catálogo base (del Google Form "Pacientes egresados"); se suman los que ya se han capturado
const HOSPITALES_BASE = ['Roma', 'Metropolitano', 'Clínica Londres', 'Mocel', 'Lindavista', 'Acoxpa', 'Santa Mónica', 'CAMI', 'Bité',
  'Durango', 'México', 'Universidad', 'CIME', 'Star Médica Centro', 'DioMed', 'MAC La Viga', 'MAC Periférico'];
const ASEGURADORAS_BASE = ['Allianz', 'Atlas', 'AXA', 'Banorte', 'BBVA', 'BUPA', 'BX+', 'GNP', 'General de Seguros', 'Inbursa', 'MetLife',
  'Mapfre', 'Monterrey', 'Prevem', 'Plan Seguro', 'SURA', 'Zurich'];
const ULT_COL_SOL = columnLetter(ENC_SOL.length); // N
const ESTADOS_SOL = ['POR SOLICITAR', 'SOLICITADA', 'RECIBIDA', 'FACTURADA', 'PAGADA', 'NO PROCEDE'];
const URL_CAPTURA = 'https://app.ogmpay.com/captura/?t=';

// Evita que dos registros simultáneos escriban en la misma fila
let colaSol = Promise.resolve();
function enCola(fn) {
  const r = colaSol.then(fn, fn);
  colaSol = r.catch(() => {});
  return r;
}

function limpiarTexto(v, max) {
  return String(v == null ? '' : v)
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .replace(/^[=+\-@\s]+/, '')          // sin fórmulas en el Sheet
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max);
}

async function leerSolicitudes() {
  let data;
  try {
    data = await leerRango(`${HOJA_SOL}!A:${ULT_COL_SOL}`);
  } catch (e) {
    throw new Error('Falta la pestaña SOLICITUDES en el Google Sheet. Créala (vacía) con ese nombre exacto.');
  }
  data = data || [];
  if (!data.length || ENC_SOL.some((h, i) => normHeader(celda(data[0], i)) !== normHeader(h))) {
    await escribirFila(`${HOJA_SOL}!A1:${ULT_COL_SOL}1`, ENC_SOL);
    if (!data.length) data = [ENC_SOL.slice()];
    else data[0] = ENC_SOL.slice();
  }
  return data;
}

function filaASolicitud(row, i) {
  return {
    fila: i + 1,
    id: celda(row, 0),
    fechaCaptura: celda(row, 1),
    rfc: celda(row, 2).toUpperCase(),
    medico: celda(row, 3),
    paciente: celda(row, 4),
    hospital: celda(row, 5),
    fechaEgreso: celda(row, 6),
    aseguradora: celda(row, 7),
    notas: celda(row, 8),
    estado: (celda(row, 9) || 'POR SOLICITAR').toUpperCase(),
    fechaActualizacion: celda(row, 10),
    folio: celda(row, 11),
    comentario: celda(row, 12),
    origen: celda(row, 13),
    pideSeguimiento: celda(row, 14),
    tipo: celda(row, 15),
    procedimiento: celda(row, 16),
    participacion: celda(row, 17),
    fechaProcedimiento: celda(row, 18)
  };
}

// Valida y normaliza los datos de un paciente (registro, edición e importación).
// Obligatorios (decisión de Adrián, 11-oct): nombre, hospital, fecha del procedimiento, aseguradora y participación.
function fechaISOaMx(v) {
  const t = limpiarTexto(v, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) return '';
  const [y, m, d] = t.split('-');
  if (+y < 2000 || +y > 2100 || +m < 1 || +m > 12 || +d < 1 || +d > 31) return '';
  return `${d}/${m}/${y}`;
}
function normalizarSolicitud(p) {
  const paciente = limpiarTexto(p.paciente, 120);
  const hospital = limpiarTexto(p.hospital, 80);
  const fechaProc = fechaISOaMx(p.fechaProcedimiento);
  const aseguradora = limpiarTexto(p.aseguradora, 80);
  const participacion = PARTICIPACIONES.find(x => x === p.participacion) || '';
  if (paciente.length < 3) return { ok: false, error: 'Falta el nombre del paciente.' };
  if (!hospital) return { ok: false, error: 'Falta el hospital.' };
  if (!fechaProc) return { ok: false, error: 'Falta la fecha del procedimiento (o no es válida).' };
  if (!aseguradora) return { ok: false, error: 'Falta la aseguradora (o "No sé").' };
  if (!participacion) return { ok: false, error: 'Falta tu participación.' };
  return {
    ok: true, paciente, hospital, fechaProc, aseguradora, participacion,
    egreso: fechaISOaMx(p.fechaEgreso),
    tipo: TIPOS_ATENCION.find(t => t === p.tipo) || '',
    procedimiento: limpiarTexto(p.procedimiento, 150),
    notas: limpiarTexto(p.notas, 500)
  };
}
// Las fechas se guardan como texto ('dd/mm/aaaa) para que Google no las voltee según el idioma de la hoja
function fechaTexto(v) { return v ? "'" + v : ''; }
function filaNuevaSolicitud(medico, d, origen, estado) {
  const id = 'P' + Date.now().toString(36).toUpperCase() + crypto.randomBytes(2).toString('hex').toUpperCase();
  return [
    id, auth.ahoraMx(), medico.rfc, medico.nombre, d.paciente, d.hospital, fechaTexto(d.egreso),
    d.aseguradora, d.notas,
    estado || 'POR SOLICITAR', auth.ahoraMx(), '', '', origen, '', d.tipo, d.procedimiento, d.participacion, fechaTexto(d.fechaProc)
  ];
}

async function agregarSolicitud(medico, p, origen) {
  const d = normalizarSolicitud(p);
  if (!d.ok) return d;
  return enCola(async () => {
    const data = await leerSolicitudes();
    const fila = filaNuevaSolicitud(medico, d, origen);
    const n = data.length + 1;
    await escribirFila(`${HOJA_SOL}!A${n}:${ULT_COL_SOL}${n}`, fila);
    avisarNuevoPaciente(medico, { paciente: d.paciente, hospital: d.hospital, egreso: d.fechaProc, origen, tipo: d.tipo, aseguradora: d.aseguradora, procedimiento: d.procedimiento, participacion: d.participacion, notas: d.notas })
      .catch(e => console.error('Aviso paciente:', e.message));
    return { ok: true, id: fila[0] };
  });
}

// ── Editar datos de un paciente ya registrado (no cambia estado ni folio) ──
async function editarDatosSolicitud(id, claveMedico, p) {
  const d = normalizarSolicitud(p);
  if (!d.ok) return d;
  return enCola(async () => {
    const data = await leerSolicitudes();
    const i = data.findIndex((r, k) => k > 0 && celda(r, 0) === String(id || '') && (!claveMedico || mismaClave(celda(r, 2), claveMedico)));
    if (i < 1) return { ok: false, error: 'No encontré ese paciente.' };
    const n = i + 1;
    await escribirFila(`${HOJA_SOL}!E${n}:I${n}`, [d.paciente, d.hospital, fechaTexto(d.egreso), d.aseguradora, d.notas]);
    await escribirFila(`${HOJA_SOL}!P${n}:S${n}`, [d.tipo, d.procedimiento, d.participacion, fechaTexto(d.fechaProc)]);
    await escribirCelda(`${HOJA_SOL}!K${n}`, auth.ahoraMx());
    return { ok: true };
  });
}
async function wa_editarPacienteMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  return await editarDatosSolicitud(p.id, login.rfc, p);
}
async function wa_editarSolicitudAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  return await editarDatosSolicitud(p.id, null, p);
}

// ── Importar pacientes desde Excel/CSV (el portal manda lotes de hasta 300) ──
// Todas las filas válidas del lote se escriben en UNA sola llamada a Google Sheets.
const ESTADOS_IMPORT = ESTADOS_SOL.slice();
async function wa_importarPacientesMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const filas = Array.isArray(p.filas) ? p.filas.slice(0, 300) : [];
  if (!filas.length) return { ok: false, error: 'No hay filas para importar.' };
  const medico = { rfc: login.rfc, nombre: login.nombre };
  const errores = [], nuevas = [];
  for (const f of filas) {
    const d = normalizarSolicitud(f);
    if (!d.ok) { errores.push({ n: f.n, error: d.error }); continue; }
    const estado = ESTADOS_IMPORT.includes(String(f.estado || '').toUpperCase()) ? String(f.estado).toUpperCase() : 'POR SOLICITAR';
    nuevas.push(filaNuevaSolicitud(medico, d, 'IMPORTACION', estado));
  }
  if (!nuevas.length) return { ok: true, agregados: 0, errores };
  try {
    await enCola(async () => {
      const data = await leerSolicitudes();
      const desde = data.length + 1, hasta = data.length + nuevas.length;
      await escribirFilas(`${HOJA_SOL}!A${desde}:${ULT_COL_SOL}${hasta}`, nuevas);
    });
  } catch (e) {
    return { ok: false, error: 'No se pudo guardar en la hoja: ' + e.message };
  }
  return { ok: true, agregados: nuevas.length, errores };
}

async function avisarNuevoPaciente(medico, s) {
  if (!process.env.RESEND_API_KEY) return;
  const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const html = `<h2>Nuevo paciente por solicitar</h2>
    <p><strong>Médico:</strong> ${esc(medico.nombre)} (${esc(medico.rfc)})</p>
    <p><strong>Paciente:</strong> ${esc(s.paciente)}<br><strong>Hospital:</strong> ${esc(s.hospital)}<br>
    <strong>Fecha del procedimiento:</strong> ${esc(s.egreso || 'sin fecha')}<br><strong>Tipo:</strong> ${esc(s.tipo)}${s.procedimiento ? ' · ' + esc(s.procedimiento) : ''}<br><strong>Participación:</strong> ${esc(s.participacion)}<br>
    <strong>Aseguradora:</strong> ${esc(s.aseguradora)}<br><strong>Registrado desde:</strong> ${esc(s.origen)}</p>
    ${s.notas ? '<p><strong>Comentarios:</strong> ' + esc(s.notas) + '</p>' : ''}
    <p>Está en la pestaña "Por solicitar" del sistema y en la hoja SOLICITUDES.</p>`;
  await auth.enviarCorreo('facturacion@ogmpay.com', `Paciente por solicitar: ${medico.nombre} - ${s.hospital}`, html);
}

// ── Liga personal de captura ──
function generarTokenCaptura() { return crypto.randomBytes(18).toString('base64url'); }

async function ligaDeCuenta(cuenta, regenerar) {
  let token = celda(cuenta.row, cuenta.idx.token);
  if (!token && !regenerar && cuenta.plan === 'cliente') {
    // si venía de cuenta gratis, conserva la misma liga que ya tiene guardada en su celular
    const gratis = await buscarCuentaGratis(cuenta.rfc);
    if (gratis) token = celda(gratis.row, IDX_CUENTA.token);
    if (token) await cuenta.set(cuenta.idx.token, token);
  }
  if (!token || regenerar) {
    token = generarTokenCaptura();
    await cuenta.set(cuenta.idx.token, token);
  }
  return { ok: true, url: URL_CAPTURA + token, nombre: cuenta.nombre };
}

async function medicoPorToken(token) {
  const t = String(token || '').trim();
  if (t.length < 20) return null;
  const { data, idx } = await leerEmisoresPortal();
  const tb = Buffer.from(t);
  const igual = g => g && g.length === t.length && crypto.timingSafeEqual(Buffer.from(g), tb);
  for (let i = 1; i < data.length; i++) {
    if (igual(celda(data[i], idx.token))) {
      return { fila: i + 1, row: data[i], rfc: celda(data[i], 1).toUpperCase(), nombre: celda(data[i], 2) };
    }
  }
  const cuentas = await leerCuentas();
  for (let i = 1; cuentas && i < cuentas.length; i++) {
    if (igual(celda(cuentas[i], IDX_CUENTA.token))) {
      const c = cuentaGratisDeFila(cuentas, i);
      // si ya es cliente, se registra a su nombre de EMISORES
      const { medico } = c.rfcReal ? await buscarMedicoPortal(c.rfcReal) : { medico: null };
      return medico ? { fila: medico.fila, row: medico.row, rfc: medico.rfc, nombre: medico.nombre } : c;
    }
  }
  return null;
}

const capturasPorToken = new Map(); // límite: 40 registros por hora por liga
function excedeLimite(token) {
  const ahora = Date.now();
  const lista = (capturasPorToken.get(token) || []).filter(t => ahora - t < 3600 * 1000);
  if (lista.length >= 40) return true;
  lista.push(ahora); capturasPorToken.set(token, lista);
  return false;
}

async function hospitalesConocidos() {
  try {
    const data = await leerSolicitudes();
    const set = new Map(HOSPITALES_BASE.map(h => [normHeader(h), h]));
    for (let i = 1; i < data.length; i++) {
      const h = celda(data[i], 5);
      if (h && set.has(normHeader(h))) continue;
      if (h) set.set(normHeader(h), h);
    }
    return [...set.values()].sort((a, b) => a.localeCompare(b, 'es'));
  } catch (e) { return HOSPITALES_BASE.slice(); }
}

async function aseguradorasCatalogo() {
  let lista = [];
  try { lista = (await wa_getAseguradoras()).map(a => a.nombre); } catch (e) {}
  const set = new Map(ASEGURADORAS_BASE.map(a => [normHeader(a), a]));
  lista.forEach(a => { if (a && !set.has(normHeader(a))) set.set(normHeader(a), a); });
  return [...set.values()].sort((a, b) => a.localeCompare(b, 'es'));
}

// Liga sin contraseña: solo muestra el nombre del médico y catálogos; nunca lista pacientes.
async function wa_capturaInfo(p) {
  const medico = await medicoPorToken(p.t);
  if (!medico) return { ok: false, error: 'Esta liga ya no es válida. Pide una nueva a OGMpay.' };
  const aseguradoras = await aseguradorasCatalogo();
  let gestoria = false;
  try { gestoria = tieneGestoria(await buscarCuenta(medico.rfc)); } catch (e) {}
  return { ok: true, nombre: medico.nombre, gestoria, hospitales: await hospitalesConocidos(), aseguradoras, tipos: TIPOS_ATENCION, participaciones: PARTICIPACIONES };
}

async function wa_capturaPaciente(p) {
  const medico = await medicoPorToken(p.t);
  if (!medico) return { ok: false, error: 'Esta liga ya no es válida. Pide una nueva a OGMpay.' };
  if (excedeLimite(p.t)) return { ok: false, error: 'Demasiados registros en poco tiempo. Intenta en un rato.' };
  return await agregarSolicitud(medico, p, 'LIGA');
}

// ── Portal médico (RFC + contraseña) ──
async function wa_getPacientesMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  if (login.plan === 'cliente') { try { await migrarCuentasConvertidas(); } catch (e) { console.error('Migración cuentas:', e.message); } }
  const data = await leerSolicitudes();
  const lista = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const s = filaASolicitud(data[i], i);
    if (s.id && mismaClave(s.rfc, login.rfc)) { delete s.fila; delete s.rfc; delete s.medico; lista.push(s); }
  }
  return { ok: true, pacientes: lista, gestoria: login.gestoria, hospitales: await hospitalesConocidos(), aseguradoras: await aseguradorasCatalogo(), tipos: TIPOS_ATENCION, participaciones: PARTICIPACIONES };
}

async function wa_registrarPacienteMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  return await agregarSolicitud({ rfc: login.rfc, nombre: login.nombre }, p, 'PORTAL');
}

async function wa_getLigaCapturaMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const cuenta = await buscarCuenta(login.rfc);
  return await ligaDeCuenta(cuenta, Boolean(p.regenerar));
}

// ── Sistema interno (usuario + contraseña del sistema) ──
async function wa_getSolicitudesAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  try { await migrarCuentasConvertidas(); } catch (e) { console.error('Migración cuentas:', e.message); }
  const data = await leerSolicitudes();
  const emisores = await leerRango('EMISORES!A:C');
  const clientes = new Set(emisores.slice(1).map(r => celda(r, 1).toUpperCase()).filter(Boolean));
  const conGestoria = await clavesConGestoria();
  const lista = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const s = filaASolicitud(data[i], i);
    if (s.id) { s.gratis = !clientes.has(s.rfc); s.gestoria = conGestoria.has(s.rfc); lista.push(s); }
  }
  return { ok: true, solicitudes: lista, estados: ESTADOS_SOL, hospitales: await hospitalesConocidos(), aseguradoras: await aseguradorasCatalogo(), tipos: TIPOS_ATENCION, participaciones: PARTICIPACIONES };
}

async function wa_actualizarSolicitud(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const estado = String(p.estado || '').toUpperCase();
  if (!ESTADOS_SOL.includes(estado)) return { ok: false, error: 'Estado no válido' };
  return enCola(async () => {
    const data = await leerSolicitudes();
    let i = data.findIndex((r, k) => k > 0 && celda(r, 0) === String(p.id || ''));
    if (i < 1) return { ok: false, error: 'No encontré ese registro (¿lo borraron del Sheet?)' };
    const n = i + 1;
    await escribirFila(`${HOJA_SOL}!J${n}:M${n}`, [estado, auth.ahoraMx(), limpiarTexto(p.folio, 80), limpiarTexto(p.comentario, 300)]);
    return { ok: true };
  });
}

async function wa_registrarPacienteAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const medico = await buscarCuenta(p.rfcMedico);
  if (!medico) return { ok: false, error: 'Elige al médico.' };
  return await agregarSolicitud(medico, p, 'OGM');
}

async function wa_getLigaCapturaAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const cuenta = await buscarCuenta(p.rfc);
  if (!cuenta) return { ok: false, error: 'RFC no encontrado' };
  return await ligaDeCuenta(cuenta, Boolean(p.regenerar));
}

// ═════════════════════════════════════════════════════════════
// CUENTA GRATIS: registro, control propio y "¿quieres que demos seguimiento?"
// ═════════════════════════════════════════════════════════════
let registrosRecientes = [];

async function wa_registrarCuentaGratis(p) {
  const ahora = Date.now();
  registrosRecientes = registrosRecientes.filter(t => ahora - t < 3600 * 1000);
  if (registrosRecientes.length >= 30) return { ok: false, error: 'Estamos recibiendo muchos registros. Intenta en un rato.' };

  const nombre = limpiarTexto(p.nombre, 100);
  const correo = String(p.correo || '').trim().toLowerCase().slice(0, 120);
  const celular = String(p.celular || '').replace(/\D/g, '').slice(-10);
  const nac = String(p.fechaNacimiento || '').trim();
  if (nombre.length < 5 || nombre.indexOf(' ') < 0) return { ok: false, error: 'Escribe tu nombre completo.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) return { ok: false, error: 'Revisa tu correo.' };
  if (celular.length !== 10) return { ok: false, error: 'Escribe tu teléfono a 10 dígitos.' };
  const m = nac.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  const edad = m ? (Date.now() - new Date(+m[1], +m[2] - 1, +m[3]).getTime()) / (365.25 * 86400000) : 0;
  if (!m || edad < 20 || edad > 95) return { ok: false, error: 'Revisa tu fecha de nacimiento.' };
  if (!p.aceptoAviso) return { ok: false, error: 'Necesitas aceptar el aviso de privacidad.' };
  const errPass = auth.validarNuevaPassword(p.contrasena, correo);
  if (errPass) return { ok: false, error: errPass };

  return enCola(async () => {
    const cuentas = await leerCuentas();
    if (!cuentas) return { ok: false, error: 'El registro aún no está disponible. Escríbenos por WhatsApp.' };
    if (cuentas.slice(1).some(r => mismaClave(celda(r, 2), correo))) {
      return { ok: false, error: 'Ya existe una cuenta con ese correo. Entra con tu contraseña o usa "¿Olvidaste tu contraseña?".' };
    }
    const { data: em, idx: ie } = await leerEmisoresPortal();
    if (em.slice(1).some(r => mismaClave(celda(r, ie.correo), correo))) {
      return { ok: false, error: 'Ya existe una cuenta con ese correo. Entra con tu contraseña o usa "¿Olvidaste tu contraseña?".' };
    }
    registrosRecientes.push(ahora);
    const n = cuentas.length + 1;
    const fila = ['', nombre, correo, celular, `${m[3]}/${m[2]}/${m[1]}`, auth.hashPassword(p.contrasena), auth.ahoraMx(), 'SÍ ' + auth.ahoraMx(),
      generarTokenCaptura(), '', '', '', ''];
    await escribirFila(`${HOJA_CUENTAS}!A${n}:${ULT_COL_CUENTAS}${n}`, fila);

    const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    auth.enviarCorreo('facturacion@ogmpay.com', `Nueva cuenta gratis: ${nombre}`,
      `<h2>Nueva cuenta gratis en OGMpay</h2><p><strong>${esc(nombre)}</strong><br>Correo: ${esc(correo)}<br>Teléfono: ${esc(celular)}</p>
       <p>Es un posible cliente: ya puede registrar pacientes. Lo ves en el sistema, pestaña "Acceso portal" → Cuentas gratis.</p>`)
      .catch(e => console.error('Aviso cuenta gratis:', e.message));
    auth.enviarCorreo(correo, 'Tu cuenta gratis de OGMpay está lista', correoCuentaGratis(nombre, correo))
      .catch(e => console.error('Bienvenida cuenta gratis:', e.message));
    return { ok: true, usuario: correo };
  });
}

function correoCuentaGratis(nombre, usuario) {
  const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  return `
<div style="font-family:Arial,Helvetica,sans-serif;background:#F4F7FB;padding:24px;">
  <div style="max-width:520px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #E3E9F2;">
    <div style="background:#0F2A52;padding:20px 28px;">
      <span style="color:#ffffff;font-size:22px;font-weight:800;">OGM</span><span style="color:#00B67A;font-size:22px;font-weight:300;">pay</span>
    </div>
    <div style="padding:28px;color:#0F2A52;font-size:15px;line-height:1.55;">
      <h2 style="margin:0 0 12px;font-size:19px;">Bienvenido a OGMpay</h2>
      <p style="margin:0 0 12px;">Hola, ${esc(nombre)}:</p>
      <p style="margin:0 0 12px;">Tu cuenta gratis ya está lista. Con ella puedes:</p>
      <ul style="margin:0 0 18px;padding-left:20px;">
        <li>Registrar a cada paciente en segundos desde tu celular.</li>
        <li>Llevar el control de qué cartas te faltan y cuáles ya te pagaron.</li>
        <li>Pedirnos que demos seguimiento cuando tú no tengas tiempo.</li>
      </ul>
      <p style="margin:0 0 18px;">Entra con tu correo <strong>${esc(usuario)}</strong> y la contraseña que creaste. En la pestaña <strong>Mis pacientes</strong> encontrarás tu liga para guardarla en la pantalla de inicio de tu celular.</p>
      <p style="text-align:center;margin:0 0 22px;">
        <a href="https://app.ogmpay.com/medico/" style="background:#00B67A;color:#062A1E;text-decoration:none;padding:12px 26px;border-radius:10px;font-weight:700;display:inline-block;">Entrar a mi cuenta</a>
      </p>
      <p style="margin:0;font-size:12px;color:#8FA3C0;">¿Dudas? Escríbenos por <a href="https://wa.me/529617079590" style="color:#2E6CC7;">WhatsApp</a>.</p>
    </div>
  </div>
</div>`;
}

// Cuando Olga convierte una cuenta gratis (le pone RFC en CUENTAS y la da de alta en EMISORES),
// sus pacientes guardados con el correo pasan a su RFC. Se ejecuta al abrir "Por solicitar" o "Mis pacientes".
async function migrarCuentasConvertidas() {
  const cuentas = await leerCuentas();
  if (!cuentas) return;
  const emisores = await leerRango('EMISORES!A:C');
  const nombrePorRfc = new Map(emisores.slice(1).map(r => [celda(r, 1).toUpperCase(), celda(r, 2)]));
  const cambio = new Map(); // correo -> { rfc, nombre }
  for (let i = 1; i < cuentas.length; i++) {
    const rfc = celda(cuentas[i], 0).toUpperCase(), correo = celda(cuentas[i], 2).toUpperCase();
    if (rfc && correo && nombrePorRfc.has(rfc)) cambio.set(correo, { rfc, nombre: nombrePorRfc.get(rfc) });
  }
  if (!cambio.size) return;
  // si contrató la gestoría siendo cuenta gratis, se conserva al volverse cliente
  const { data: em, idx: ie } = await leerEmisoresPortal();
  for (let i = 1; i < cuentas.length; i++) {
    const rfc = celda(cuentas[i], 0).toUpperCase();
    if (!rfc || !celda(cuentas[i], IDX_CUENTA.gestoria).toUpperCase().startsWith('S')) continue;
    const k = em.findIndex((r, j) => j > 0 && celda(r, 1).toUpperCase() === rfc);
    if (k > 0 && !celda(em[k], ie.gestoria)) {
      await setCeldaEmisor(k + 1, ie.gestoria, 'SÍ');
      await setCeldaEmisor(k + 1, ie.gestoriaFecha, celda(cuentas[i], IDX_CUENTA.gestoriaFecha));
      await setCeldaEmisor(k + 1, ie.gestoriaVersion, celda(cuentas[i], IDX_CUENTA.gestoriaVersion));
    }
  }
  await enCola(async () => {
    const data = await leerSolicitudes();
    for (let i = 1; i < data.length; i++) {
      const c = cambio.get(celda(data[i], 2).toUpperCase());
      if (c) await escribirFila(`${HOJA_SOL}!C${i + 1}:D${i + 1}`, [c.rfc, c.nombre]);
    }
  });
}

async function solicitudDelMedico(data, id, rfc) {
  const i = data.findIndex((r, k) => k > 0 && celda(r, 0) === String(id || '') && mismaClave(celda(r, 2), rfc));
  return i;
}

// La cuenta gratis lleva su propio control: ella misma cambia el estado
async function wa_actualizarPacienteMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  if (login.gestoria) return { ok: false, error: 'Tienes la Gestoría de cartas activa: OGMpay actualiza el estado de tus pacientes por ti.' };
  const estado = String(p.estado || '').toUpperCase();
  if (!ESTADOS_SOL.includes(estado)) return { ok: false, error: 'Estado no válido' };
  return enCola(async () => {
    const data = await leerSolicitudes();
    const i = await solicitudDelMedico(data, p.id, login.rfc);
    if (i < 1) return { ok: false, error: 'No encontré ese paciente.' };
    await escribirFila(`${HOJA_SOL}!J${i + 1}:K${i + 1}`, [estado, auth.ahoraMx()]);
    return { ok: true };
  });
}

// "¿Quieres que nosotros demos seguimiento?"
async function wa_pedirSeguimientoPaciente(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const res = await enCola(async () => {
    const data = await leerSolicitudes();
    const i = await solicitudDelMedico(data, p.id, login.rfc);
    if (i < 1) return { ok: false, error: 'No encontré ese paciente.' };
    if (celda(data[i], 14)) return { ok: true, yaEstaba: true };
    await escribirCelda(`${HOJA_SOL}!O${i + 1}`, auth.ahoraMx());
    return { ok: true, s: filaASolicitud(data[i], i) };
  });
  if (res.ok && res.s) {
    const cuenta = await buscarCuenta(login.rfc);
    if (cuenta && cuenta.plan === 'gratis') await cuenta.set(IDX_CUENTA.seguimiento, auth.ahoraMx());
    const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    auth.enviarCorreo('facturacion@ogmpay.com', `Pide seguimiento: ${login.nombre} - ${res.s.paciente}`,
      `<h2>Un médico quiere que demos seguimiento</h2><p><strong>${esc(login.nombre)}</strong> (${esc(login.rfc)}) ${login.plan === 'gratis' ? '<strong>— cuenta gratis, posible cliente</strong>' : ''}${cuenta && cuenta.celular ? '<br>Celular: ' + esc(cuenta.celular) : ''}</p>
       <p>Paciente: ${esc(res.s.paciente)}<br>Hospital: ${esc(res.s.hospital)}<br>Egreso: ${esc(res.s.fechaEgreso || 'sin fecha')}</p>
       <p>Contáctalo por WhatsApp para explicarle cómo trabajamos.</p>`)
      .catch(e => console.error('Aviso seguimiento:', e.message));
  }
  return { ok: true };
}

// Para Olga: lista de cuentas gratis (posibles clientes)
async function wa_getCuentasGratisAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const cuentas = await leerCuentas();
  if (!cuentas) return { ok: true, cuentas: [], falta: true };
  const emisores = await leerRango('EMISORES!A:C');
  const clientes = new Set(emisores.slice(1).map(r => celda(r, 1).toUpperCase()).filter(Boolean));
  let sol = [];
  try { sol = await leerSolicitudes(); } catch (e) {}
  const lista = [];
  for (let i = cuentas.length - 1; i >= 1; i--) {
    const c = cuentaGratisDeFila(cuentas, i);
    if (!c.rfc) continue;
    const suyas = sol.slice(1).filter(r => mismaClave(celda(r, 2), c.rfc) || (c.rfcReal && mismaClave(celda(r, 2), c.rfcReal)));
    lista.push({
      rfc: c.rfcReal, nombre: c.nombre, correo: c.correo, celular: c.celular,
      nacimiento: celda(c.row, IDX_CUENTA.nacimiento),
      fechaRegistro: celda(c.row, IDX_CUENTA.fecha),
      pidioSeguimiento: celda(c.row, IDX_CUENTA.seguimiento),
      pacientes: suyas.length,
      sinCarta: suyas.filter(r => ['POR SOLICITAR', 'SOLICITADA'].includes((celda(r, 9) || 'POR SOLICITAR').toUpperCase())).length,
      yaEsCliente: Boolean(c.rfcReal && clientes.has(c.rfcReal)),
      gestoria: tieneGestoria(c),
      clave: c.rfc
    });
  }
  return { ok: true, cuentas: lista };
}

// ═════════════════════════════════════════════════════════════
// GESTORÍA DE CARTAS (servicio adicional, $200 por carta conseguida)
// El médico la contrata en su portal aceptando electrónicamente el texto.
// Se guarda: GESTORIA = SÍ, GESTORIA ACEPTADA (fecha y hora), GESTORIA VERSION.
// Con eso se genera la "Carta de autorización" (sin firma autógrafa).
// ═════════════════════════════════════════════════════════════
const PRECIO_GESTORIA = 200;
const GESTORIA_VERSION = 'v2-2026-11';
const TEXTO_GESTORIA = [
  'Autorizo a OGMpay a solicitar, dar seguimiento y recibir en mi nombre mis cartas de honorarios médicos, así como a gestionar su cobro, ante hospitales, aseguradoras y sus áreas administrativas, respecto de los pacientes que yo registre en mi cuenta OGMpay.',
  `Acepto que OGMpay me cobre $${PRECIO_GESTORIA} MXN por cada carta de honorarios que consiga. Si OGMpay no consigue la carta, no se me cobra nada. Si en 6 meses desde el registro del paciente el hospital no entrega la carta, el seguimiento se cierra sin costo.`,
  'El cobro se incluye en mi estado de cuenta mensual de OGMpay. Puedo cancelar este servicio en cualquier momento escribiendo a OGMpay; la cancelación no afecta las cartas ya conseguidas.'
];

function tieneGestoria(cuenta) {
  if (!cuenta || !cuenta.idx || cuenta.idx.gestoria == null || cuenta.idx.gestoria < 0) return false;
  return celda(cuenta.row, cuenta.idx.gestoria).toUpperCase().startsWith('S');
}

async function clavesConGestoria() {
  const set = new Set();
  const { data, idx } = await leerEmisoresPortal();
  for (let i = 1; i < data.length; i++) if (celda(data[i], idx.gestoria).toUpperCase().startsWith('S')) set.add(celda(data[i], 1).toUpperCase());
  const cuentas = await leerCuentas();
  for (let i = 1; cuentas && i < cuentas.length; i++) {
    if (celda(cuentas[i], IDX_CUENTA.gestoria).toUpperCase().startsWith('S')) set.add(cuentaGratisDeFila(cuentas, i).rfc.toUpperCase());
  }
  return set;
}

function folioAceptacion(clave, fecha, version) {
  return crypto.createHash('sha256').update(`${String(clave).toUpperCase()}|${fecha}|${version}`).digest('hex').slice(0, 12).toUpperCase();
}

async function wa_getInfoGestoria() {
  return { ok: true, precio: PRECIO_GESTORIA, version: GESTORIA_VERSION, texto: TEXTO_GESTORIA };
}

async function wa_contratarGestoria(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  if (!p.acepto || p.version !== GESTORIA_VERSION) return { ok: false, error: 'Necesitas aceptar la autorización para contratar.' };
  const cuenta = await buscarCuenta(login.rfc);
  if (!cuenta) return { ok: false, error: 'No encontré tu cuenta.' };
  if (tieneGestoria(cuenta)) return { ok: true, yaActiva: true };
  const fecha = auth.ahoraMx();
  await cuenta.set(cuenta.idx.gestoria, 'SÍ');
  await cuenta.set(cuenta.idx.gestoriaFecha, fecha);
  await cuenta.set(cuenta.idx.gestoriaVersion, GESTORIA_VERSION);
  const folio = folioAceptacion(cuenta.rfc, fecha, GESTORIA_VERSION);

  const esc = v => String(v || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  auth.enviarCorreo('facturacion@ogmpay.com', `Gestoría contratada: ${cuenta.nombre}`,
    `<h2>Nuevo contrato de Gestoría de cartas</h2><p><strong>${esc(cuenta.nombre)}</strong> (${esc(cuenta.rfc)}) ${cuenta.plan === 'gratis' ? '— cuenta gratis' : ''}${cuenta.celular ? '<br>Teléfono: ' + esc(cuenta.celular) : ''}</p>
     <p>Aceptado electrónicamente el ${esc(fecha)} · versión ${GESTORIA_VERSION} · folio ${folio}</p>
     <p>Sus pacientes ya aparecen en "Por solicitar". La carta de autorización se descarga desde "Acceso portal".</p>`)
    .catch(e => console.error('Aviso gestoría:', e.message));
  if (cuenta.correo) {
    auth.enviarCorreo(cuenta.correo, 'Gestoría de cartas activada — OGMpay',
      `<div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;color:#0F2A52;line-height:1.55;">
        <h2 style="font-size:19px;">Tu Gestoría de cartas está activa</h2>
        <p>Hola, ${esc(cuenta.nombre)}:</p>
        <p>A partir de hoy nosotros conseguimos las cartas de honorarios de los pacientes que registres. Verás cómo va cada una en tu portal, pestaña <strong>Mis pacientes</strong>.</p>
        <p style="font-size:13px;color:#4A6080;"><strong>Lo que aceptaste:</strong><br>${TEXTO_GESTORIA.map(esc).join('<br><br>')}</p>
        <p style="font-size:12px;color:#8FA3C0;">Aceptado electrónicamente el ${esc(fecha)} · folio ${folio}</p>
      </div>`)
      .catch(e => console.error('Confirmación gestoría:', e.message));
  }
  return { ok: true, fecha, folio };
}

function datosCarta(cuenta) {
  const fecha = celda(cuenta.row, cuenta.idx.gestoriaFecha);
  const version = celda(cuenta.row, cuenta.idx.gestoriaVersion) || GESTORIA_VERSION;
  return {
    ok: true, nombre: cuenta.nombre, rfc: cuenta.plan === 'cliente' ? cuenta.rfc : (cuenta.rfcReal || ''),
    correo: cuenta.correo || '', fecha, version, folio: folioAceptacion(cuenta.rfc, fecha, version),
    // la carta que se entrega al hospital solo lleva la autorización (sin precios ni términos comerciales)
    texto: [TEXTO_GESTORIA[0]]
  };
}

async function wa_getCartaAutorizacionMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const cuenta = await buscarCuenta(login.rfc);
  if (!tieneGestoria(cuenta)) return { ok: false, error: 'No tienes la Gestoría de cartas activa.' };
  return datosCarta(cuenta);
}

async function wa_getCartaAutorizacionAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const cuenta = await buscarCuenta(p.clave || p.rfcMedico);
  if (!tieneGestoria(cuenta)) return { ok: false, error: 'Este médico no tiene la Gestoría de cartas activa.' };
  return datosCarta(cuenta);
}

// ═════════════════════════════════════════════════════════════
// TIEMPOS DE PAGO POR ASEGURADORA (estadística con todas las facturas)
// Días naturales entre "Fecha de subida al portal" y "FECHA DE TRANSFERENCIA".
// Solo devuelve números agregados (no expone datos de otros médicos).
// ═════════════════════════════════════════════════════════════
function claveAseguradora(a) {
  return String(a || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim()
    .replace(/^SEGUROS\s+/, '').replace(/\s+SEGUROS$/, '').replace(/\s+/g, ' ');
}
function percentil(arr, q) {
  const a = arr.slice().sort((x, y) => x - y);
  const pos = (a.length - 1) * q, b = Math.floor(pos);
  return a[b + 1] !== undefined ? a[b] + (pos - b) * (a[b + 1] - a[b]) : a[b];
}
let cacheTiempos = { t: 0, data: null };
async function wa_getTiemposPago(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  const admin = !login.ok && (await verificarAdmin(p));
  if (!login.ok && !admin) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  if (cacheTiempos.data && Date.now() - cacheTiempos.t < 30 * 60 * 1000) return cacheTiempos.data;
  const facturas = await wa_getFacturas();
  const grupos = {}, todos = [];
  for (const f of facturas) {
    const subida = fechaIsoCob(f.fechaSubida), pago = fechaIsoCob(f.fechaPago);
    if (!subida || !pago) continue;
    const dias = Math.round((Date.parse(pago) - Date.parse(subida)) / 86400000);
    if (!(dias >= 0 && dias <= 200)) continue; // descarta fechas capturadas mal
    const k = claveAseguradora(f.aseguradora);
    if (!k) continue;
    (grupos[k] = grupos[k] || []).push(dias);
    todos.push(dias);
  }
  const porAseguradora = {};
  for (const [k, arr] of Object.entries(grupos)) {
    porAseguradora[k] = { n: arr.length, mediana: Math.round(percentil(arr, 0.5)), p80: Math.ceil(percentil(arr, 0.8)) };
  }
  const res = {
    ok: true, minimo: 3, porAseguradora,
    global: todos.length ? { n: todos.length, mediana: Math.round(percentil(todos, 0.5)), p80: Math.ceil(percentil(todos, 0.8)) } : null
  };
  cacheTiempos = { t: Date.now(), data: res };
  return res;
}

// ─── Helper: número de columna → letra ───────────────────────
function columnLetter(n) {
  let result = '';
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}


// ════════════════════════════════════════════════════════════
// ALTA DE NUEVO CLIENTE (desde el sistema → pestaña Acceso portal)
// El sistema lee la Constancia de Situación Fiscal en el navegador y manda
// solo los datos. El CSD (si se incluye) se sube directo a Facturama y NO se
// guarda en ningún lado (ni en el Sheet ni en Drive).
// ════════════════════════════════════════════════════════════
const RE_RFC = /^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$/;

async function subirCsdFacturama(rfc, cerB64, keyB64, password) {
  const fetch = require('node-fetch');
  const { getAuth } = require('./facturama');
  const BASE_F = process.env.FACTURAMA_BASE || 'https://api.facturama.mx';
  const cuerpo = JSON.stringify({ Rfc: rfc, Certificate: cerB64, PrivateKey: keyB64, PrivateKeyPassword: password });
  const opts = m => ({ method: m, headers: { 'Authorization': getAuth(), 'Content-Type': 'application/json', 'Accept': 'application/json' }, body: cuerpo });
  let res = await fetch(`${BASE_F}/api-lite/csds`, opts('POST'));
  let raw = await res.text();
  if (res.status === 200 || res.status === 201) return { ok: true, accion: 'cargado' };
  // Si ese RFC ya tenía un CSD en Facturama, se reemplaza
  if (/exist|registrad/i.test(raw)) {
    res = await fetch(`${BASE_F}/api-lite/csds/${encodeURIComponent(rfc)}`, opts('PUT'));
    raw = await res.text();
    if (res.status === 200 || res.status === 201 || res.status === 204) return { ok: true, accion: 'actualizado' };
  }
  let msg = raw;
  try { const j = JSON.parse(raw); msg = j.Message || j.message || (j.ModelState ? JSON.stringify(j.ModelState) : raw); } catch (e) {}
  return { ok: false, error: 'Facturama no aceptó el CSD: ' + String(msg).slice(0, 300) };
}

async function wa_altaClienteAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const rfc = String(p.rfc || '').trim().toUpperCase();
  const nombre = String(p.nombre || '').trim().replace(/\s+/g, ' ').toUpperCase();
  const cp = String(p.cp || '').trim();
  const regimen = String(p.regimen || '').trim();
  const serie = String(p.serie || '').trim().toUpperCase();
  const correo = String(p.correo || '').trim().toLowerCase();
  const celular = String(p.celular || '').replace(/\D/g, '');
  const plan = ['ESENCIAL', 'INTEGRAL', 'SIN COBRO'].includes(String(p.plan || '').toUpperCase()) ? String(p.plan).toUpperCase() : 'ESENCIAL';
  if (!RE_RFC.test(rfc)) return { ok: false, error: 'El RFC no es válido.' };
  if (nombre.length < 5) return { ok: false, error: 'Falta el nombre como aparece en la constancia.' };
  if (!/^\d{5}$/.test(cp)) return { ok: false, error: 'El código postal debe tener 5 dígitos.' };
  if (!/^\d{3}$/.test(regimen)) return { ok: false, error: 'Elige el régimen fiscal.' };
  if (!/^[A-Z0-9]{1,10}$/.test(serie)) return { ok: false, error: 'La serie debe tener de 1 a 10 letras o números, sin espacios.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) return { ok: false, error: 'El correo no es válido.' };

  const { data, idx } = await leerEmisoresPortal();
  // Columna CELULAR (se crea sola si no existe)
  const enc = (data[0] || []).map(normHeader);
  let idxCel = enc.findIndex(h => /CELULAR|TELEFONO|WHATSAPP/.test(h));
  if (idxCel < 0 && celular) {
    idxCel = Math.max(enc.length, IDX_PASS + 1, ...Object.values(idx).map(v => v + 1));
    await escribirCelda(`EMISORES!${columnLetter(idxCel + 1)}1`, 'CELULAR');
  }
  for (let i = 1; i < data.length; i++) {
    if (celda(data[i], 1).toUpperCase() === rfc) return { ok: false, error: 'Ese RFC ya está dado de alta (fila ' + (i + 1) + ' de EMISORES).' };
    if (celda(data[i], 2).toUpperCase() === nombre) return { ok: false, error: 'Ya hay un emisor con ese mismo nombre (fila ' + (i + 1) + ').' };
  }

  // 1) CSD a Facturama (opcional). Si falla, no se da de alta para poder corregir y reintentar.
  let csd = null;
  if (p.csd && (p.csd.cer || p.csd.key || p.csd.password)) {
    if (!p.csd.cer || !p.csd.key || !p.csd.password) return { ok: false, error: 'Para el CSD faltan archivos o la contraseña (.cer, .key y contraseña).' };
    csd = await subirCsdFacturama(rfc, String(p.csd.cer), String(p.csd.key), String(p.csd.password));
    if (!csd.ok) return csd;
  }

  // 2) Fila nueva en EMISORES (A=nombre, B=RFC, C=nombre fiscal, D=CP, E=régimen, F=serie, CORREO, CELULAR)
  let fila;
  await enCola(async () => {
    const actual = await leerRango('EMISORES!A:Z');
    fila = actual.length + 1;
    const ult = Math.max(5, idx.correo, idx.plan, celular ? idxCel : 0);
    const row = new Array(ult + 1).fill('');
    row[0] = nombre; row[1] = rfc; row[2] = nombre; row[3] = "'" + cp; row[4] = regimen; row[5] = serie;
    row[idx.correo] = correo;
    row[idx.plan] = plan;
    if (celular) row[idxCel] = "'" + celular;
    await escribirFila(`EMISORES!A${fila}:${columnLetter(ult + 1)}${fila}`, row);
  });

  // 3) Si ya tenía cuenta gratis con ese correo, se liga para que conserve pacientes, contraseña y liga
  let cuentaGratis = false;
  try {
    const cuentas = await leerCuentas();
    if (cuentas) {
      for (let i = 1; i < cuentas.length; i++) {
        if (celda(cuentas[i], 2).toLowerCase() === correo && !celda(cuentas[i], 0)) {
          await escribirCelda(`${HOJA_CUENTAS}!A${i + 1}`, rfc);
          cuentaGratis = true;
          break;
        }
      }
      if (cuentaGratis) await migrarCuentasConvertidas();
    }
  } catch (e) { console.error('Alta: no se pudo ligar la cuenta gratis:', e.message); }

  // 4) Enviar acceso al portal (opcional)
  let acceso = null;
  if (p.enviarAcceso && !cuentaGratis) {
    try { acceso = await wa_enviarAccesoPortal(rfc); } catch (e) { acceso = { ok: false, error: e.message }; }
  }
  console.log(`Alta de cliente: ${nombre} (${rfc}) en fila ${fila}${csd ? ', CSD ' + csd.accion : ''}${cuentaGratis ? ', venía de cuenta gratis' : ''}`);
  return { ok: true, fila, csd: csd ? csd.accion : null, cuentaGratis, acceso, correoEnJ: idx.correo === 9 };
}


// ─── Editar datos de un médico (desde el sistema) ─────────────
// El RFC no se cambia (es la llave de todo). Si cambia el nombre, se actualiza
// también en Registro (A y C) y en SOLICITUDES (D) para no perder su historial.
async function wa_editarClienteAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const rfc = String(p.rfc || '').trim().toUpperCase();
  const nombre = String(p.nombre || '').trim().replace(/\s+/g, ' ').toUpperCase();
  const cp = String(p.cp || '').trim();
  const regimen = String(p.regimen || '').trim();
  const serie = String(p.serie || '').trim().toUpperCase();
  const correo = String(p.correo || '').trim().toLowerCase();
  const celular = String(p.celular || '').replace(/\D/g, '');
  if (nombre.length < 5) return { ok: false, error: 'Falta el nombre como aparece en la constancia.' };
  if (!/^\d{5}$/.test(cp)) return { ok: false, error: 'El código postal debe tener 5 dígitos.' };
  if (!/^\d{3}$/.test(regimen)) return { ok: false, error: 'Elige el régimen fiscal.' };
  if (!/^[A-Z0-9]{1,10}$/.test(serie)) return { ok: false, error: 'La serie debe tener de 1 a 10 letras o números, sin espacios.' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) return { ok: false, error: 'El correo no es válido.' };
  const plan = ['ESENCIAL', 'INTEGRAL', 'SIN COBRO'].includes(String(p.plan || '').toUpperCase()) ? String(p.plan).toUpperCase() : '';

  const { data, idx } = await leerEmisoresPortal();
  const k = data.findIndex((r, i) => i > 0 && celda(r, 1).toUpperCase() === rfc);
  if (k < 0) return { ok: false, error: 'No encontré ese RFC en EMISORES.' };
  const nombreAnterior = celda(data[k], 2);
  const cambiaNombre = nombreAnterior.toUpperCase() !== nombre;
  if (cambiaNombre && data.some((r, i) => i > 0 && i !== k && celda(r, 2).toUpperCase() === nombre)) {
    return { ok: false, error: 'Ya hay otro emisor con ese nombre.' };
  }

  let csd = null;
  if (p.csd && (p.csd.cer || p.csd.key || p.csd.password)) {
    if (!p.csd.cer || !p.csd.key || !p.csd.password) return { ok: false, error: 'Para el CSD faltan archivos o la contraseña (.cer, .key y contraseña).' };
    csd = await subirCsdFacturama(rfc, String(p.csd.cer), String(p.csd.key), String(p.csd.password));
    if (!csd.ok) return csd;
  }

  const fila = k + 1;
  const encE = (data[0] || []).map(normHeader);
  let idxCel = encE.findIndex(h => /CELULAR|TELEFONO|WHATSAPP/.test(h));
  if (idxCel < 0 && celular) {
    idxCel = Math.max(encE.length, IDX_PASS + 1, ...Object.values(idx).map(v => v + 1));
    await escribirCelda(`EMISORES!${columnLetter(idxCel + 1)}1`, 'CELULAR');
  }
  const cambios = [];
  const sheets = await getSheets();
  const upd = [
    { range: `EMISORES!C${fila}:F${fila}`, values: [[nombre, "'" + cp, regimen, serie]] },
    { range: `EMISORES!${columnLetter(idx.correo + 1)}${fila}`, values: [[correo]] }
  ];
  if (cambiaNombre) upd.push({ range: `EMISORES!A${fila}`, values: [[nombre]] });
  if (plan) { if ((celda(data[k], idx.plan) || 'ESENCIAL').toUpperCase() !== plan) cambios.push('plan'); upd.push({ range: `EMISORES!${columnLetter(idx.plan + 1)}${fila}`, values: [[plan]] }); }
  if (idxCel >= 0) upd.push({ range: `EMISORES!${columnLetter(idxCel + 1)}${fila}`, values: [[celular ? "'" + celular : '']] });
  if (celda(data[k], 3) !== cp) cambios.push('CP');
  if (celda(data[k], 4) !== regimen) cambios.push('régimen');
  if (celda(data[k], 5).toUpperCase() !== serie) cambios.push('serie');
  if (celda(data[k], idx.correo).toLowerCase() !== correo) cambios.push('correo');
  if (idxCel >= 0 && celda(data[k], idxCel).replace(/\D/g, '') !== celular) cambios.push('celular');

  // Si cambia el nombre, se arrastra a sus facturas y pacientes
  let facturasRenombradas = 0;
  if (cambiaNombre) {
    cambios.unshift('nombre');
    const reg = await leerRango('Registro!A:C');
    const viejo = nombreAnterior.toUpperCase();
    for (let i = 1; i < reg.length; i++) {
      if (celda(reg[i], 0).toUpperCase() === viejo || (celda(reg[i], 1).toUpperCase() === rfc && celda(reg[i], 2).toUpperCase() === viejo)) {
        upd.push({ range: `Registro!A${i + 1}`, values: [[nombre]] });
        upd.push({ range: `Registro!C${i + 1}`, values: [[nombre]] });
        facturasRenombradas++;
      }
    }
    try {
      const sol = await leerSolicitudes();
      for (let i = 1; i < sol.length; i++) {
        if (celda(sol[i], 2).toUpperCase() === rfc) upd.push({ range: `${HOJA_SOL}!D${i + 1}`, values: [[nombre]] });
      }
    } catch (e) { /* sin hoja de pacientes */ }
  }
  if (csd) cambios.push('CSD ' + csd.accion);

  await enCola(async () => {
    // Google permite muchas celdas por llamada; se manda en bloques para no exceder límites
    for (let i = 0; i < upd.length; i += 400) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data: upd.slice(i, i + 400) }
      });
    }
  });
  console.log(`Datos editados: ${nombre} (${rfc}) → ${cambios.join(', ') || 'sin cambios'}`);
  return { ok: true, cambios, facturasRenombradas, tasaIsr: tasaIsrPorRegimen(regimen) };
}


// ════════════════════════════════════════════════════════════
// COBRANZA: estados de cuenta mensuales de OGMpay a cada médico
// ════════════════════════════════════════════════════════════
// Hoja COBRANZA = un estado de cuenta por médico y mes (se crea sola).
// Hoja CONFIG   = datos de cobro y facturación de OGMpay (clave | valor).
// Qué se cobra en el estado del mes P (solo lo que aún no se cobró en otro estado):
//   · COMISIÓN: % sobre el BRUTO de cada factura cobrada (fecha de transferencia ≤ fin de P y ≥ COBRO DESDE).
//   · GESTORÍA: $ por cada carta conseguida (paciente en RECIBIDA/FACTURADA/PAGADA) de médicos con gestoría.
//   · INTEGRAL: cuota mensual si el médico tiene PLAN = INTEGRAL.
// Médicos con PLAN = SIN COBRO no generan estado de cuenta.
// Crea una pestaña en el Spreadsheet si no existe
async function asegurarHoja(titulo) {
  try { await leerRango(`${titulo}!A1:A1`); return; } catch (e) {}
  const sheets = await getSheets();
  try {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: titulo } } }] } });
    console.log('Pestaña creada:', titulo);
  } catch (e) { if (!/already exists|ya existe/i.test(e.message)) throw e; }
}
const HOJA_COB = 'COBRANZA';
const ENC_COB = ['ID', 'PERIODO', 'RFC', 'MEDICO', 'SUBTOTAL', 'IVA', 'TOTAL', 'ESTADO', 'FECHA GENERADO', 'FECHA ENVIO',
  'PAGO REPORTADO', 'FECHA TRANSFERENCIA', 'METODO PAGO', 'REFERENCIA PAGO', 'FACTURA UUID', 'ID FACTURAMA', 'DETALLE', 'NOTAS'];
const ULT_COL_COB = 'R';
const IC = { id: 0, periodo: 1, rfc: 2, medico: 3, subtotal: 4, iva: 5, total: 6, estado: 7, generado: 8, envio: 9,
  reportado: 10, fechaPago: 11, metodo: 12, refPago: 13, uuid: 14, idFact: 15, detalle: 16, notas: 17 };

const HOJA_CONFIG = 'CONFIG';
const CONFIG_DEFAULT = {
  'COMISION %': '4',
  'CUOTA INTEGRAL': '1000',
  'PRECIO GESTORIA': '200',
  'IVA %': '16',
  'DIAS PARA PAGAR': '5',
  'COBRO DESDE': '',            // aaaa-mm-dd: no se cobran facturas pagadas antes de esta fecha
  'BANCO': '',
  'CLABE': '',
  'TITULAR': '',
  'FACTURAR AL PAGAR': 'SI',
  'OGM RFC': '',
  'OGM NOMBRE': '',
  'OGM CP': '',
  'OGM REGIMEN': '626',
  'OGM SERIE': 'OGM',
  'CLAVE PROD SERV': '84111506',   // respaldo si falta alguna de las de abajo
  'CLAVE COMISION': '84101704',    // Servicios de cobro de la deuda (servicios de cobranza)
  'CLAVE GESTORIA': '80161500',    // Servicios de apoyo gerencial
  'CLAVE INTEGRAL': '84111506',    // Servicios de facturación (cuota fija)
  'CLAVE UNIDAD': 'E48'
};
// Claves que el médico puede ver en su portal
const CONFIG_PUBLICA = ['BANCO', 'CLABE', 'TITULAR', 'DIAS PARA PAGAR'];

async function leerConfig() {
  await asegurarHoja(HOJA_CONFIG);
  let data = [];
  try { data = await leerRango(`${HOJA_CONFIG}!A:B`); } catch (e) { data = []; }
  const cfg = Object.assign({}, CONFIG_DEFAULT);
  const presentes = new Set();
  for (let i = 1; i < (data || []).length; i++) {
    const k = celda(data[i], 0).toUpperCase();
    if (k) { cfg[k] = celda(data[i], 1); presentes.add(k); }
  }
  // Si la hoja está vacía o le faltan claves, se escriben (para que se puedan ver y editar)
  const faltan = Object.keys(CONFIG_DEFAULT).filter(k => !presentes.has(k));
  if (faltan.length) {
    try {
      const filas = [];
      if (!data || !data.length) filas.push(['CLAVE', 'VALOR']);
      faltan.forEach(k => filas.push([k, CONFIG_DEFAULT[k]]));
      const desde = (data && data.length ? data.length : 0) + 1;
      await escribirFilas(`${HOJA_CONFIG}!A${desde}:B${desde + filas.length - 1}`, filas);
    } catch (e) { console.error('CONFIG: no se pudo inicializar (¿existe la pestaña CONFIG?):', e.message); }
  }
  return cfg;
}
function num(v, def = 0) { const n = Number(String(v == null ? '' : v).replace(/[$,\s]/g, '')); return isNaN(n) ? def : n; }
function r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

async function guardarConfig(valores) {
  let data = [];
  try { data = await leerRango(`${HOJA_CONFIG}!A:B`); } catch (e) { data = []; }
  if (!data || !data.length) { await leerConfig(); data = await leerRango(`${HOJA_CONFIG}!A:B`); }
  const upd = [];
  let siguiente = data.length + 1;
  for (const [k, v] of Object.entries(valores)) {
    const K = k.toUpperCase();
    if (!(K in CONFIG_DEFAULT)) continue;
    const i = data.findIndex((r, j) => j > 0 && celda(r, 0).toUpperCase() === K);
    const valor = (K === 'CLABE' || K === 'OGM CP') ? "'" + String(v || '').trim() : String(v == null ? '' : v).trim();
    if (i > 0) upd.push({ range: `${HOJA_CONFIG}!B${i + 1}`, values: [[valor]] });
    else { upd.push({ range: `${HOJA_CONFIG}!A${siguiente}:B${siguiente}`, values: [[K, valor]] }); siguiente++; }
  }
  if (!upd.length) return;
  const sheets = await getSheets();
  await sheets.spreadsheets.values.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { valueInputOption: 'USER_ENTERED', data: upd } });
}

// Fechas: acepta aaaa-mm-dd, dd/mm/aaaa, dd/mm/aa, dd.mm.aa (con o sin hora) → 'aaaa-mm-dd' o ''
function fechaIsoCob(v) {
  const t = String(v || '').replace(/^'/, '').trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
  if (/^\d{5}(\.\d+)?$/.test(t)) { // número de serie de Sheets/Excel
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(t)) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  return '';
}
function hoyIsoMx() {
  return new Date().toLocaleDateString('en-CA', { timeZone: 'America/Mexico_City' }); // aaaa-mm-dd
}
function finDePeriodo(periodo) { // 'aaaa-mm' → 'aaaa-mm-último día'
  const [y, m] = periodo.split('-').map(Number);
  return `${periodo}-${String(new Date(y, m, 0).getDate()).padStart(2, '0')}`;
}
const MESES_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
function nombrePeriodo(periodo) { const [y, m] = String(periodo).split('-').map(Number); return (MESES_ES[m - 1] || '') + ' ' + y; }

async function leerCobranza() {
  await asegurarHoja(HOJA_COB);
  let data;
  try { data = await leerRango(`${HOJA_COB}!A:${ULT_COL_COB}`); } catch (e) { data = []; }
  data = data || [];
  const enc = data[0] || [];
  if (ENC_COB.some((h, i) => normHeader(enc[i]) !== normHeader(h))) {
    await escribirFila(`${HOJA_COB}!A1:${ULT_COL_COB}1`, ENC_COB);
    if (!data.length) data.push(ENC_COB.slice()); else data[0] = ENC_COB.slice();
  }
  return data;
}
function filaACobro(row, i) {
  let detalle = [];
  try { detalle = JSON.parse(celda(row, IC.detalle) || '[]'); } catch (e) { detalle = []; }
  return {
    fila: i + 1, id: celda(row, IC.id), periodo: celda(row, IC.periodo).replace(/^'/, ''), rfc: celda(row, IC.rfc).toUpperCase(),
    medico: celda(row, IC.medico), subtotal: num(celda(row, IC.subtotal)), iva: num(celda(row, IC.iva)), total: num(celda(row, IC.total)),
    estado: (celda(row, IC.estado) || 'PENDIENTE').toUpperCase(), generado: celda(row, IC.generado), envio: celda(row, IC.envio),
    reportado: celda(row, IC.reportado), fechaPago: celda(row, IC.fechaPago).replace(/^'/, ''), metodo: celda(row, IC.metodo),
    refPago: celda(row, IC.refPago), uuid: celda(row, IC.uuid), idFacturama: celda(row, IC.idFact), detalle, notas: celda(row, IC.notas)
  };
}
function filaDeCobro(c) {
  const row = new Array(ENC_COB.length).fill('');
  row[IC.id] = c.id; row[IC.periodo] = "'" + c.periodo; row[IC.rfc] = c.rfc; row[IC.medico] = c.medico;
  row[IC.subtotal] = c.subtotal; row[IC.iva] = c.iva; row[IC.total] = c.total; row[IC.estado] = c.estado;
  row[IC.generado] = c.generado || ''; row[IC.envio] = c.envio || ''; row[IC.reportado] = c.reportado || '';
  row[IC.fechaPago] = c.fechaPago ? "'" + c.fechaPago : ''; row[IC.metodo] = c.metodo || ''; row[IC.refPago] = c.refPago || '';
  row[IC.uuid] = c.uuid || ''; row[IC.idFact] = c.idFacturama || ''; row[IC.detalle] = JSON.stringify(c.detalle || []); row[IC.notas] = c.notas || '';
  return row;
}
async function guardarCobro(c) {
  await escribirFila(`${HOJA_COB}!A${c.fila}:${ULT_COL_COB}${c.fila}`, filaDeCobro(c));
}

// Médicos que se cobran (EMISORES) con su plan y gestoría
async function medicosParaCobro() {
  const { data, idx } = await leerEmisoresPortal();
  const lista = [];
  for (let i = 1; i < data.length; i++) {
    const rfc = celda(data[i], 1).toUpperCase(), nombre = celda(data[i], 2);
    if (!rfc || !nombre) continue;
    const plan = (celda(data[i], idx.plan) || 'ESENCIAL').toUpperCase();
    lista.push({
      rfc, nombre, plan, serie: celda(data[i], 5), cp: celda(data[i], 3), regimen: celda(data[i], 4),
      correo: celda(data[i], idx.correo),
      gestoria: celda(data[i], idx.gestoria).toUpperCase().startsWith('S'),
      gestoriaDesde: fechaIsoCob(celda(data[i], idx.gestoriaFecha))
    });
  }
  return lista;
}

// Calcula los renglones a cobrar de un médico en un periodo (sin repetir lo ya cobrado en otros estados)
function calcularRenglones(med, periodo, cfg, facturas, solicitudes, yaCobrado) {
  const fin = finDePeriodo(periodo);
  const desde = fechaIsoCob(cfg['COBRO DESDE']);
  const pct = num(cfg['COMISION %'], 4) / 100;
  const renglones = [];
  if (med.plan === 'SIN COBRO') return renglones;

  for (const f of facturas) {
    if (String(f.medico || '').toUpperCase() !== med.nombre.toUpperCase()) continue;
    if (/CANCEL/.test(String(f.estado || '').toUpperCase())) continue;
    const pago = fechaIsoCob(f.fechaPago);
    if (!pago || pago > fin || (desde && pago < desde)) continue;
    const ref = f.uuid || ('F' + f.fila);
    if (yaCobrado.has('COM:' + ref)) continue;
    const bruto = r2(f.monto);
    if (!(bruto > 0)) continue;
    const paciente = (String(f.concepto1 || '').split('_').pop() || '').trim();
    renglones.push({
      tipo: 'COMISION', ref, fecha: pago,
      descripcion: `${(cfg['COMISION %'] || '4')}% de factura ${f.aseguradora || ''}${paciente && paciente !== f.concepto1 ? ' · ' + paciente : ''}`,
      base: bruto, importe: r2(bruto * pct)
    });
  }

  if (med.gestoria) {
    const precio = num(cfg['PRECIO GESTORIA'], 200);
    for (const s of solicitudes) {
      if (s.rfc !== med.rfc) continue;
      if (!['RECIBIDA', 'FACTURADA', 'PAGADA'].includes(s.estado)) continue;
      const cuando = fechaIsoCob(s.fechaActualizacion) || fechaIsoCob(s.fechaCaptura);
      if (cuando && cuando > fin) continue;
      if (med.gestoriaDesde && cuando && cuando < med.gestoriaDesde) continue;
      if (desde && cuando && cuando < desde) continue;
      if (yaCobrado.has('GES:' + s.id)) continue;
      renglones.push({ tipo: 'GESTORIA', ref: s.id, fecha: cuando, descripcion: `Carta conseguida · ${s.paciente}${s.hospital ? ' · ' + s.hospital : ''}`, base: 0, importe: r2(precio) });
    }
  }

  if (med.plan === 'INTEGRAL' && !yaCobrado.has('INT:' + periodo)) {
    renglones.push({ tipo: 'INTEGRAL', ref: periodo, fecha: '', descripcion: `Plan Integral · ${nombrePeriodo(periodo)}`, base: 0, importe: r2(num(cfg['CUOTA INTEGRAL'], 1000)) });
  }
  return renglones;
}
function claveCobro(r) { return (r.tipo === 'COMISION' ? 'COM:' : r.tipo === 'GESTORIA' ? 'GES:' : 'INT:') + r.ref; }

// Genera (o recalcula, si sigue pendiente y sin reporte de pago) los estados de cuenta de un periodo
async function wa_generarCobranza(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const periodo = String(p.periodo || '').trim();
  if (!/^\d{4}-\d{2}$/.test(periodo)) return { ok: false, error: 'Periodo no válido (aaaa-mm).' };
  const cfg = await leerConfig();
  const ivaPct = num(cfg['IVA %'], 16) / 100;
  const medicos = await medicosParaCobro();
  const facturas = await wa_getFacturas();
  let solicitudes = [];
  try { const sd = await leerSolicitudes(); solicitudes = sd.slice(1).map((r, i) => filaASolicitud(r, i + 1)).filter(s => s.id); } catch (e) {}

  const resumen = { creados: 0, actualizados: 0, sinCambios: 0, bloqueados: 0, sinCargos: 0, enviados: 0, errores: [] };
  const generados = [];
  await enCola(async () => {
    const data = await leerCobranza();
    const cobros = data.slice(1).map((r, i) => filaACobro(r, i + 1)).filter(c => c.id);
    for (const med of medicos) {
      if (med.plan === 'SIN COBRO') continue;
      const id = `EC-${(med.serie || med.rfc.slice(0, 4)).toUpperCase()}-${periodo.replace('-', '').slice(2)}`;
      const existente = cobros.find(c => c.id === id);
      if (existente && (existente.estado !== 'PENDIENTE' || existente.reportado)) { resumen.bloqueados++; continue; }
      // Lo ya cobrado en OTROS estados de cuenta no se vuelve a cobrar
      const yaCobrado = new Set();
      cobros.filter(c => c.rfc === med.rfc && c.id !== id && c.estado !== 'CANCELADO').forEach(c => c.detalle.forEach(r => yaCobrado.add(claveCobro(r))));
      const renglones = calcularRenglones(med, periodo, cfg, facturas, solicitudes, yaCobrado);
      if (!renglones.length) { if (!existente) resumen.sinCargos++; continue; }
      const subtotal = r2(renglones.reduce((s, r) => s + r.importe, 0));
      const iva = r2(subtotal * ivaPct);
      const cobro = Object.assign(existente || { fila: data.length + 1 + resumen.creados, generado: auth.ahoraMx() }, {
        id, periodo, rfc: med.rfc, medico: med.nombre, subtotal, iva, total: r2(subtotal + iva), estado: 'PENDIENTE', detalle: renglones
      });
      if (existente && JSON.stringify(existente.detalle) === JSON.stringify(renglones)) { resumen.sinCambios++; generados.push(cobro); continue; }
      await guardarCobro(cobro);
      if (existente) resumen.actualizados++; else resumen.creados++;
      generados.push(cobro);
    }
  });
  if (p.enviar) {
    for (const c of generados) {
      if (c.envio) continue;
      try { await enviarCorreoEstadoCuenta(c, cfg); resumen.enviados++; } catch (e) { resumen.errores.push(c.medico + ': ' + e.message); }
    }
  }
  console.log(`Cobranza ${periodo}:`, JSON.stringify(resumen));
  return Object.assign({ ok: true, periodo }, resumen);
}

function htmlEstadoCuenta(c, cfg) {
  const esc = v => String(v == null ? '' : v).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const fmt = n => '$' + Number(n || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const filas = c.detalle.map(r => `<tr><td style="padding:8px;border-bottom:1px solid #E5EAF2;font-size:13px;">${esc(r.descripcion)}${r.base ? `<div style="color:#8FA3C0;font-size:12px;">Factura por ${fmt(r.base)}${r.fecha ? ' · pagada el ' + esc(r.fecha.split('-').reverse().join('/')) : ''}</div>` : ''}</td><td style="padding:8px;border-bottom:1px solid #E5EAF2;text-align:right;font-size:13px;white-space:nowrap;">${fmt(r.importe)}</td></tr>`).join('');
  const datosPago = cfg['CLABE'] ? `<p style="margin:16px 0 4px;font-size:14px;"><strong>Para pagar por transferencia:</strong><br>Banco: ${esc(cfg['BANCO'])}<br>CLABE: <strong>${esc(cfg['CLABE'])}</strong><br>Titular: ${esc(cfg['TITULAR'])}<br>Concepto / referencia: <strong>${esc(c.id)}</strong></p>` : '';
  return `<div style="font-family:Arial,sans-serif;background:#F4F7FB;padding:24px;">
  <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:16px;padding:28px;">
    <h2 style="margin:0 0 4px;color:#0F2A52;">Tu estado de cuenta de ${esc(nombrePeriodo(c.periodo))}</h2>
    <p style="margin:0 0 16px;color:#5B6B82;font-size:14px;">Hola, ${esc(c.medico)}. Este es el resumen de los servicios de OGMpay del mes.</p>
    <table style="width:100%;border-collapse:collapse;">${filas}
      <tr><td style="padding:8px;text-align:right;font-size:13px;color:#5B6B82;">Subtotal</td><td style="padding:8px;text-align:right;font-size:13px;">${fmt(c.subtotal)}</td></tr>
      <tr><td style="padding:8px;text-align:right;font-size:13px;color:#5B6B82;">IVA</td><td style="padding:8px;text-align:right;font-size:13px;">${fmt(c.iva)}</td></tr>
      <tr><td style="padding:8px;text-align:right;font-size:15px;"><strong>Total a pagar</strong></td><td style="padding:8px;text-align:right;font-size:15px;"><strong>${fmt(c.total)}</strong></td></tr>
    </table>
    ${datosPago}
    <p style="margin:12px 0;font-size:14px;">También puedes pagar con tarjeta o en OXXO desde tu portal, en la pestaña <strong>Estado de cuenta</strong>.</p>
    <p style="margin:20px 0;"><a href="${auth.PORTAL_URL}" style="background:#10B981;color:#fff;text-decoration:none;padding:12px 20px;border-radius:10px;font-weight:bold;">Ver y pagar en mi portal</a></p>
    <p style="margin:0;font-size:12px;color:#8FA3C0;">Plazo para pagar: ${esc(cfg['DIAS PARA PAGAR'] || 5)} días. Al recibir tu pago te enviamos la factura.</p>
  </div></div>`;
}
async function enviarCorreoEstadoCuenta(c, cfg) {
  const meds = await medicosParaCobro();
  const med = meds.find(m => m.rfc === c.rfc);
  if (!med || !med.correo) throw new Error('sin correo');
  await auth.enviarCorreo(med.correo, `Tu estado de cuenta OGMpay · ${nombrePeriodo(c.periodo)}`, htmlEstadoCuenta(c, cfg));
  c.envio = auth.ahoraMx();
  await escribirCelda(`${HOJA_COB}!${columnLetter(IC.envio + 1)}${c.fila}`, c.envio);
}

async function buscarCobro(id) {
  const data = await leerCobranza();
  const i = data.findIndex((r, k) => k > 0 && celda(r, IC.id) === String(id || ''));
  return i > 0 ? filaACobro(data[i], i) : null;
}

// ── Sistema ──
async function wa_getCobranzaAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const cfg = await leerConfig();
  const data = await leerCobranza();
  const cobros = data.slice(1).map((r, i) => filaACobro(r, i + 1)).filter(c => c.id).reverse();
  const meds = await medicosParaCobro();
  return {
    ok: true, cobros, config: cfg,
    mercadoPago: Boolean(process.env.MP_ACCESS_TOKEN),
    medicos: meds.map(m => ({ rfc: m.rfc, nombre: m.nombre, plan: m.plan, gestoria: m.gestoria, correo: m.correo }))
  };
}
async function wa_guardarConfigCobranza(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const v = p.valores || {};
  if (v['CLABE'] && !/^\d{18}$/.test(String(v['CLABE']).replace(/\s/g, ''))) return { ok: false, error: 'La CLABE debe tener 18 dígitos.' };
  if (v['CLABE']) v['CLABE'] = String(v['CLABE']).replace(/\s/g, '');
  if (v['OGM RFC'] && !RE_RFC.test(String(v['OGM RFC']).toUpperCase())) return { ok: false, error: 'El RFC de OGMpay no es válido.' };
  if (v['OGM RFC']) v['OGM RFC'] = String(v['OGM RFC']).toUpperCase();
  if (v['OGM NOMBRE']) v['OGM NOMBRE'] = String(v['OGM NOMBRE']).toUpperCase().replace(/\s+/g, ' ').trim();
  if (v['COBRO DESDE'] && !/^\d{4}-\d{2}-\d{2}$/.test(v['COBRO DESDE'])) return { ok: false, error: 'Fecha "cobrar desde" no válida.' };
  await guardarConfig(v);
  let csd = null;
  if (p.csd && p.csd.cer && p.csd.key && p.csd.password) {
    const cfg = await leerConfig();
    if (!cfg['OGM RFC']) return { ok: false, error: 'Primero guarda el RFC de OGMpay.' };
    csd = await subirCsdFacturama(cfg['OGM RFC'], String(p.csd.cer), String(p.csd.key), String(p.csd.password));
    if (!csd.ok) return csd;
    await guardarConfig({ 'FACTURAR AL PAGAR': cfg['FACTURAR AL PAGAR'] || 'SI' });
  }
  return { ok: true, csd: csd ? csd.accion : null };
}
async function wa_marcarPagadoCobranza(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const c = await buscarCobro(p.id);
  if (!c) return { ok: false, error: 'No encontré ese estado de cuenta.' };
  if (p.deshacer) {
    if (c.uuid) return { ok: false, error: 'Ya tiene factura emitida; cancélala primero en Facturama.' };
    Object.assign(c, { estado: 'PENDIENTE', fechaPago: '', metodo: '', refPago: '' });
    await guardarCobro(c);
    return { ok: true };
  }
  const fecha = fechaIsoCob(p.fecha) || hoyIsoMx();
  Object.assign(c, { estado: 'PAGADO', fechaPago: fecha, metodo: String(p.metodo || 'TRANSFERENCIA').toUpperCase(), refPago: String(p.referencia || c.refPago || '') });
  await guardarCobro(c);
  let factura = null;
  if (p.facturar) factura = await emitirFacturaCobro(c);
  return { ok: true, factura };
}
async function wa_facturarCobranza(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const c = await buscarCobro(p.id);
  if (!c) return { ok: false, error: 'No encontré ese estado de cuenta.' };
  if (c.estado !== 'PAGADO') return { ok: false, error: 'Primero márcalo como pagado.' };
  return await emitirFacturaCobro(c);
}
async function wa_enviarEstadoCuentaAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  const c = await buscarCobro(p.id);
  if (!c) return { ok: false, error: 'No encontré ese estado de cuenta.' };
  try { await enviarCorreoEstadoCuenta(c, await leerConfig()); } catch (e) { return { ok: false, error: e.message === 'sin correo' ? 'Ese médico no tiene correo.' : e.message }; }
  return { ok: true };
}
async function wa_archivoFacturaCobroAdmin(p) {
  if (!(await verificarAdmin(p))) return { ok: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' };
  return await archivosFacturaCobro(await buscarCobro(p.id), p.tipo);
}

// ── Factura de OGMpay al médico (PUE, IVA trasladado) ──
const FORMA_PAGO_METODO = { 'TRANSFERENCIA': '03', 'TARJETA DE CREDITO': '04', 'TARJETA DE DEBITO': '28', 'EFECTIVO': '01', 'OXXO': '01', 'MERCADO PAGO': '06' };
function formaPagoDe(metodo) {
  const m = String(metodo || '').toUpperCase();
  if (/CREDIT/.test(m)) return '04';
  if (/DEBIT|DÉBIT/.test(m)) return '28';
  if (/OXXO|TICKET|EFECTIVO/.test(m)) return '01';
  if (/ACCOUNT_MONEY|SALDO/.test(m)) return '06';
  return FORMA_PAGO_METODO[m] || '03';
}
async function emitirFacturaCobro(c) {
  if (c.uuid) return { ok: true, uuid: c.uuid, yaExistia: true };
  const cfg = await leerConfig();
  if (!cfg['OGM RFC'] || !cfg['OGM NOMBRE'] || !/^\d{5}$/.test(cfg['OGM CP'] || '') || !cfg['OGM REGIMEN']) {
    return { ok: false, error: 'Faltan los datos fiscales de OGMpay (RFC, nombre, CP y régimen) en la configuración de cobranza.' };
  }
  const meds = await medicosParaCobro();
  const med = meds.find(m => m.rfc === c.rfc);
  if (!med) return { ok: false, error: 'El médico ya no está en EMISORES.' };
  const ivaPct = num(cfg['IVA %'], 16) / 100;
  const regEmisor = (String(cfg['OGM REGIMEN']).match(/\d{3}/) || ['626'])[0];
  const receptorMoral = med.rfc.length === 12;
  // Un concepto por tipo de cargo
  const grupos = {};
  for (const r of c.detalle) (grupos[r.tipo] = grupos[r.tipo] || []).push(r);
  const textos = {
    COMISION: n => `Comisión mercantil por gestión, trámite y cobro de facturas de honorarios médicos ante aseguradoras (${cfg['COMISION %']}% sobre el monto facturado) · ${n} factura${n === 1 ? '' : 's'} · ${nombrePeriodo(c.periodo)}`,
    GESTORIA: n => `Gestoría de cartas de honorarios · ${n} carta${n === 1 ? '' : 's'} conseguida${n === 1 ? '' : 's'} · ${nombrePeriodo(c.periodo)}`,
    INTEGRAL: () => `Plan Integral (asesoría y servicios administrativos) · ${nombrePeriodo(c.periodo)}`
  };
  const items = Object.entries(grupos).map(([tipo, rs]) => {
    const base = r2(rs.reduce((s, r) => s + r.importe, 0));
    const iva = r2(base * ivaPct);
    const taxes = [{ Name: 'IVA', Rate: ivaPct, Total: iva, Base: base, IsRetention: false }];
    let total = base + iva;
    if (receptorMoral && regEmisor === '626') { // persona moral retiene 1.25% de ISR al RESICO
      const ret = r2(base * 0.0125);
      taxes.push({ Name: 'ISR', Rate: 0.0125, Total: ret, Base: base, IsRetention: true });
      total -= ret;
    }
    return {
      ProductCode: cfg['CLAVE ' + tipo] || cfg['CLAVE PROD SERV'] || '84111506', Description: (textos[tipo] || (() => tipo))(rs.length),
      UnitCode: cfg['CLAVE UNIDAD'] || 'E48', Unit: 'Servicio', UnitPrice: base, Quantity: 1, Subtotal: base,
      TaxObject: '02', Taxes: taxes, Total: r2(total)
    };
  });
  const factura = {
    Serie: String(cfg['OGM SERIE'] || 'OGM').trim(), Folio: String(c.fila),
    CfdiType: 'I', PaymentForm: formaPagoDe(c.metodo), PaymentMethod: 'PUE',
    ExpeditionPlace: String(cfg['OGM CP']).padStart(5, '0'),
    Issuer: { Rfc: cfg['OGM RFC'], Name: cfg['OGM NOMBRE'], FiscalRegime: regEmisor },
    Receiver: { Rfc: med.rfc, Name: med.nombre, CfdiUse: 'G03', FiscalRegime: (String(med.regimen).match(/\d{3}/) || ['612'])[0], TaxZipCode: String(med.cp).padStart(5, '0') },
    Items: items
  };
  const res = await timbrar(factura);
  if (!res.uuid) {
    let msg = res.error || 'Error desconocido';
    try { const j = JSON.parse(msg); msg = j.Message || (j.ModelState ? JSON.stringify(j.ModelState) : msg); } catch (e) {}
    c.notas = 'Error al facturar: ' + String(msg).slice(0, 300);
    await guardarCobro(c);
    return { ok: false, error: 'Facturama no timbró la factura: ' + String(msg).slice(0, 300) };
  }
  c.uuid = res.uuid; c.idFacturama = res.id; c.notas = '';
  await guardarCobro(c);
  // Se le manda al médico con el PDF y XML
  try {
    const arch = await archivosFacturaCobro(c);
    if (arch.ok && med.correo && process.env.RESEND_API_KEY) {
      const fetch = require('node-fetch');
      await fetch('https://api.resend.com/emails', {
        method: 'POST',
        headers: { 'Authorization': `Bearer ${process.env.RESEND_API_KEY}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          from: 'OGMpay <notificaciones@ogmpay.com>', to: [med.correo], reply_to: 'facturacion@ogmpay.com',
          subject: `Tu factura OGMpay · ${nombrePeriodo(c.periodo)}`,
          html: `<div style="font-family:Arial,sans-serif;font-size:14px;color:#0F2A52;"><p>Hola, ${med.nombre}:</p><p>Gracias por tu pago. Te enviamos la factura de los servicios de OGMpay de ${nombrePeriodo(c.periodo)} (PDF y XML).</p><p>Total: $${c.total.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</p></div>`,
          attachments: [
            arch.pdf ? { filename: arch.nombre + '.pdf', content: arch.pdf } : null,
            arch.xml ? { filename: arch.nombre + '.xml', content: arch.xml } : null
          ].filter(Boolean)
        })
      });
    }
  } catch (e) { console.error('No se pudo mandar la factura por correo:', e.message); }
  return { ok: true, uuid: res.uuid };
}
async function archivosFacturaCobro(c, tipo) {
  if (!c) return { ok: false, error: 'No encontré ese estado de cuenta.' };
  if (!c.idFacturama) return { ok: false, error: 'Ese estado de cuenta aún no tiene factura.' };
  const nombre = `OGMpay_${c.id}`;
  const quiere = t => !tipo || tipo === t;
  const pdf = quiere('pdf') ? await descargarArchivo(c.idFacturama, 'pdf') : null;
  const xml = quiere('xml') ? await descargarArchivo(c.idFacturama, 'xml') : null;
  return { ok: true, nombre, pdf: pdf ? pdf.toString('base64') : null, xml: xml ? xml.toString('base64') : null };
}

// ── Portal del médico ──
async function wa_getEstadoCuentaMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { ok: false, error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const cfg = await leerConfig();
  const data = await leerCobranza();
  const cobros = data.slice(1).map((r, i) => filaACobro(r, i + 1)).filter(c => c.id && c.rfc === login.rfc && c.estado !== 'CANCELADO').reverse()
    .map(c => ({ id: c.id, periodo: c.periodo, subtotal: c.subtotal, iva: c.iva, total: c.total, estado: c.estado, reportado: c.reportado,
      fechaPago: c.fechaPago, metodo: c.metodo, facturada: Boolean(c.uuid), detalle: c.detalle }));
  const pago = {};
  CONFIG_PUBLICA.forEach(k => { pago[k] = cfg[k]; });
  return { ok: true, cobros, pago, mercadoPago: Boolean(process.env.MP_ACCESS_TOKEN) };
}
async function cobroDelMedico(p) {
  const login = await wa_loginMedico(p.rfc, p.contrasena);
  if (!login.ok) return { error: 'Sesión no válida. Vuelve a iniciar sesión.' };
  const c = await buscarCobro(p.id);
  if (!c || c.rfc !== login.rfc) return { error: 'No encontré ese estado de cuenta.' };
  return { c, login };
}
async function wa_reportarPagoMedico(p) {
  const { c, login, error } = await cobroDelMedico(p);
  if (error) return { ok: false, error };
  if (c.estado === 'PAGADO') return { ok: true, yaPagado: true };
  const fecha = fechaIsoCob(p.fecha) || hoyIsoMx();
  c.reportado = `${auth.ahoraMx()} · transferencia del ${fecha.split('-').reverse().join('/')}${p.nota ? ' · ' + String(p.nota).slice(0, 200) : ''}`;
  c.estado = 'REPORTADO';
  await guardarCobro(c);
  try {
    await auth.enviarCorreo('facturacion@ogmpay.com', `💰 Pago reportado: ${login.nombre} · ${c.id}`,
      `<div style="font-family:Arial,sans-serif;font-size:14px;"><p><strong>${login.nombre}</strong> (${login.rfc}) reporta que ya transfirió su estado de cuenta <strong>${c.id}</strong> por <strong>$${c.total.toLocaleString('es-MX', { minimumFractionDigits: 2 })}</strong>.</p><p>${c.reportado}</p><p>Revisa tu banco y márcalo como pagado en el sistema → Cobranza.</p></div>`);
  } catch (e) { console.error('Aviso de pago reportado:', e.message); }
  return { ok: true };
}
async function wa_archivoFacturaCobroMedico(p) {
  const { c, error } = await cobroDelMedico(p);
  if (error) return { ok: false, error };
  return await archivosFacturaCobro(c, p.tipo);
}

// ── Mercado Pago (tarjeta y OXXO). Requiere MP_ACCESS_TOKEN en Railway. ──
async function mpApi(metodo, ruta, cuerpo) {
  const fetch = require('node-fetch');
  const res = await fetch('https://api.mercadopago.com' + ruta, {
    method: metodo,
    headers: { 'Authorization': `Bearer ${process.env.MP_ACCESS_TOKEN}`, 'Content-Type': 'application/json' },
    body: cuerpo ? JSON.stringify(cuerpo) : undefined
  });
  const txt = await res.text();
  let data = null; try { data = JSON.parse(txt); } catch (e) {}
  return { ok: res.status >= 200 && res.status < 300, status: res.status, data, raw: txt };
}
async function wa_crearPagoMercadoPago(p) {
  if (!process.env.MP_ACCESS_TOKEN) return { ok: false, error: 'El pago con tarjeta aún no está disponible. Usa transferencia.' };
  const { c, login, error } = await cobroDelMedico(p);
  if (error) return { ok: false, error };
  if (c.estado === 'PAGADO') return { ok: false, error: 'Este estado de cuenta ya está pagado.' };
  const meds = await medicosParaCobro();
  const med = meds.find(m => m.rfc === login.rfc) || {};
  const api = process.env.API_PUBLIC_URL || 'https://api.ogmpay.com';
  const r = await mpApi('POST', '/checkout/preferences', {
    items: [{ id: c.id, title: `OGMpay · Estado de cuenta ${nombrePeriodo(c.periodo)}`, quantity: 1, unit_price: c.total, currency_id: 'MXN' }],
    external_reference: c.id,
    payer: med.correo ? { email: med.correo } : undefined,
    back_urls: { success: auth.PORTAL_URL + '?pago=ok&ec=' + encodeURIComponent(c.id), failure: auth.PORTAL_URL + '?pago=error', pending: auth.PORTAL_URL + '?pago=pendiente&ec=' + encodeURIComponent(c.id) },
    auto_return: 'approved',
    notification_url: api + '/api/mp/webhook',
    statement_descriptor: 'OGMPAY'
  });
  if (!r.ok || !r.data || !r.data.init_point) return { ok: false, error: 'Mercado Pago no respondió. Intenta de nuevo o usa transferencia.' };
  return { ok: true, url: r.data.init_point };
}
// Aplica un pago de Mercado Pago (lo consulta directo a MP: no se confía en lo que manda el navegador)
async function aplicarPagoMercadoPago(pago) {
  if (!pago || pago.status !== 'approved' || !pago.external_reference) return { ok: false };
  const c = await buscarCobro(pago.external_reference);
  if (!c) return { ok: false };
  if (c.estado === 'PAGADO') return { ok: true, yaPagado: true };
  if (Number(pago.transaction_amount) + 0.5 < c.total) {
    c.notas = `Pago de Mercado Pago por $${pago.transaction_amount} menor al total (${pago.id})`;
    await guardarCobro(c);
    return { ok: false };
  }
  const tipo = pago.payment_type_id || '';
  const metodo = tipo === 'credit_card' ? 'TARJETA DE CREDITO' : tipo === 'debit_card' ? 'TARJETA DE DEBITO' : tipo === 'ticket' ? 'OXXO' : tipo === 'account_money' ? 'MERCADO PAGO SALDO' : 'MERCADO PAGO';
  Object.assign(c, { estado: 'PAGADO', fechaPago: fechaIsoCob(pago.date_approved) || hoyIsoMx(), metodo, refPago: 'MP ' + pago.id });
  await guardarCobro(c);
  const cfg = await leerConfig();
  if (String(cfg['FACTURAR AL PAGAR']).toUpperCase().startsWith('S')) {
    try { await emitirFacturaCobro(c); } catch (e) { console.error('Factura automática:', e.message); }
  }
  try { await auth.enviarCorreo('facturacion@ogmpay.com', `✅ Pago recibido (Mercado Pago): ${c.medico} · ${c.id}`, `<p>${c.medico} pagó <strong>$${c.total}</strong> con ${metodo}. Referencia MP ${pago.id}.</p>`); } catch (e) {}
  return { ok: true };
}
async function wa_verificarPagoMercadoPago(p) {
  if (!process.env.MP_ACCESS_TOKEN) return { ok: false };
  const { c, error } = await cobroDelMedico(p);
  if (error) return { ok: false, error };
  if (c.estado === 'PAGADO') return { ok: true, pagado: true };
  const r = await mpApi('GET', '/v1/payments/search?sort=date_created&criteria=desc&external_reference=' + encodeURIComponent(c.id));
  const aprobado = r.ok && r.data && (r.data.results || []).find(x => x.status === 'approved');
  if (aprobado) { await aplicarPagoMercadoPago(aprobado); return { ok: true, pagado: true }; }
  const pendiente = r.ok && r.data && (r.data.results || []).find(x => x.status === 'pending' || x.status === 'in_process');
  return { ok: true, pagado: false, pendiente: Boolean(pendiente) };
}
router.post('/mp/webhook', async (req, res) => {
  res.status(200).json({ ok: true }); // MP solo necesita el 200; el trabajo sigue después
  try {
    if (!process.env.MP_ACCESS_TOKEN) return;
    const tipo = (req.body && (req.body.type || req.body.topic)) || req.query.type || req.query.topic;
    const id = (req.body && req.body.data && req.body.data.id) || req.query['data.id'] || req.query.id;
    if (!id || (tipo && !/payment/.test(tipo))) return;
    const r = await mpApi('GET', '/v1/payments/' + encodeURIComponent(id));
    if (r.ok) await aplicarPagoMercadoPago(r.data);
  } catch (e) { console.error('Webhook Mercado Pago:', e.message); }
});

// ─── Dispatcher principal ─────────────────────────────────────
const ACCIONES_ADMIN = new Set(['getEmisores', 'getAseguradoras', 'getFacturas', 'getTimbres', 'timbrarFila', 'timbrarNueva',
  'timbrarComplemento', 'guardarDatos', 'cancelarFactura', 'cancelarComplemento', 'enviarAccesoPortal', 'getEstadoPortalMedicos',
  'getArchivosFactura', 'getLinkDrive']);
const ACCIONES_FRENO = new Set(['login', 'loginMedico', 'registrarCuentaGratis', 'recuperarPasswordMedico', 'capturaPaciente', 'capturaInfo']);
const frenos = new Map();
function frenoIp(req, accion) {
  const ip = String(req.headers['x-forwarded-for'] || req.socket?.remoteAddress || '').split(',')[0].trim() || 'sin-ip';
  const limite = accion === 'capturaInfo' ? 120 : 40; // por IP cada 10 minutos
  const k = accion + '|' + ip, ahora = Date.now();
  let r = frenos.get(k);
  if (!r || ahora - r.desde > 10 * 60 * 1000) { r = { n: 0, desde: ahora }; frenos.set(k, r); }
  r.n++;
  if (frenos.size > 5000) for (const [kk, v] of frenos) if (ahora - v.desde > 10 * 60 * 1000) frenos.delete(kk);
  return r.n > limite;
}
router.post('/accion', async (req, res) => {
const body = req.body || {};
const accion = body.accion;
const params = body.params || {};
  try {
    let result;
    // Acciones del sistema interno: exigen usuario y contraseña del sistema
    if (ACCIONES_ADMIN.has(accion) && !(await verificarAdmin(params))) {
      return res.status(401).json({ ok: false, sesion: false, error: 'Sesión del sistema no válida. Vuelve a iniciar sesión.' });
    }
    // Freno básico por IP en acciones públicas (inicio de sesión, registro, recuperación)
    if (ACCIONES_FRENO.has(accion) && frenoIp(req, accion)) {
      return res.status(429).json({ ok: false, error: 'Demasiadas solicitudes. Espera unos minutos e intenta de nuevo.' });
    }
    switch (accion) {
      case 'login':                  result = await wa_login(params.usuario, params.contrasena); break;
      case 'getEmisores':            result = await wa_getEmisores(); break;
      case 'getAseguradoras':        result = await wa_getAseguradoras(); break;
      case 'getFacturas':            result = await wa_getFacturas(); break;
      case 'getTimbres':             result = await wa_getTimbres(); break;
      case 'timbrarFila':            result = await wa_timbrarFila(params.fila); break;
      case 'timbrarNueva':           result = await wa_timbrarNueva(params); break;
     case 'timbrarComplemento':     result = await wa_timbrarComplemento(params.idPago, params.fechaPago, params.filaEspecifica); break;
      case 'guardarDatos':           result = await wa_guardarDatos(params); break;
      case 'cancelarFactura':        result = await wa_cancelarFactura(params.fila, params.motive); break;
      case 'cancelarComplemento':    result = await wa_cancelarComplemento(params); break;
      case 'loginMedico':            result = await wa_loginMedico(params.rfc, params.contrasena); break;
      case 'cambiarPasswordMedico':  result = await wa_cambiarPasswordMedico(params.rfc, params.passwordActual, params.passwordNueva); break;
      case 'getFacturasMedico':      result = await wa_getFacturasMedico(params.rfc, params.contrasena); break;
      case 'enviarAccesoPortal':     result = await wa_enviarAccesoPortal(params.rfc); break;
      case 'altaClienteAdmin':       result = await wa_altaClienteAdmin(params); break;
      case 'editarClienteAdmin':     result = await wa_editarClienteAdmin(params); break;
      case 'getCobranzaAdmin':       result = await wa_getCobranzaAdmin(params); break;
      case 'generarCobranza':        result = await wa_generarCobranza(params); break;
      case 'guardarConfigCobranza':  result = await wa_guardarConfigCobranza(params); break;
      case 'marcarPagadoCobranza':   result = await wa_marcarPagadoCobranza(params); break;
      case 'facturarCobranza':       result = await wa_facturarCobranza(params); break;
      case 'enviarEstadoCuentaAdmin': result = await wa_enviarEstadoCuentaAdmin(params); break;
      case 'archivoFacturaCobroAdmin': result = await wa_archivoFacturaCobroAdmin(params); break;
      case 'getEstadoCuentaMedico':  result = await wa_getEstadoCuentaMedico(params); break;
      case 'reportarPagoMedico':     result = await wa_reportarPagoMedico(params); break;
      case 'archivoFacturaCobroMedico': result = await wa_archivoFacturaCobroMedico(params); break;
      case 'crearPagoMercadoPago':   result = await wa_crearPagoMercadoPago(params); break;
      case 'verificarPagoMercadoPago': result = await wa_verificarPagoMercadoPago(params); break;
      case 'recuperarPasswordMedico': result = await wa_recuperarPasswordMedico(params.rfc); break;
      case 'getEstadoPortalMedicos': result = await wa_getEstadoPortalMedicos(); break;
      case 'enviarAclaracion':       result = await wa_enviarAclaracion(params); break;
      case 'getArchivosFactura':     result = await wa_getArchivosFactura(params.fila, params.tipo); break;
      case 'getLinkDrive':           result = await wa_getLinkDrive(params.fila); break;
      case 'getReporteMensualMedico': result = await wa_getReporteMensualMedico(params); break;
      case 'getArchivoCfdiMedico':   result = await wa_getArchivoCfdiMedico(params); break;
      case 'getReporteMensualAdmin': result = await wa_getReporteMensualAdmin(params); break;
      case 'getArchivoCfdiAdmin':    result = await wa_getArchivoCfdiAdmin(params); break;
      case 'capturaInfo':            result = await wa_capturaInfo(params); break;
      case 'capturaPaciente':        result = await wa_capturaPaciente(params); break;
      case 'getPacientesMedico':     result = await wa_getPacientesMedico(params); break;
      case 'registrarPacienteMedico': result = await wa_registrarPacienteMedico(params); break;
      case 'getLigaCapturaMedico':   result = await wa_getLigaCapturaMedico(params); break;
      case 'getSolicitudesAdmin':    result = await wa_getSolicitudesAdmin(params); break;
      case 'actualizarSolicitud':    result = await wa_actualizarSolicitud(params); break;
      case 'registrarPacienteAdmin': result = await wa_registrarPacienteAdmin(params); break;
      case 'getLigaCapturaAdmin':    result = await wa_getLigaCapturaAdmin(params); break;
      case 'registrarCuentaGratis':  result = await wa_registrarCuentaGratis(params); break;
      case 'actualizarPacienteMedico': result = await wa_actualizarPacienteMedico(params); break;
      case 'pedirSeguimientoPaciente': result = await wa_pedirSeguimientoPaciente(params); break;
      case 'getCuentasGratisAdmin':  result = await wa_getCuentasGratisAdmin(params); break;
      case 'editarPacienteMedico':   result = await wa_editarPacienteMedico(params); break;
      case 'editarSolicitudAdmin':   result = await wa_editarSolicitudAdmin(params); break;
      case 'importarPacientesMedico': result = await wa_importarPacientesMedico(params); break;
      case 'getTiemposPago':         result = await wa_getTiemposPago(params); break;
      case 'getInfoGestoria':        result = await wa_getInfoGestoria(); break;
      case 'contratarGestoria':      result = await wa_contratarGestoria(params); break;
      case 'getCartaAutorizacionMedico': result = await wa_getCartaAutorizacionMedico(params); break;
      case 'getCartaAutorizacionAdmin':  result = await wa_getCartaAutorizacionAdmin(params); break;
      default: result = { error: 'Acción no reconocida: ' + accion };
    }
    res.json(result);
  } catch (e) {
    console.error(`Error en acción ${accion}:`, e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

router.get('/ping', (req, res) => res.json({ ok: true, service: 'OGMpay API' }));

module.exports = router;
