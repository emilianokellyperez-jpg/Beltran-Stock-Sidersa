// ═══════════════════════════════════════════════════════════════
//  Beltran S.A. — Control de Stock (Códigos de Barras 1D)
//  Google Apps Script
//
//  PARA ACTUALIZAR:
//  Extensiones → Apps Script → pegá este código completo
//  → Implementar → Administrar implementaciones → editar → Nueva versión → Implementar
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = 'PEGAR_ID_DE_TU_SPREADSHEET_AQUI';

// ── Punto de entrada POST ─────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const accion = body.accion;

    let resultado;
    if      (accion === 'generar')       resultado = generarStock(body);
    else if (accion === 'consumir')      resultado = consumirViga(body);
    else if (accion === 'stock_resumen') resultado = stockResumen();
    else throw new Error('Acción desconocida: ' + accion);

    return responder(resultado);
  } catch (err) {
    return responder({ success: false, message: err.message });
  }
}

function doGet(e) {
  return responder({ success: true, message: 'API Beltran S.A. activa.' });
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Generar Stock ─────────────────────────────────────────────
function generarStock(body) {
  const tipoViga = body.tipo_viga;
  const cantidad = parseInt(body.cantidad);

  if (!tipoViga) throw new Error('Falta el tipo de viga.');
  if (!cantidad || cantidad < 1 || cantidad > 500)
    throw new Error('La cantidad debe ser entre 1 y 500.');

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName('Stock_Real');
  if (!hoja) throw new Error('No se encontró la hoja "Stock_Real".');

  hoja.getRange('A:A').setNumberFormat('@');

  const fecha  = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });
  const codes  = [];
  const usados = new Set();

  const datosExistentes = hoja.getDataRange().getValues();
  datosExistentes.forEach(fila => {
    if (fila[0]) usados.add(String(fila[0]).toUpperCase().trim());
  });

  for (let i = 0; i < cantidad; i++) {
    let code;
    let intentos = 0;
    do {
      code = generarCodigo();
      intentos++;
      if (intentos > 200) throw new Error('No se pudo generar código único. Intentá de nuevo.');
    } while (usados.has(code.toUpperCase().trim()));

    usados.add(code.toUpperCase().trim());
    codes.push(code);

    const nuevaFila = hoja.getLastRow() + 1;
    hoja.getRange(nuevaFila, 1).setNumberFormat('@').setValue(code.trim());
    hoja.getRange(nuevaFila, 2).setValue(tipoViga);
    hoja.getRange(nuevaFila, 3).setValue('Disponible');
    hoja.getRange(nuevaFila, 4).setValue(fecha);
  }

  return { success: true, codes, message: cantidad + ' viga(s) ingresadas correctamente.' };
}

// ── Consumir Viga ─────────────────────────────────────────────
function consumirViga(body) {
  const barCode = body.qr_code;
  const nroObra = body.numero_obra || 'Sin asignar';

  if (!barCode) throw new Error('Falta el código de barras.');

  const ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaStock  = ss.getSheetByName('Stock_Real');
  const hojaConsum = ss.getSheetByName('Consumido');

  if (!hojaStock)  throw new Error('No se encontró la hoja "Stock_Real".');
  if (!hojaConsum) throw new Error('No se encontró la hoja "Consumido".');

  hojaStock.getRange('A:A').setNumberFormat('@');

  const datos = hojaStock.getDataRange().getValues();
  const codigoBuscado = String(barCode).toUpperCase().trim();

  let filaEncontrada = -1;
  let tipoViga       = '';
  let estadoActual   = '';

  for (let i = 1; i < datos.length; i++) {
    const codigoEnHoja = String(datos[i][0]).toUpperCase().trim();
    if (codigoEnHoja === codigoBuscado) {
      filaEncontrada = i + 1;
      tipoViga       = datos[i][1];
      estadoActual   = datos[i][2];
      break;
    }
  }

  if (filaEncontrada === -1) {
    return { success: false, message: 'Código no encontrado en el inventario: ' + barCode };
  }
  if (estadoActual === 'Consumido') {
    return { success: false, message: 'La viga ' + barCode + ' ya fue registrada como consumida.' };
  }

  hojaStock.getRange(filaEncontrada, 3).setValue('Consumido');

  const idMov = generarIdMovimiento();
  const fecha = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });
  hojaConsum.appendRow([idMov, fecha, barCode.trim(), tipoViga, nroObra]);

  return {
    success:       true,
    message:       'Perfil registrado correctamente.',
    tipo_viga:     tipoViga,
    id_movimiento: idMov
  };
}

// ── Stock Resumen ─────────────────────────────────────────────
// Devuelve un array con disponibles/consumidos/total por tipo de perfil
function stockResumen() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName('Stock_Real');
  if (!hoja) throw new Error('No se encontró la hoja "Stock_Real".');

  const datos = hoja.getDataRange().getValues();
  // datos[0] = encabezados

  // Acumular conteos por tipo
  const mapa = {};  // { tipo: { disponibles, consumidos } }

  for (let i = 1; i < datos.length; i++) {
    const tipo   = String(datos[i][1] || '').trim();
    const estado = String(datos[i][2] || '').trim();
    if (!tipo) continue;

    if (!mapa[tipo]) mapa[tipo] = { disponibles: 0, consumidos: 0 };

    if (estado === 'Disponible') mapa[tipo].disponibles++;
    else if (estado === 'Consumido') mapa[tipo].consumidos++;
  }

  // Convertir a array ordenado por nombre de perfil
  const resumen = Object.keys(mapa)
    .sort()
    .map(tipo => ({
      tipo,
      disponibles: mapa[tipo].disponibles,
      consumidos:  mapa[tipo].consumidos,
      total:       mapa[tipo].disponibles + mapa[tipo].consumidos
    }));

  return { success: true, resumen };
}

// ── Helpers ───────────────────────────────────────────────────
function generarCodigo() {
  const d    = new Date();
  const aa   = String(d.getFullYear()).slice(2);
  const mm   = pad(d.getMonth() + 1);
  const dd   = pad(d.getDate());
  const rand = String(Math.floor(Math.random() * 9000) + 1000);
  return 'BEL' + aa + mm + dd + rand;
}

function generarIdMovimiento() {
  const d = new Date();
  return 'MOV'
    + d.getFullYear()
    + pad(d.getMonth() + 1)
    + pad(d.getDate())
    + pad(d.getHours())
    + pad(d.getMinutes())
    + pad(d.getSeconds());
}

function pad(n) { return String(n).padStart(2, '0'); }
