// ═══════════════════════════════════════════════════════════════
//  Beltran S.A. — Control de Stock (Códigos de Barras 1D)
//  Google Apps Script — Versión completa (Perfiles + Bobinas)
//
//  PARA ACTUALIZAR:
//  Extensiones → Apps Script → pegá este código completo
//  → Implementar → Administrar implementaciones → editar → Nueva versión → Implementar
//
//  Hojas necesarias en el Spreadsheet:
//    - Stock_Real   (QR_Code | Tipo_Viga | Estado | Fecha_Ingreso)
//    - Consumido    (ID_Movimiento | Fecha_Consumo | QR_Code | Tipo_Viga | Numero_Obra)
//    - Bobinas      (Codigo | Medida | Peso_kg | Estado | Fecha_Ingreso | Fecha_Proceso | Fecha_Finalizada)
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1xczrzJvQ8RCuGcQTwmE-Pl8OKnaBab3j_O0UwTITu1E';

// ── Punto de entrada POST ─────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const accion = body.accion;

    let resultado;
    if      (accion === 'generar')          resultado = generarStock(body);
    else if (accion === 'consumir')         resultado = consumirViga(body);
    else if (accion === 'stock_resumen')    resultado = stockResumen();
    else if (accion === 'generar_bobina')   resultado = generarBobina(body);
    else if (accion === 'escanear_bobina')  resultado = escanearBobina(body);
    else if (accion === 'resumen_bobinas')  resultado = resumenBobinas();
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

// ═══════════════════════════════════════════════════════════════
//  PERFILES — Generar Stock
// ═══════════════════════════════════════════════════════════════
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
      code = generarCodigo('BEL');
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

// ═══════════════════════════════════════════════════════════════
//  PERFILES — Consumir Viga
// ═══════════════════════════════════════════════════════════════
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

// ═══════════════════════════════════════════════════════════════
//  PERFILES — Stock Resumen
// ═══════════════════════════════════════════════════════════════
function stockResumen() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName('Stock_Real');
  if (!hoja) throw new Error('No se encontró la hoja "Stock_Real".');

  const datos = hoja.getDataRange().getValues();
  const mapa = {};

  for (let i = 1; i < datos.length; i++) {
    const tipo   = String(datos[i][1] || '').trim();
    const estado = String(datos[i][2] || '').trim();
    if (!tipo) continue;

    if (!mapa[tipo]) mapa[tipo] = { disponibles: 0, consumidos: 0 };

    if (estado === 'Disponible') mapa[tipo].disponibles++;
    else if (estado === 'Consumido') mapa[tipo].consumidos++;
  }

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

// ═══════════════════════════════════════════════════════════════
//  BOBINAS — Generar Bobina
// ═══════════════════════════════════════════════════════════════
function generarBobina(body) {
  const medida = body.medida;
  const peso   = parseFloat(body.peso);

  if (!medida)         throw new Error('Falta la medida de la bobina.');
  if (!peso || peso <= 0) throw new Error('Falta el peso de la bobina.');

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  let hoja   = ss.getSheetByName('Bobinas');

  // Crear la hoja si no existe
  if (!hoja) {
    hoja = ss.insertSheet('Bobinas');
    hoja.getRange(1, 1, 1, 7).setValues([[
      'Codigo', 'Medida', 'Peso_kg', 'Estado', 'Fecha_Ingreso', 'Fecha_Proceso', 'Fecha_Finalizada'
    ]]);
    hoja.getRange('A:A').setNumberFormat('@');
  }

  hoja.getRange('A:A').setNumberFormat('@');

  // Verificar unicidad
  const datosExistentes = hoja.getDataRange().getValues();
  const usados = new Set();
  datosExistentes.forEach(fila => {
    if (fila[0]) usados.add(String(fila[0]).toUpperCase().trim());
  });

  let codigo;
  let intentos = 0;
  do {
    codigo = generarCodigo('BOB');
    intentos++;
    if (intentos > 200) throw new Error('No se pudo generar código único para bobina.');
  } while (usados.has(codigo.toUpperCase().trim()));

  const fecha = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });
  const nuevaFila = hoja.getLastRow() + 1;
  hoja.getRange(nuevaFila, 1).setNumberFormat('@').setValue(codigo.trim());
  hoja.getRange(nuevaFila, 2).setValue(medida);
  hoja.getRange(nuevaFila, 3).setValue(peso);
  hoja.getRange(nuevaFila, 4).setValue('Disponible');
  hoja.getRange(nuevaFila, 5).setValue(fecha);
  hoja.getRange(nuevaFila, 6).setValue('');
  hoja.getRange(nuevaFila, 7).setValue('');

  return {
    success: true,
    codigo:  codigo,
    message: 'Bobina ' + codigo + ' ingresada correctamente.'
  };
}

// ═══════════════════════════════════════════════════════════════
//  BOBINAS — Escanear Bobina (ciclo: Disponible → En proceso → Finalizada)
// ═══════════════════════════════════════════════════════════════
function escanearBobina(body) {
  const codigoEscaneado = body.codigo;
  if (!codigoEscaneado) throw new Error('Falta el código de bobina.');

  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName('Bobinas');
  if (!hoja) throw new Error('No existe la hoja "Bobinas". Generá una bobina primero.');

  hoja.getRange('A:A').setNumberFormat('@');

  const datos = hoja.getDataRange().getValues();
  const buscar = String(codigoEscaneado).toUpperCase().trim();

  let filaEncontrada = -1;
  let estadoActual   = '';

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).toUpperCase().trim() === buscar) {
      filaEncontrada = i + 1;
      estadoActual   = datos[i][3];
      break;
    }
  }

  if (filaEncontrada === -1) {
    return { success: false, message: 'Código de bobina no encontrado: ' + codigoEscaneado };
  }

  const fecha = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });

  if (estadoActual === 'Disponible') {
    hoja.getRange(filaEncontrada, 4).setValue('En proceso');
    hoja.getRange(filaEncontrada, 6).setValue(fecha);
    return { success: true, message: 'Bobina ' + codigoEscaneado + ' → En proceso' };

  } else if (estadoActual === 'En proceso') {
    hoja.getRange(filaEncontrada, 4).setValue('Finalizada');
    hoja.getRange(filaEncontrada, 7).setValue(fecha);
    return { success: true, message: 'Bobina ' + codigoEscaneado + ' → Finalizada' };

  } else if (estadoActual === 'Finalizada') {
    return { success: false, message: 'La bobina ' + codigoEscaneado + ' ya está Finalizada.' };

  } else {
    return { success: false, message: 'Estado desconocido para la bobina: ' + estadoActual };
  }
}

// ═══════════════════════════════════════════════════════════════
//  BOBINAS — Resumen
// ═══════════════════════════════════════════════════════════════
function resumenBobinas() {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName('Bobinas');

  if (!hoja) return { success: true, bobinas: [] };

  const datos = hoja.getDataRange().getValues();
  const bobinas = [];

  for (let i = 1; i < datos.length; i++) {
    const codigo = String(datos[i][0] || '').trim();
    if (!codigo) continue;
    bobinas.push({
      codigo:  codigo,
      medida:  datos[i][1] || '',
      peso:    datos[i][2] || '',
      estado:  datos[i][3] || '',
      fecha:   datos[i][4] || ''
    });
  }

  // Orden: Disponible primero, luego En proceso, luego Finalizada
  const orden = { 'Disponible': 0, 'En proceso': 1, 'Finalizada': 2 };
  bobinas.sort((a, b) => (orden[a.estado] ?? 3) - (orden[b.estado] ?? 3));

  return { success: true, bobinas };
}

// ═══════════════════════════════════════════════════════════════
//  Helpers
// ═══════════════════════════════════════════════════════════════
function generarCodigo(prefijo) {
  const d    = new Date();
  const aa   = String(d.getFullYear()).slice(2);
  const mm   = pad(d.getMonth() + 1);
  const dd   = pad(d.getDate());
  const rand = String(Math.floor(Math.random() * 9000) + 1000);
  return prefijo + aa + mm + dd + rand;
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
