const CONFIG = {
  SPREADSHEET_ID: '',
  SHEET_REGISTROS: 'Registros',
  SHEET_SELLOS: 'Sellos',
  SHEET_JUNTAS: 'Juntas JVPLC',
  SHEET_TRANSPORTISTAS: 'Transportistas',
  DRIVE_FOLDER_NAME: 'Sellos JVPL - Imágenes',
  VENTANA_DUPLICADOS_SEGUNDOS: 20
};

const REGISTROS_HEADERS = [
  'Marca temporal',
  'Sucursal',
  'Fecha de envío',
  'N° Junta JVPLC',
  'Dorados',
  'Morados',
  'Alícuotas',
  'Cultivos',
  'Orina en tubo (pruebas químicas)',
  'Alícuotas de suero',
  'MX (¿Se derramó?)',
  'Nombre del paciente',
  'Código MX',
  'Hora preparación',
  'Temperatura (°C)',
  'Hora contacto',
  'Hora entrega',
  'Transportista'
];

const SELLOS_HEADERS = [
  'Marca temporal',
  'Número de responsable',
  'Imagen del sello (URL)'
];

const JUNTAS_HEADERS = ['N° Junta JVPLC'];
const TRANSPORTISTAS_HEADERS = ['Transportista'];

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getRegistros') return jsonResponse(getRegistros());
  if (action === 'getJuntas') return jsonResponse(getJuntas());
  if (action === 'getTransportistas') return jsonResponse(getTransportistas());

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Formulario de registro de envíos de muestras.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'registrarEnvio';
    const data = body.data || body;

    let result;
    switch (action) {
      case 'registrarEnvio':
        result = registrarEnvio(data);
        break;
      case 'guardarSello':
        result = guardarSello(data);
        break;
      case 'getRegistros':
        result = getRegistros();
        break;
      case 'getJuntas':
        result = getJuntas();
        break;
      case 'agregarJunta':
        result = agregarJunta(data.nombre || data);
        break;
      case 'eliminarJunta':
        result = eliminarJunta(data.nombre || data);
        break;
      case 'getTransportistas':
        result = getTransportistas();
        break;
      case 'agregarTransportista':
        result = agregarTransportista(data.nombre || data);
        break;
      case 'eliminarTransportista':
        result = eliminarTransportista(data.nombre || data);
        break;
      default:
        result = { success: false, error: 'Acción no reconocida: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function registrarEnvio(data) {
  try {
    const validationError = validarEnvio(data);
    if (validationError) return { success: false, error: validationError };

    const sheet = getOrCreateSheet(CONFIG.SHEET_REGISTROS, REGISTROS_HEADERS);

    if (esRegistroDuplicado(sheet, data)) {
      return {
        success: false,
        error: 'Este envío ya se registró hace unos segundos. Evita hacer doble clic en "Registrar Envío".',
        duplicado: true
      };
    }

    sheet.appendRow([
      new Date(),
      data.sucursal || '',
      data.fechaEnvio || '',
      data.juntaJVPLC || '',
      toNumber(data.dorados),
      toNumber(data.morados),
      toNumber(data.alicuotas),
      toNumber(data.cultivos),
      toNumber(data.orinaTubo),
      toNumber(data.alicuotasSuero),
      data.mx || 'No',
      data.mx === 'Si' || data.mx === 'Sí' ? (data.mxNombrePaciente || '') : '',
      data.mx === 'Si' || data.mx === 'Sí' ? (data.mxCodigo || '') : '',
      data.horaPreparacion || '',
      data.temperatura || '',
      data.horaContacto || '',
      data.horaEntrega || '',
      data.transportista || ''
    ]);

    return { success: true, message: 'Envío registrado correctamente.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function validarEnvio(data) {
  if (!data) return 'No se recibieron datos.';
  if (!data.sucursal) return 'Falta seleccionar la Sucursal.';
  if (!data.fechaEnvio) return 'Falta la Fecha de envío.';
  if (!data.juntaJVPLC) return 'Falta seleccionar el N° Junta JVPLC.';
  if (!data.transportista) return 'Falta seleccionar el Transportista.';

  if (data.mx === 'Si' || data.mx === 'Sí') {
    if (!data.mxNombrePaciente) return 'Falta el Nombre del paciente (MX = Sí).';
    if (!data.mxCodigo) return 'Falta el Código (MX = Sí).';
  }

  return null;
}

function esRegistroDuplicado(sheet, data) {
  const ultimaFila = sheet.getLastRow();
  if (ultimaFila < 2) return false;

  const headers = REGISTROS_HEADERS;
  const valoresUltimaFila = sheet.getRange(ultimaFila, 1, 1, headers.length).getValues()[0];

  const indice = {};
  headers.forEach(function (h, i) { indice[h] = i; });

  const marcaTemporalAnterior = valoresUltimaFila[indice['Marca temporal']];
  if (!(marcaTemporalAnterior instanceof Date)) return false;

  const segundosTranscurridos = (new Date().getTime() - marcaTemporalAnterior.getTime()) / 1000;
  if (segundosTranscurridos > CONFIG.VENTANA_DUPLICADOS_SEGUNDOS) return false;

  const camposClave = [
    ['Sucursal', data.sucursal || ''],
    ['Fecha de envío', data.fechaEnvio || ''],
    ['N° Junta JVPLC', data.juntaJVPLC || ''],
    ['Dorados', toNumber(data.dorados)],
    ['Morados', toNumber(data.morados)],
    ['Alícuotas', toNumber(data.alicuotas)],
    ['Cultivos', toNumber(data.cultivos)],
    ['Orina en tubo (pruebas químicas)', toNumber(data.orinaTubo)],
    ['Alícuotas de suero', toNumber(data.alicuotasSuero)],
    ['Hora preparación', data.horaPreparacion || ''],
    ['Hora contacto', data.horaContacto || ''],
    ['Hora entrega', data.horaEntrega || ''],
    ['Transportista', data.transportista || '']
  ];

  return camposClave.every(function (par) {
    const valorAnterior = valoresUltimaFila[indice[par[0]]];
    return String(valorAnterior).trim() === String(par[1]).trim();
  });
}

function getRegistros() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_REGISTROS, REGISTROS_HEADERS);
    const data = sheet.getDataRange().getValues();

    if (data.length < 2) return { success: true, headers: REGISTROS_HEADERS, records: [] };

    const headers = data[0];
    const timeZone = Session.getScriptTimeZone();

    const records = data.slice(1).map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = serializeCellValue(row[i], timeZone); });
      return obj;
    });

    return { success: true, headers: headers, records: records };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function serializeCellValue(value, timeZone) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    const esSoloFecha = value.getHours() === 0 && value.getMinutes() === 0 && value.getSeconds() === 0;
    return Utilities.formatDate(value, timeZone, esSoloFecha ? 'dd/MM/yyyy' : 'dd/MM/yyyy HH:mm:ss');
  }
  return value;
}

function guardarSello(data) {
  try {
    if (!data || !data.numeroResponsable) {
      return { success: false, error: 'Falta el Número de responsable.' };
    }

    const sheet = getOrCreateSheet(CONFIG.SHEET_SELLOS, SELLOS_HEADERS);

    let imageUrl = '';
    if (data.imageBase64) {
      imageUrl = saveImageToDrive(data.imageBase64, data.imageMimeType || 'image/png', data.numeroResponsable);
    }

    sheet.appendRow([new Date(), data.numeroResponsable, imageUrl]);

    return { success: true, message: 'Sello guardado correctamente.', imageUrl: imageUrl };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveImageToDrive(base64Data, mimeType, label) {
  const folder = getOrCreateFolder(CONFIG.DRIVE_FOLDER_NAME);
  const cleanBase64 = base64Data.indexOf('base64,') > -1 ? base64Data.split('base64,')[1] : base64Data;
  const bytes = Utilities.base64Decode(cleanBase64);
  const fileName = 'sello_' + (label || 'sin_numero') + '_' + new Date().getTime();
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function getJuntas() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_JUNTAS, JUNTAS_HEADERS);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, juntas: [] };

    const juntas = data.slice(1).map(function (row) { return String(row[0] || '').trim(); }).filter(Boolean);
    return { success: true, juntas: juntas };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function agregarJunta(nombre) {
  try {
    const valor = String(nombre || '').trim();
    if (!valor) return { success: false, error: 'El N° Junta JVPLC no puede estar vacío.' };

    const sheet = getOrCreateSheet(CONFIG.SHEET_JUNTAS, JUNTAS_HEADERS);
    const data = sheet.getDataRange().getValues();

    const yaExiste = data.slice(1).some(function (row) {
      return String(row[0] || '').trim().toLowerCase() === valor.toLowerCase();
    });
    if (yaExiste) return { success: false, error: 'Ese N° Junta JVPLC ya está registrado.' };

    sheet.appendRow([valor]);
    return { success: true, message: 'N° Junta JVPLC agregado correctamente.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function eliminarJunta(nombre) {
  try {
    const valor = String(nombre || '').trim();
    if (!valor) return { success: false, error: 'Falta indicar qué N° Junta JVPLC eliminar.' };

    const sheet = getOrCreateSheet(CONFIG.SHEET_JUNTAS, JUNTAS_HEADERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === valor) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'N° Junta JVPLC eliminado correctamente.' };
      }
    }

    return { success: false, error: 'No se encontró ese N° Junta JVPLC en la lista.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getTransportistas() {
  try {
    const sheet = getOrCreateSheet(CONFIG.SHEET_TRANSPORTISTAS, TRANSPORTISTAS_HEADERS);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, transportistas: [] };

    const transportistas = data.slice(1).map(function (row) { return String(row[0] || '').trim(); }).filter(Boolean);
    return { success: true, transportistas: transportistas };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function agregarTransportista(nombre) {
  try {
    const valor = String(nombre || '').trim();
    if (!valor) return { success: false, error: 'El nombre del Transportista no puede estar vacío.' };

    const sheet = getOrCreateSheet(CONFIG.SHEET_TRANSPORTISTAS, TRANSPORTISTAS_HEADERS);
    const data = sheet.getDataRange().getValues();

    const yaExiste = data.slice(1).some(function (row) {
      return String(row[0] || '').trim().toLowerCase() === valor.toLowerCase();
    });
    if (yaExiste) return { success: false, error: 'Ese Transportista ya está registrado.' };

    sheet.appendRow([valor]);
    return { success: true, message: 'Transportista agregado correctamente.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function eliminarTransportista(nombre) {
  try {
    const valor = String(nombre || '').trim();
    if (!valor) return { success: false, error: 'Falta indicar qué Transportista eliminar.' };

    const sheet = getOrCreateSheet(CONFIG.SHEET_TRANSPORTISTAS, TRANSPORTISTAS_HEADERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === valor) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Transportista eliminado correctamente.' };
      }
    }

    return { success: false, error: 'No se encontró ese Transportista en la lista.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getSpreadsheet() {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No se encontró la hoja de cálculo. Si este script es independiente, pega el ID de tu Google Sheet en CONFIG.SPREADSHEET_ID.');
  }
  return ss;
}

function getOrCreateSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function toNumber(value) {
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}
