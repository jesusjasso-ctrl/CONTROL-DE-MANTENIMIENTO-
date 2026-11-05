// Formulario de registro para la pestaña MTTO CORRECTIVO.
// Se publica como aplicación web: doGet sirve Formulario.html y registrar()
// agrega la unidad como una fila nueva debajo de la última registrada.
// El tablero de claude.ai abre esta app con los datos en la URL
// (?placas=...&entrada=2026-11-05...) para que solo falte confirmar.

const ID_HOJA = '1sn1TANcrOYBhzfjdqS8El2wDVhbQDoiqq9etiIHvd_g';
const PESTANA = 'MTTO CORRECTIVO';
const FILA_ENCABEZADOS = 2;
const PRIMERA_FILA = 3;

// Nombre corto de cada campo en la URL, según el encabezado de la fila 2.
const CLAVES = [
  ['tipo', /^tipo/i],
  ['marca', /marca/i],
  ['placas', /placa/i],
  ['proveedor', /proveedor/i],
  ['svc', /^svc/i],
  ['entrada', /fecha\s+entrada/i],
  ['fallas', /falla/i],
  ['semaforo', /sem[aá]n?foro/i],
  ['req', /^req/i],
  ['fecha_req', /^fecha\s+req/i],
  ['fecha_surtido', /^fecha\s+surtido\s*$/i],
  ['fecha_surtido_taller', /surtido\s+a\s+taller/i],
  ['monto', /monto/i],
  ['refacciones', /refacciones/i],
  ['observaciones', /observaciones/i],
];

function doGet(e) {
  const plantilla = HtmlService.createTemplateFromFile('Formulario');
  const datos = {};
  CLAVES.forEach(([clave]) => {
    const v = e && e.parameter && e.parameter[clave];
    if (v) datos[clave] = String(v).slice(0, 2000);
  });
  plantilla.datosIniciales = JSON.stringify(datos).replace(/</g, '\\u003c');
  return plantilla.evaluate()
    .setTitle('Registro MTTO correctivo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function clave_(titulo) {
  const par = CLAVES.find(([, patron]) => patron.test(titulo));
  return par ? par[0] : null;
}

function hoja_() {
  const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(PESTANA);
  if (!hoja) throw new Error(`No encontré la pestaña "${PESTANA}" en la hoja.`);
  return hoja;
}

// Columnas con encabezado en la fila 2 (las demás se ignoran).
function numColumnas_(hoja) {
  const encabezados = hoja.getRange(FILA_ENCABEZADOS, 1, 1, hoja.getLastColumn()).getDisplayValues()[0];
  let n = 0;
  encabezados.forEach((t, i) => { if (t.trim()) n = i + 1; });
  return n;
}

function columnaPorTitulo_(encabezados, patron) {
  return encabezados.findIndex((t) => patron.test(t));
}

// Última fila con placas capturadas, o la fila de encabezados si no hay ninguna.
function ultimaFila_(hoja, colPlacas) {
  const total = hoja.getMaxRows() - PRIMERA_FILA + 1;
  const placas = hoja.getRange(PRIMERA_FILA, colPlacas + 1, total, 1).getDisplayValues();
  for (let i = placas.length - 1; i >= 0; i--) {
    if (placas[i][0].trim()) return PRIMERA_FILA + i;
  }
  return FILA_ENCABEZADOS;
}

function opciones_(regla) {
  if (!regla) return [];
  const tipo = regla.getCriteriaType();
  const valores = regla.getCriteriaValues();
  let lista = [];
  if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) lista = valores[0];
  if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) lista = valores[0].getDisplayValues().flat();
  return [...new Set(lista.map((v) => String(v).trim()).filter(Boolean))];
}

function tipo_(titulo) {
  if (/fecha/i.test(titulo)) return 'date';
  if (/monto/i.test(titulo)) return 'number';
  if (/observaciones|refacciones/i.test(titulo)) return 'textarea';
  return 'text';
}

// La fila que sirve de modelo: la última registrada, o la primera de datos.
function plantilla_(hoja) {
  const n = numColumnas_(hoja);
  const encabezados = hoja.getRange(FILA_ENCABEZADOS, 1, 1, n).getDisplayValues()[0].map((t) => t.trim());
  const colPlacas = columnaPorTitulo_(encabezados, /placa/i);
  const colEntrada = columnaPorTitulo_(encabezados, /fecha\s+entrada/i);
  if (colPlacas < 0 || colEntrada < 0) throw new Error('No encontré las columnas Placas y Fecha Entrada Taller en la fila 2.');
  const ultima = ultimaFila_(hoja, colPlacas);
  const fila = Math.max(ultima, PRIMERA_FILA);
  const rango = hoja.getRange(fila, 1, 1, n);
  return { n, encabezados, colPlacas, colEntrada, ultima, fila, formulas: rango.getFormulas()[0], reglas: rango.getDataValidations()[0] };
}

// Campos que muestra el formulario: todas las columnas que no son fórmula.
function obtenerCampos() {
  const p = plantilla_(hoja_());
  return p.encabezados
    .map((titulo, i) => ({
      col: i,
      clave: clave_(titulo),
      titulo,
      tipo: tipo_(titulo),
      opciones: opciones_(p.reglas[i]),
      requerido: i === p.colPlacas || i === p.colEntrada,
      calculado: Boolean(p.formulas[i]),
    }))
    .filter((c) => c.titulo && !c.calculado);
}

function aFecha_(texto) {
  const [y, m, d] = texto.split('-').map(Number);
  return new Date(y, m - 1, d);
}

// datos: { "<índice de columna>": "valor" } tal como lo envía el formulario.
function registrar(datos) {
  const candado = LockService.getScriptLock();
  candado.waitLock(20000);
  try {
    const hoja = hoja_();
    const p = plantilla_(hoja);
    const placas = String(datos[p.colPlacas] || '').trim().toUpperCase();
    if (!placas) throw new Error('Escribe las placas de la unidad.');
    if (!datos[p.colEntrada]) throw new Error('Escribe la fecha de entrada a taller.');

    if (p.ultima >= PRIMERA_FILA) {
      const actuales = hoja.getRange(PRIMERA_FILA, p.colPlacas + 1, p.ultima - PRIMERA_FILA + 1, 1).getDisplayValues();
      const repetida = actuales.findIndex((r) => r[0].trim().toUpperCase() === placas);
      if (repetida >= 0) throw new Error(`La unidad ${placas} ya está registrada en taller (fila ${PRIMERA_FILA + repetida}).`);
    }

    const nueva = p.ultima + 1;
    if (nueva > hoja.getMaxRows()) hoja.insertRowAfter(hoja.getMaxRows());
    const destino = hoja.getRange(nueva, 1, 1, p.n);
    // Copia formato, listas y fórmulas de la fila modelo; luego escribe los valores.
    if (p.ultima >= PRIMERA_FILA) hoja.getRange(p.fila, 1, 1, p.n).copyTo(destino);

    const valores = p.encabezados.map((titulo, i) => {
      if (p.formulas[i]) return null;
      const v = String(datos[i] == null ? '' : datos[i]).trim();
      if (!v) return '';
      if (i === p.colPlacas) return placas;
      if (tipo_(titulo) === 'date') return aFecha_(v);
      if (tipo_(titulo) === 'number') return Number(v);
      return v;
    });
    valores.forEach((v, i) => { if (v !== null) destino.getCell(1, i + 1).setValue(v); });

    return { fila: nueva, placas };
  } finally {
    candado.releaseLock();
  }
}
