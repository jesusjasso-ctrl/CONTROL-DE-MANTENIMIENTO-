// Convierte las filas extraídas de Google Sheets (texto separado por "|")
// en los CSV que usa el tablero "Mantenimiento correctivo 2026".
// Uso: node convertir.cjs  (lee entregadas.txt y en_taller.txt de esta carpeta)
const fs = require('fs');
const path = require('path');

const dir = __dirname;
const iso = (d) => { const [a, b, c] = d.split('/'); return `${c}-${b.padStart(2, '0')}-${a.padStart(2, '0')}`; };
const q = (s) => /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
const lines = (f) => fs.readFileSync(path.join(dir, f), 'utf8').split(/\r?\n/).filter((l) => l.trim()).map((l) => l.split('|').map((x) => x.trim()));

// entregadas.txt: tipo|marca|placa|proveedor|fecha entrada|precio sin IVA|fecha entrega|dias parada|falla
const ej = lines('entregadas.txt')
  .filter((r) => r[4] && r[6])
  .map(([tipo, marca, placa, proveedor, ent, monto, sal, dias, falla]) =>
    [tipo, marca, placa === '-' ? '' : placa, proveedor, iso(ent), (monto || '').replace(/[$,]/g, ''), iso(sal), dias, falla]);
fs.writeFileSync(path.join(dir, 'unidades_entregadas.csv'),
  ['tipo,marca,placa,proveedor,fecha_entrada,monto_sin_iva,fecha_entrega,dias_parada,falla', ...ej.map((r) => r.map(q).join(','))].join('\n') + '\n');

// en_taller.txt: tipo|marca|placa|proveedor|fecha entrada|dias parada|falla|semaforo de salida (ej. 60%)
const et = lines('en_taller.txt')
  .filter((r) => r[2] && r[4])
  .map(([tipo, marca, placa, proveedor, ent, dias, falla, avance]) =>
    [tipo, marca, placa, proveedor, iso(ent), dias, falla, String(Number((avance || '0').replace('%', '')) / 100)]);
fs.writeFileSync(path.join(dir, 'unidades_en_taller.csv'),
  ['tipo,marca,placa,proveedor,fecha_entrada,dias_parada,falla,avance_salida', ...et.map((r) => r.map(q).join(','))].join('\n') + '\n');

console.log(`entregadas: ${ej.length} filas, en taller: ${et.length} filas`);
