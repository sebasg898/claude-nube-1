import * as XLSX from 'xlsx';
import { HOJAS } from './excel';

export interface SaldoLeido {
  alias: string;
  minutos: number; // positivo = le debemos; negativo = nos debe
}

/**
 * Lee los saldos finales de un Excel ya trabajado (hoja Resumen), para usarlos como
 * saldo anterior del mes siguiente. Usa los valores que Excel guardó al grabar el archivo.
 */
export function leerSaldosDeExcel(data: ArrayBuffer | Uint8Array): SaldoLeido[] {
  const wb = XLSX.read(data, { type: 'array' });
  const ws = wb.Sheets[HOJAS.resumen];
  if (!ws) throw new Error('Ese archivo no tiene la hoja "Resumen". Elige un Excel generado por esta página.');
  const filas = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: null });
  const iCab = filas.findIndex((f) => f.some((v) => String(v ?? '').trim() === 'Nombre') && f.some((v) => String(v ?? '').startsWith('Le debemos')));
  if (iCab < 0) throw new Error('No encontré la tabla del Resumen (columnas "Nombre" y "Le debemos").');
  const cab = filas[iCab].map((v) => String(v ?? '').trim());
  const cN = cab.indexOf('Nombre');
  const cF = cab.findIndex((v) => v.startsWith('Le debemos'));
  const cC = cab.findIndex((v) => v.startsWith('Nos debe'));
  const out: SaldoLeido[] = [];
  for (let r = iCab + 1; r < filas.length; r++) {
    const alias = String(filas[r][cN] ?? '').trim();
    if (!alias || alias === 'TOTAL') break;
    const f = Number(filas[r][cF] ?? 0);
    const c = Number(filas[r][cC] ?? 0);
    if (!Number.isFinite(f) || !Number.isFinite(c)) throw new Error(`El saldo de ${alias} no es un número. Abre el archivo en Excel, guárdalo y vuelve a intentar.`);
    out.push({ alias, minutos: Math.round(f * 1440) - Math.round(c * 1440) });
  }
  if (!out.length) throw new Error('El Resumen no tiene personas.');
  return out;
}
