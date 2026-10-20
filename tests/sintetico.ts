/** Genera un reporte del reloj sintético (mismo formato que el real) para pruebas sin datos privados. */
import * as XLSX from 'xlsx';
import { daysBetween } from '../src/lib/time';

/**
 * `dias`: marcadas por día. La clave puede ser el número de día (si el período es de un solo mes)
 * o la fecha completa "AAAA-MM-DD".
 */
export function reporteSintetico(personas: { id: string; nombre: string; dias: Record<string | number, string> }[], inicio = '2026-09-01', fin = '2026-09-12'): Uint8Array {
  const fechas = daysBetween(inicio, fin);
  const nums = fechas.map((f) => Number(f.slice(8)));
  const filas: unknown[][] = [
    ['Reporte de Eventos de Asistencia'],
    [],
    ['Periodo:', '', `${inicio} ~ ${fin}`, '', '', 'Fecha actual:', fin],
    nums,
  ];
  for (const p of personas) {
    filas.push(['ID:', '', p.id, '', 'Nombre:', '', p.nombre, '', 'Departamento:', '', 'Empresa']);
    filas.push(fechas.map((f, i) => p.dias[f] ?? p.dias[nums[i]] ?? ''));
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['x']]), 'Reporte de Turnos');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(filas), 'Reporte de Asistencia');
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'biff8' }));
}
