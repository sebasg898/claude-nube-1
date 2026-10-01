/** Genera un reporte del reloj sintético (mismo formato que el real) para pruebas sin datos privados. */
import * as XLSX from 'xlsx';

export function reporteSintetico(personas: { id: string; nombre: string; dias: Record<number, string> }[], inicio = '2026-09-01', fin = '2026-09-12'): Uint8Array {
  const d0 = Number(inicio.slice(8));
  const d1 = Number(fin.slice(8));
  const nums = Array.from({ length: d1 - d0 + 1 }, (_, i) => d0 + i);
  const filas: unknown[][] = [
    ['Reporte de Eventos de Asistencia'],
    [],
    ['Periodo:', '', `${inicio} ~ ${fin}`, '', '', 'Fecha actual:', fin],
    nums,
  ];
  for (const p of personas) {
    filas.push(['ID:', '', p.id, '', 'Nombre:', '', p.nombre, '', 'Departamento:', '', 'Empresa']);
    filas.push(nums.map((n) => p.dias[n] ?? ''));
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['x']]), 'Reporte de Turnos');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(filas), 'Reporte de Asistencia');
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'biff8' }));
}
