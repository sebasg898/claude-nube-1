/** Lectura de valores del Excel recalculado y comparación con el cálculo de la página. */
import * as XLSX from 'xlsx';
import { expect } from 'vitest';
import { FILA, filaBloque, HOJAS, PRIMERA_COL_DIA, RES } from '../src/lib/excel';
import type { Mes } from '../src/lib/model';

export const celda = (ws: XLSX.WorkSheet, r: number, c: number) => ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })]?.v;
export const min = (ws: XLSX.WorkSheet, r: number, c: number) => {
  const v = celda(ws, r, c);
  return typeof v === 'number' ? Math.round(v * 1440) : 0;
};
/** Lee el texto "+h:mm" / "-h:mm" del saldo. */
export const saldo = (ws: XLSX.WorkSheet, r: number) => {
  const m = /^([+-])(\d+):(\d{2})$/.exec(String(celda(ws, r, RES.saldo)));
  if (!m) throw new Error(`Saldo ilegible: ${celda(ws, r, RES.saldo)}`);
  return (m[1] === '-' ? -1 : 1) * (Number(m[2]) * 60 + Number(m[3]));
};

/** Compara todo el Resumen y cada día de Asistencia con lo que calculó la página. */
export function compararConPagina(x: XLSX.WorkBook, mes: Mes) {
  const res = x.Sheets[HOJAS.resumen];
  const asis = x.Sheets[HOJAS.asis];
  mes.filas.forEach((fp, i) => {
    const r = RES.primeraFila + i;
    const t = fp.tot;
    expect({ p: fp.persona.alias, fav: min(res, r, RES.favor), con: min(res, r, RES.contra), cmp: min(res, r, RES.compensado), trab: min(res, r, RES.trabajado), base: min(res, r, RES.base), saldo: saldo(res, r), faltas: Number(celda(res, r, RES.faltas) ?? 0), pend: Number(celda(res, r, RES.pendientes) ?? 0) })
      .toEqual({ p: fp.persona.alias, fav: t.favor, con: t.contra, cmp: t.compensado, trab: t.trabajado, base: t.base, saldo: t.saldo, faltas: t.faltas, pend: t.pendientes });
    fp.dias.forEach((d, j) => {
      const c = PRIMERA_COL_DIA + j;
      const b = filaBloque(i);
      expect([fp.persona.alias, d.fecha, min(asis, b + FILA.favor, c), min(asis, b + FILA.contra, c), min(asis, b + FILA.compensado, c), min(asis, b + FILA.trabajado, c)])
        .toEqual([fp.persona.alias, d.fecha, d.calc.favor, d.calc.contra, d.calc.compensado, d.calc.trabajado]);
    });
  });
}

