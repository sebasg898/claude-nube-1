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
export const num = (ws: XLSX.WorkSheet, r: number, c: number) => Number(celda(ws, r, c) ?? 0);

/** Compara todo el Resumen y cada día de Asistencia con lo que calculó la página. */
export function compararConPagina(x: XLSX.WorkBook, mes: Mes) {
  const res = x.Sheets[HOJAS.resumen];
  const asis = x.Sheets[HOJAS.asis];
  mes.filas.forEach((fp, i) => {
    const r = RES.primeraFila + i;
    const t = fp.tot;
    const leido = {
      p: fp.persona.alias,
      trab: min(res, r, RES.trabajado), base: min(res, r, RES.base),
      e1: min(res, r, RES.extra1), e2: min(res, r, RES.extra2), cmp: min(res, r, RES.compensado),
      n1: min(res, r, RES.neto1), n2: min(res, r, RES.neto2), fer: min(res, r, RES.feriado),
      monto: num(res, r, RES.monto), debe: min(res, r, RES.nosDebe), con: min(res, r, RES.contra),
      faltas: num(res, r, RES.faltas), pend: num(res, r, RES.pendientes),
    };
    expect(leido).toEqual({
      p: fp.persona.alias, trab: t.trabajado, base: t.base, e1: t.extra1, e2: t.extra2, cmp: t.compensado,
      n1: t.neto1, n2: t.neto2, fer: t.enFeriado, monto: t.monto, debe: t.nosDebe, con: t.contra, faltas: t.faltas, pend: t.pendientes,
    });
    const b = filaBloque(i);
    fp.dias.forEach((d, j) => {
      const c = PRIMERA_COL_DIA + j;
      const filas = ['trabajado', 'extra1', 'extra2', 'feriado', 'contra', 'compensado', 'neto1', 'neto2'] as const;
      const esperado = [d.calc.trabajado, d.calc.extra1, d.calc.extra2, d.calc.enFeriado, d.calc.contra, d.calc.compensado, d.neto.neto1, d.neto.neto2];
      expect([fp.persona.alias, d.fecha, ...filas.map((k) => min(asis, b + FILA[k], c))]).toEqual([fp.persona.alias, d.fecha, ...esperado]);
    });
  });
}
