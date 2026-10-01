/** Coherencia página ↔ Excel con datos sintéticos (corre en cualquier máquina con LibreOffice). */
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { configInicial } from '../src/lib/config';
import { excelABuffer, FILA, filaBloque, generarExcel, HOJAS, PRIMERA_COL_DIA, RES } from '../src/lib/excel';
import { construirMes, sincronizarPersonas } from '../src/lib/model';
import { leerReporte } from '../src/lib/parser';
import { leerSaldosDeExcel } from '../src/lib/saldos';
import { hayLibreOffice, recalcular } from './libreoffice';
import { reporteSintetico } from './sintetico';

const min = (ws: XLSX.WorkSheet, r: number, c: number) => {
  const v = ws[XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })]?.v;
  return typeof v === 'number' ? Math.round(v * 1440) : 0;
};

describe.skipIf(!hayLibreOffice)('Excel generado (datos sintéticos)', () => {
  const rep = leerReporte(reporteSintetico([
    { id: '10', nombre: 'Ana Uno', dias: { 1: '08:2918:14', 2: '08:4019:05', 3: '08:1013:0113:45', 4: '09:4118:20', 5: '08:0415:37', 6: '09:0013:00', 7: '08:30', 8: '08:3017:52' } },
    { id: '11', nombre: 'Beto Dos', dias: { 1: '06:3718:01', 2: '08:2918:2918:30', 3: '07:5021:00', 9: '08:2018:00' } },
  ]));
  const cfg0 = configInicial(2026);
  const { config: cfg } = sincronizarPersonas(rep, cfg0);
  cfg.personas[1].tempranoAntesDe = 7 * 60 + 45;
  cfg.personas[1].saldoAnterior = -100;
  cfg.personas[0].reparto = true;
  const mes = construirMes(rep, cfg, '2026-09-12');
  const dir = mkdtempSync(join(tmpdir(), 'qalen-sint-'));

  it('Resumen y cada día coinciden con la página (recalculado sin valores guardados)', async () => {
    const f = join(dir, 'a.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep, { sinResultados: true }))));
    const x = XLSX.read(readFileSync(recalcular(f)));
    const res = x.Sheets[HOJAS.resumen];
    const asis = x.Sheets[HOJAS.asis];
    mes.filas.forEach((fp, i) => {
      const r = RES.primeraFila + i;
      expect([min(res, r, RES.favor), min(res, r, RES.contra), min(res, r, RES.saldoFavor) - min(res, r, RES.saldoContra)]).toEqual([fp.tot.favor, fp.tot.contra, fp.tot.saldoFinal]);
      fp.dias.forEach((d, j) => {
        const c = PRIMERA_COL_DIA + j;
        expect([d.fecha, min(asis, filaBloque(i) + FILA.favor, c), min(asis, filaBloque(i) + FILA.contra, c), min(asis, filaBloque(i) + FILA.trabajado, c)])
          .toEqual([d.fecha, d.calc.favor, d.calc.contra, d.calc.trabajado]);
      });
    });
    // El total de Ana incluye: 3/9 sin salida (0), domingo 6/9 trabajado (+240), 4/9 tardanza 71 + extras 20
    expect(mes.filas[0].tot.enFeriado).toBe(240);
    expect(min(res, RES.primeraFila, RES.enFeriado)).toBe(240);
  });

  it('completar una celda amarilla y escribir un motivo recalcula todo; la salida en formato 12 h se marca', async () => {
    const f = join(dir, 'b.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep))));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(f);
    const ws = wb.getWorksheet(HOJAS.asis)!;
    const col = (fecha: string) => PRIMERA_COL_DIA + mes.fechas.findIndex((x) => x.fecha === fecha);
    // Ana 3/9: escribe salida 19:00 (+60). Ana 7/9 (una marcada): salida 6:14 → error de 12 h
    ws.getCell(filaBloque(0) + FILA.salida, col('2026-09-03')).value = 19 / 24;
    ws.getCell(filaBloque(0) + FILA.salida, col('2026-09-07')).value = (6 * 60 + 14) / 1440;
    // Beto 4/9 sin marcadas: Permiso (−480)
    ws.getCell(filaBloque(1) + FILA.motivo, col('2026-09-04')).value = 'Permiso';
    const g = join(dir, 'b2.xlsx');
    await wb.xlsx.writeFile(g);
    const x = XLSX.read(readFileSync(recalcular(g)));
    const res = x.Sheets[HOJAS.resumen];
    const asis = x.Sheets[HOJAS.asis];
    const r0 = RES.primeraFila;
    expect(min(res, r0, RES.favor)).toBe(mes.filas[0].tot.favor + 60);
    expect(min(res, r0 + 1, RES.contra)).toBe(mes.filas[1].tot.contra + 480);
    const estado = asis[XLSX.utils.encode_cell({ r: filaBloque(0) + FILA.estado - 1, c: col('2026-09-07') - 1 })]?.v;
    expect(estado).toBe('REVISAR HORAS');
    expect(Number(res[XLSX.utils.encode_cell({ r: r0 - 1, c: RES.alertas - 1 })]?.v)).toBe(1);
    expect(Number(res[XLSX.utils.encode_cell({ r: r0 - 1, c: RES.pendientes - 1 })]?.v)).toBe(mes.filas[0].tot.pendientes - 2);
    // Los saldos se pueden leer para el mes siguiente
    const saldos = leerSaldosDeExcel(readFileSync(recalcular(g)));
    expect(saldos.map((s) => s.alias)).toEqual(['ANA', 'BETO']);
    expect(saldos[1].minutos).toBe(mes.filas[1].tot.saldoFinal - 480);
  });
});
