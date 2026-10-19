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
import { compararConPagina, min, num } from './comparar';
import { hayLibreOffice, recalcular } from './libreoffice';
import { reporteSintetico } from './sintetico';

describe.skipIf(!hayLibreOffice)('Excel generado (datos sintéticos)', () => {
  const rep = leerReporte(reporteSintetico([
    // Ana: extras el 1 (0), 2 (65), 3 (sin salida), 4 (20 + tardanza 71), 5 sáb (217), dom 6 (240), 7 (una marcada), 8 (17:52), 10 (21:30 → 2h + 1:30)
    { id: '10', nombre: 'Ana Uno', dias: { 1: '08:2918:14', 2: '08:4019:05', 3: '08:1013:0113:45', 4: '09:4118:20', 5: '08:0415:37', 6: '09:0013:00', 7: '08:30', 8: '08:3017:52', 10: '08:2521:30' } },
    { id: '11', nombre: 'Beto Dos', dias: { 1: '06:3718:01', 2: '08:2918:2918:30', 3: '07:5021:00', 9: '08:2018:00' } },
  ]));
  const { config: cfg } = sincronizarPersonas(rep, configInicial(2026));
  cfg.personas[0].reparto = true;
  cfg.personas[0].valorHora = 7.5;
  cfg.personas[1].fechaIngreso = '2026-09-02';
  cfg.personas[1].valorHora = 10;
  const mes = construirMes(rep, cfg, '2026-09-12');
  const dir = mkdtempSync(join(tmpdir(), 'qalen-sint-'));

  it('Resumen y cada día coinciden con la página (recalculado sin valores guardados)', async () => {
    const f = join(dir, 'a.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep, { sinResultados: true }))));
    const x = XLSX.read(readFileSync(recalcular(f)));
    expect(x.SheetNames).toEqual(['Leeme', 'Asistencia', 'Semanal', 'Resumen']);
    compararConPagina(x, mes);
    const a = mes.filas[0].tot;
    expect([a.extra1, a.extra2, a.enFeriado]).toEqual([65 + 20 + 120 + 120, 97 + 90, 240]);
    expect(a.monto).toBeGreaterThan(0);
  });

  it('compensadas: se descuentan desde los primeros días; si sobran, "nos debe"; todo se recalcula', async () => {
    const f = join(dir, 'b.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep))));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(f);
    const ws = wb.getWorksheet(HOJAS.asis)!;
    const col = (fecha: string) => PRIMERA_COL_DIA + mes.fechas.findIndex((x) => x.fecha === fecha);
    const set = (i: number, fila: keyof typeof FILA, fecha: string, v: ExcelJS.CellValue) => {
      ws.getCell(filaBloque(i) + FILA[fila], col(fecha)).value = v;
    };
    // Ana 7/9 (una marcada 08:30): Compensado, salió 13:00 → 300 min compensados.
    // Se descuentan: día 2 (65), día 4 (20), día 5 sáb (97 al 35% y 118 de los 120 al 25%) → quedan 2 min al 25% del sábado.
    set(0, 'motivo', '2026-09-07', 'Compensado');
    set(0, 'salida', '2026-09-07', 13 / 24);
    // Beto: 4/9 y 5/9 sin marcas → Compensado (480 + 210) y 7/9 → Falta (aparte, en contra)
    set(1, 'motivo', '2026-09-04', 'Compensado');
    set(1, 'motivo', '2026-09-05', 'Compensado');
    set(1, 'motivo', '2026-09-07', 'Falta');
    // Valor hora corregido en el Resumen
    wb.getWorksheet(HOJAS.resumen)!.getCell(RES.primeraFila + 1, RES.valorHora).value = 12;
    const g = join(dir, 'b2.xlsx');
    await wb.xlsx.writeFile(g);
    const x = XLSX.read(readFileSync(recalcular(g)));
    const res = x.Sheets[HOJAS.resumen];
    const asis = x.Sheets[HOJAS.asis];
    const r0 = RES.primeraFila;
    const a = mes.filas[0].tot;
    expect([min(res, r0, RES.compensado), min(res, r0, RES.neto2), min(res, r0, RES.neto1), min(res, r0, RES.nosDebe)])
      .toEqual([300, a.extra2 - 97, a.extra1 - (65 + 20 + 118), 0]);
    expect(min(asis, filaBloque(0) + FILA.neto1, col('2026-09-05'))).toBe(2);
    expect(min(asis, filaBloque(0) + FILA.neto2, col('2026-09-10'))).toBe(90);
    // Beto: 690 compensadas contra sus extras (todas consumidas), el resto "nos debe"; la falta va en contra
    const b = mes.filas[1].tot;
    expect([min(res, r0 + 1, RES.neto1), min(res, r0 + 1, RES.neto2), min(res, r0 + 1, RES.nosDebe)]).toEqual([0, 0, 690 - b.extra1 - b.extra2]);
    expect([min(res, r0 + 1, RES.contra), num(res, r0 + 1, RES.faltas)]).toEqual([b.contra + 480, 1]);
    expect(num(res, r0 + 1, RES.monto)).toBe(0);
    // Monto de Ana = netos × valor hora × recargo
    const n1 = min(res, r0, RES.neto1), n2 = min(res, r0, RES.neto2), fer = min(res, r0, RES.feriado);
    expect(num(res, r0, RES.monto)).toBeCloseTo(Math.round(((n1 / 60) * 7.5 * 1.25 + (n2 / 60) * 7.5 * 1.35 + (fer / 60) * 7.5 * 2) * 100) / 100, 2);
    expect(num(res, r0, RES.pendientes)).toBe(a.pendientes - 1);
  });
});
