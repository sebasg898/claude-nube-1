/** Coherencia página ↔ Excel con datos sintéticos (corre en cualquier máquina con LibreOffice). */
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { configInicial } from '../src/lib/config';
import { excelABuffer, FILA, filaBloque, generarExcel, HOJAS, PRIMERA_COL_DIA, RES } from '../src/lib/excel';
import { construirMes, corteMes, sincronizarPersonas } from '../src/lib/model';
import { leerReporte } from '../src/lib/parser';
import { compararConPagina, min, num } from './comparar';
import { hayLibreOffice, recalcular } from './libreoffice';
import { reporteSintetico } from './sintetico';

describe.skipIf(!hayLibreOffice)('Excel generado (datos sintéticos)', () => {
  const rep = leerReporte(reporteSintetico([
    // Ana: extras el 1 (0), 2 (65), 3 (sin salida), 4 (18:20 < 18:30: 0; tardanza 71), 5 sáb (217), dom 6 (240), 7 (una marcada), 8 (17:52), 10 (21:30 → 2h + 1:30)
    { id: '10', nombre: 'Ana Uno', dias: { 1: '08:2918:14', 2: '08:4019:05', 3: '08:1013:0113:45', 4: '09:4118:20', 5: '08:0415:37', 6: '09:0013:00', 7: '08:30', 8: '08:3017:52', 10: '08:2521:30' } },
    { id: '11', nombre: 'Beto Dos', dias: { 1: '06:3718:01', 2: '08:2918:2918:30', 3: '07:5021:00', 9: '08:2018:00' } },
  ]));
  const { config: cfg } = sincronizarPersonas(rep, configInicial(2026));
  cfg.personas[0].reparto = true;
  cfg.personas[0].sueldo = 1800; // valor hora 7.50
  cfg.personas[1].fechaIngreso = '2026-09-02';
  cfg.personas[1].sueldo = 2287; // + asignación 113 = 2400 → 10.00
  cfg.personas[1].asignacionFamiliar = true;
  cfg.personas[1].horasSemanales = 40;
  cfg.personas[1].valorHoraSegun = 'reales'; // 2400 ÷ 30 ÷ 6:40 h = 12.00
  const mes = construirMes(rep, cfg, '2026-09-12');
  const dir = mkdtempSync(join(tmpdir(), 'qalen-sint-'));

  it('Resumen y cada día coinciden con la página (recalculado sin valores guardados)', async () => {
    const f = join(dir, 'a.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep, { sinResultados: true }))));
    const x = XLSX.read(readFileSync(recalcular(f)));
    expect(x.SheetNames).toEqual(['Leeme', 'Asistencia', 'Semanal', 'Resumen']);
    compararConPagina(x, mes);
    const a = mes.filas[0].tot;
    expect([a.extra1, a.extra2, a.enFeriado]).toEqual([65 + 120 + 120, 97 + 90, 240]);
    expect(a.monto).toBeGreaterThan(0);
    expect([a.valorHora, mes.filas[1].tot.valorHora]).toEqual([7.5, 12]);
    expect(mes.filas[1].tot.totalMes).toBe(2400 + mes.filas[1].tot.monto);
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
    // Se descuentan: día 2 (65), día 5 sáb (97 al 35% y 120 al 25%) y 18 min del 35% del día 10 (de 1:30 quedan 1:12).
    set(0, 'motivo', '2026-09-07', 'Compensado');
    set(0, 'salida', '2026-09-07', 13 / 24);
    // Beto (40 h/semana → jornada 6:40): 4/9 y 5/9 sin marcas → Compensado (400 + 210) y 7/9 → Falta (400, aparte)
    set(1, 'motivo', '2026-09-04', 'Compensado');
    set(1, 'motivo', '2026-09-05', 'Compensado');
    set(1, 'motivo', '2026-09-07', 'Falta');
    // En el Resumen: Ana pasa a sueldo 2400 con asignación (valor hora 10.47); Beto vuelve a 30 días × 8 h (10.00)
    const wr = wb.getWorksheet(HOJAS.resumen)!;
    wr.getCell(RES.primeraFila, RES.sueldo).value = 2400;
    wr.getCell(RES.primeraFila, RES.asignacion).value = 'Sí';
    wr.getCell(RES.primeraFila + 1, RES.segun).value = '30 días × 8 h';
    const g = join(dir, 'b2.xlsx');
    await wb.xlsx.writeFile(g);
    const x = XLSX.read(readFileSync(recalcular(g)));
    const res = x.Sheets[HOJAS.resumen];
    const asis = x.Sheets[HOJAS.asis];
    const r0 = RES.primeraFila;
    const a = mes.filas[0].tot;
    expect([min(res, r0, RES.compensado), min(res, r0, RES.neto2), min(res, r0, RES.neto1), min(res, r0, RES.nosDebe)])
      .toEqual([300, a.extra2 - 97 - 18, a.extra1 - (65 + 120), 0]);
    expect(min(asis, filaBloque(0) + FILA.neto1, col('2026-09-05'))).toBe(0);
    expect(min(asis, filaBloque(0) + FILA.neto2, col('2026-09-10'))).toBe(72);
    expect(min(asis, filaBloque(0) + FILA.neto1, col('2026-09-10'))).toBe(120);
    // Beto: 610 compensadas contra sus extras (todas consumidas), el resto "nos debe"; la falta va en contra
    const b = mes.filas[1].tot;
    expect([min(res, r0 + 1, RES.neto1), min(res, r0 + 1, RES.neto2), min(res, r0 + 1, RES.nosDebe)]).toEqual([0, 0, 610 - b.extra1 - b.extra2]);
    expect([min(res, r0 + 1, RES.contra), num(res, r0 + 1, RES.faltas)]).toEqual([b.contra + 400, 1]);
    expect([num(res, r0 + 1, RES.valorHora), num(res, r0 + 1, RES.monto), num(res, r0 + 1, RES.totalMes)]).toEqual([10, 0, 2400]);
    // Monto de Ana = netos × valor hora × recargo
    const n1 = min(res, r0, RES.neto1), n2 = min(res, r0, RES.neto2), fer = min(res, r0, RES.feriado);
    expect(num(res, r0, RES.valorHora)).toBe(10.47);
    // El pago es la suma de los soles de cada día (cada parte redondeada a céntimos)
    const asisCol = (fecha: string) => col(fecha);
    const diaTotal = (fecha: string) => num(asis, filaBloque(0) + FILA.solesTotal, asisCol(fecha));
    expect(num(asis, filaBloque(0) + FILA.soles2, col('2026-09-10'))).toBe(Math.round((72 / 60) * 10.47 * 1.35 * 100) / 100);
    expect(num(asis, filaBloque(0) + FILA.soles1, col('2026-09-10'))).toBe(Math.round(2 * 10.47 * 1.25 * 100) / 100);
    expect(diaTotal('2026-09-06')).toBe(Math.round(4 * 10.47 * 2 * 100) / 100); // domingo: 4 h al 100%
    expect(diaTotal('2026-09-02')).toBe(0); // sus extras se usaron para compensar
    const pago = mes.fechas.reduce((acc, f) => acc + diaTotal(f.fecha), 0);
    expect(num(res, r0, RES.monto)).toBeCloseTo(pago, 2);
    expect(num(res, r0, RES.totalMes)).toBeCloseTo(2400 + 113 + pago, 2);
    expect(n1 + n2 + fer).toBeGreaterThan(0);
    expect(num(res, r0, RES.pendientes)).toBe(a.pendientes - 1);
  });
});

describe.skipIf(!hayLibreOffice)('Reporte que abarca dos meses', () => {
  const rep = leerReporte(reporteSintetico([
    { id: '20', nombre: 'Carla Tres', dias: { '2026-09-29': '08:2519:00', '2026-09-30': '08:30', '2026-10-01': '08:2821:30', '2026-10-03': '08:2513:00' } },
  ], '2026-09-28', '2026-10-03'));
  const { config: cfg } = sincronizarPersonas(rep, configInicial(2026));
  cfg.personas[0].sueldo = 2400;
  const mes = construirMes(rep, cfg, corteMes(rep));

  it('analiza todo el rango y el Excel coincide con la página', async () => {
    expect(mes.fechas.map((f) => f.fecha)).toEqual(['2026-09-28', '2026-09-29', '2026-09-30', '2026-10-01', '2026-10-02', '2026-10-03']);
    const t = mes.filas[0].tot;
    expect([t.extra1, t.extra2]).toEqual([60 + 120 + 60, 90]); // 29/09: 1 h; 01/10: 3:30 (2 h + 1:30); sáb 03/10: 1 h
    const dir = mkdtempSync(join(tmpdir(), 'qalen-2m-'));
    const f = join(dir, 'a.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep, { sinResultados: true }))));
    const x = XLSX.read(readFileSync(recalcular(f)));
    compararConPagina(x, mes);
    expect(String(x.Sheets[HOJAS.asis]['D1']?.v)).toBe('ASISTENCIA SETIEMBRE – OCTUBRE 2026');
  });
});
