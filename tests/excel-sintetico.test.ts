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
import { celda, compararConPagina, min, saldo } from './comparar';
import { leerReporte } from '../src/lib/parser';
import { hayLibreOffice, recalcular } from './libreoffice';
import { reporteSintetico } from './sintetico';

describe.skipIf(!hayLibreOffice)('Excel generado (datos sintéticos)', () => {
  const rep = leerReporte(reporteSintetico([
    { id: '10', nombre: 'Ana Uno', dias: { 1: '08:2918:14', 2: '08:4019:05', 3: '08:1013:0113:45', 4: '09:4118:20', 5: '08:0415:37', 6: '09:0013:00', 7: '08:30', 8: '08:3017:52' } },
    { id: '11', nombre: 'Beto Dos', dias: { 1: '06:3718:01', 2: '08:2918:2918:30', 3: '07:5021:00', 9: '08:2018:00' } },
  ]));
  const { config: cfg } = sincronizarPersonas(rep, configInicial(2026));
  cfg.personas[0].reparto = true;
  cfg.personas[1].fechaIngreso = '2026-09-02';
  const mes = construirMes(rep, cfg, '2026-09-12');
  const dir = mkdtempSync(join(tmpdir(), 'qalen-sint-'));

  it('Resumen y cada día coinciden con la página (recalculado sin valores guardados)', async () => {
    const f = join(dir, 'a.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep, { sinResultados: true }))));
    const x = XLSX.read(readFileSync(recalcular(f)));
    expect(x.SheetNames).toEqual(['Leeme', 'Asistencia', 'Semanal', 'Resumen']);
    compararConPagina(x, mes);
    expect(mes.filas[0].tot.enFeriado).toBe(240); // domingo 6/9 trabajado
    expect(mes.filas[0].tot.pendientes).toBeGreaterThan(0);
  });

  it('completar celdas amarillas y elegir motivos recalcula todo y baja el conteo', async () => {
    const f = join(dir, 'b.xlsx');
    writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, cfg, rep))));
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(f);
    const ws = wb.getWorksheet(HOJAS.asis)!;
    const col = (fecha: string) => PRIMERA_COL_DIA + mes.fechas.findIndex((x) => x.fecha === fecha);
    const set = (i: number, fila: keyof typeof FILA, fecha: string, v: ExcelJS.CellValue) => {
      ws.getCell(filaBloque(i) + FILA[fila], col(fecha)).value = v;
    };
    // Ana: 3/9 escribe salida 19:00 (+60 a favor); 7/9 una marcada → Compensado con salida 13:00 (300 compensadas)
    set(0, 'salida', '2026-09-03', 19 / 24);
    set(0, 'motivo', '2026-09-07', 'Compensado');
    set(0, 'salida', '2026-09-07', 13 / 24);
    // Beto: 4/9 sin marcadas → Falta (−480, 1 falta); 5/9 sábado sin marcas → Vacaciones; 7/9 → Estudio (−480)
    set(1, 'motivo', '2026-09-04', 'Falta');
    set(1, 'motivo', '2026-09-05', 'Vacaciones');
    set(1, 'motivo', '2026-09-07', 'Estudio');
    const g = join(dir, 'b2.xlsx');
    await wb.xlsx.writeFile(g);
    const x = XLSX.read(readFileSync(recalcular(g)));
    const res = x.Sheets[HOJAS.resumen];
    const r0 = RES.primeraFila;
    const [a, b] = mes.filas.map((f) => f.tot);
    expect([min(res, r0, RES.favor), min(res, r0, RES.compensado), saldo(res, r0)]).toEqual([a.favor + 60, 300, a.saldo + 60 - 300]);
    expect([min(res, r0 + 1, RES.contra), Number(celda(res, r0 + 1, RES.faltas)), saldo(res, r0 + 1)]).toEqual([b.contra + 960, 1, b.saldo - 960]);
    expect(Number(celda(res, r0, RES.pendientes))).toBe(a.pendientes - 2);
    expect(Number(celda(res, r0 + 1, RES.pendientes))).toBe(b.pendientes - 3);
    // Vacaciones del sábado cuenta como día trabajado (3:30)
    expect(min(res, r0 + 1, RES.trabajado)).toBe(b.trabajado + 210);
    // Semanal: semana 2 de Beto (7/9 a 12/9) tiene las 8 h de Estudio en contra
    const sem = x.Sheets[HOJAS.semanal];
    const semBase = mes.filas[1].dias.filter((_, j) => mes.fechas[j].semana === 2).reduce((s, d) => s + d.calc.contra, 0);
    expect(min(sem, 8, 3 + 5 + 2)).toBe(semBase + 480);
  });
});
