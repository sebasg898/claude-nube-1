/** Continuar un Excel ya trabajado con un reporte nuevo del reloj. */
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { describe, expect, it } from 'vitest';
import { aplicarSueldos, combinarConAnterior, leerExcelAnterior } from '../src/lib/anterior';
import { configInicial } from '../src/lib/config';
import { excelABuffer, FILA, filaBloque, generarExcel, HOJAS, PRIMERA_COL_DIA, RES } from '../src/lib/excel';
import { construirMes, corteMes, sincronizarPersonas, type Mes } from '../src/lib/model';
import { leerReporte } from '../src/lib/parser';
import { compararConPagina } from './comparar';
import { hayLibreOffice, recalcular } from './libreoffice';
import { reporteSintetico } from './sintetico';

const h = (s: string) => {
  const [a, b] = s.split(':').map(Number);
  return a * 60 + b;
};

// Primer corte: 01/09 al 12/09
const ANA_1 = { '2026-09-01': '08:2518:14', '2026-09-02': '08:30', '2026-09-03': '08:2519:00', '2026-09-04': '', '2026-09-08': '08:2918:05', '2026-09-10': '08:3018:00' };
// Reporte nuevo: 08/09 al 19/09 (se traslapa del 08 al 12). El 10/09 ahora trae una salida distinta (18:45).
const ANA_2 = { '2026-09-08': '08:2918:05', '2026-09-10': '08:3018:45', '2026-09-15': '08:2020:00', '2026-09-16': '08:30' };

async function excelCompletado(dir: string) {
  const rep1 = leerReporte(reporteSintetico([{ id: '10', nombre: 'Ana Uno', dias: ANA_1 }, { id: '11', nombre: 'Beto Dos', dias: { '2026-09-01': '08:3018:00' } }], '2026-09-01', '2026-09-12'));
  const { config: cfg } = sincronizarPersonas(rep1, configInicial(2026));
  const mes1 = construirMes(rep1, cfg, corteMes(rep1));
  const f = join(dir, 'corte1.xlsx');
  writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes1, cfg, rep1))));
  // Yasmín completa y guarda
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(f);
  const ws = wb.getWorksheet(HOJAS.asis)!;
  const col = (fecha: string) => PRIMERA_COL_DIA + mes1.fechas.findIndex((x) => x.fecha === fecha);
  ws.getCell(filaBloque(0) + FILA.salida, col('2026-09-02')).value = h('18:40') / 1440; // a mano (avisó por WhatsApp)
  ws.getCell(filaBloque(0) + FILA.motivo, col('2026-09-04')).value = 'Vacaciones';
  ws.getCell(filaBloque(0) + FILA.salida, col('2026-09-10')).value = h('18:00') / 1440; // ya era del reloj: no es edición
  ws.getCell(filaBloque(1) + FILA.motivo, col('2026-09-02')).value = 'Falta';
  const res = wb.getWorksheet(HOJAS.resumen)!;
  res.getCell(RES.primeraFila, RES.sueldo).value = 2400;
  res.getCell(RES.primeraFila, RES.asignacion).value = 'Sí';
  const g = join(dir, 'corte1-completado.xlsx');
  await wb.xlsx.writeFile(g);
  return { cfg, archivo: g };
}

describe('Continuar un Excel ya trabajado', () => {
  const dir = mkdtempSync(join(tmpdir(), 'qalen-cont-'));

  it('lee lo que ella completó (horas a mano, motivos y sueldos)', async () => {
    const { archivo } = await excelCompletado(dir);
    const ant = leerExcelAnterior(readFileSync(archivo));
    expect([ant.inicio, ant.fin]).toEqual(['2026-09-01', '2026-09-12']);
    const ana = ant.personas.find((p) => p.alias === 'ANA')!;
    expect(ana.dias['2026-09-02']).toEqual({ reloj: [h('08:30')], entrada: h('08:30'), salida: h('18:40'), motivo: '' });
    expect(ana.dias['2026-09-04'].motivo).toBe('Vacaciones');
    expect(ant.sueldos.find((s) => s.alias === 'ANA')).toEqual({ alias: 'ANA', sueldo: 2400, asignacion: true, segun: 'ocho' });
  });

  it('cortes que se traslapan: rango completo, se copia lo hecho y solo quedan pendientes los días nuevos', async () => {
    const { cfg, archivo } = await excelCompletado(dir);
    const ant = leerExcelAnterior(readFileSync(archivo));
    const { config } = aplicarSueldos(ant, cfg);
    expect(config.personas[0]).toMatchObject({ sueldo: 2400, asignacionFamiliar: true });
    const rep2 = leerReporte(reporteSintetico([{ id: '10', nombre: 'Ana Uno', dias: ANA_2 }, { id: '11', nombre: 'Beto Dos', dias: {} }], '2026-09-08', '2026-09-19'));
    const c = combinarConAnterior(rep2, ant, config);
    expect([c.rep.inicio, c.rep.fin, c.rep.sinDatos]).toEqual(['2026-09-01', '2026-09-19', []]);
    expect(c.copiados).toBe(3); // salida a mano 02/09, Vacaciones 04/09, Falta de Beto 02/09
    const mes = construirMes(c.rep, config, corteMes(c.rep), c.ediciones);
    const dia = (m: Mes, i: number, f: string) => m.filas[i].dias.find((d) => d.fecha === f)!;
    // Días que solo estaban en el Excel anterior: marcadas del Excel + lo que completó
    expect(dia(mes, 0, '2026-09-01')).toMatchObject({ entrada: h('08:25'), salida: h('18:14') });
    expect(dia(mes, 0, '2026-09-02')).toMatchObject({ entrada: h('08:30'), salida: h('18:40'), calc: { pendiente: false, extra1: 40 } });
    expect(dia(mes, 0, '2026-09-04')).toMatchObject({ motivo: 'Vacaciones', calc: { pendiente: false } });
    expect(dia(mes, 1, '2026-09-02').calc).toMatchObject({ falta: true, contra: 480 });
    // Días del traslape: manda el reloj nuevo (el 10/09 era del reloj, ahora 18:45)
    expect(dia(mes, 0, '2026-09-10').salida).toBe(h('18:45'));
    // Días nuevos
    expect(dia(mes, 0, '2026-09-15').calc.extra1).toBe(120);
    expect(dia(mes, 0, '2026-09-16').calc.pendiente).toBe(true);
    // Lo que ella ya completó deja de estar pendiente: Ana tiene 2 pendientes menos que sin el Excel anterior (02/09 y 04/09)
    const sinAnterior = construirMes(c.rep, config, corteMes(c.rep));
    expect(sinAnterior.filas[0].tot.pendientes - mes.filas[0].tot.pendientes).toBe(2);
    expect(sinAnterior.filas[1].tot.pendientes - mes.filas[1].tot.pendientes).toBe(1);
    expect(mes.filas[0].tot.valorHora).toBe(10.47);

    if (hayLibreOffice) {
      const f = join(dir, 'final.xlsx');
      writeFileSync(f, Buffer.from(await excelABuffer(await generarExcel(mes, config, c.rep, { sinResultados: true }))));
      compararConPagina(XLSX.read(readFileSync(recalcular(f))), mes);
    }
  });

  it('si ella escribió una hora y ahora el reloj trae otra, se respeta la suya con aviso', async () => {
    const { cfg, archivo } = await excelCompletado(dir);
    const ant = leerExcelAnterior(readFileSync(archivo));
    // El reporte nuevo ahora sí trae la salida del 02/09 (18:10), distinta de la que ella escribió (18:40)
    const rep2 = leerReporte(reporteSintetico([{ id: '10', nombre: 'Ana Uno', dias: { '2026-09-02': '08:3018:10' } }], '2026-09-01', '2026-09-19'));
    const c = combinarConAnterior(rep2, ant, cfg);
    const mes = construirMes(c.rep, cfg, corteMes(c.rep), c.ediciones);
    expect(mes.filas[0].dias.find((d) => d.fecha === '2026-09-02')!.salida).toBe(h('18:40'));
    expect(mes.avisos.some((a) => a.includes('ANA 02/09') && a.includes('18:40') && a.includes('18:10'))).toBe(true);
  });

  it('con un hueco de días entre el Excel anterior y el reporte nuevo, avisa y deja esos días en amarillo', async () => {
    const { cfg, archivo } = await excelCompletado(dir);
    const ant = leerExcelAnterior(readFileSync(archivo));
    const rep3 = leerReporte(reporteSintetico([{ id: '10', nombre: 'Ana Uno', dias: { '2026-09-17': '08:3018:00' } }], '2026-09-17', '2026-09-19'));
    const c = combinarConAnterior(rep3, ant, cfg);
    expect([c.rep.inicio, c.rep.fin]).toEqual(['2026-09-01', '2026-09-19']);
    expect(c.rep.sinDatos).toEqual(['2026-09-14', '2026-09-15', '2026-09-16']); // el domingo 13 no hace falta
    expect(c.avisos.some((a) => a.includes('14/09') && a.includes('16/09'))).toBe(true);
    const mes = construirMes(c.rep, cfg, corteMes(c.rep), c.ediciones);
    const d14 = mes.filas[0].dias.find((d) => d.fecha === '2026-09-14')!;
    expect(d14.calc.pendiente).toBe(true);
    expect(d14.casos[0].mensaje).toContain('no está en ningún reporte');
  });

  it('rechaza archivos que no son un Excel de esta página', () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['hola']]), 'Hoja1');
    expect(() => leerExcelAnterior(new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' })))).toThrow(/Asistencia/);
  });
});
