import { describe, expect, it } from 'vitest';
import { dayType, feriadosPeru } from '../src/lib/calendar';
import { corteMes } from '../src/lib/model';
import { nombrePeriodo } from '../src/lib/time';
import { ErrorReporte, interpretarFilas, leerReporte } from '../src/lib/parser';
import { reporteSintetico } from './sintetico';

describe('lectura del reporte', () => {
  it('separa las horas pegadas y respeta las columnas', () => {
    const rep = leerReporte(reporteSintetico([{ id: '90000001', nombre: 'ana prueba', dias: { 1: '08:2913:4114:1218:01', 3: '08:25' } }]));
    expect(rep.inicio).toBe('2026-09-01');
    expect(rep.fechas).toHaveLength(12);
    const p = rep.personas[0];
    expect(p.id).toBe('90000001');
    expect(p.marcadas['2026-09-01']).toEqual([509, 821, 852, 1081]);
    expect(p.marcadas['2026-09-02']).toBeUndefined();
    expect(p.marcadas['2026-09-03']).toEqual([505]);
  });

  it('período que cruza de mes', () => {
    const filas = [['Periodo:', '2026-09-28 ~ 2026-10-02'], [28, 29, 30, 1, 2], ['ID:', '5', 'Nombre:', 'A'], ['', '', '', '08:00', '']];
    const rep = interpretarFilas(filas);
    expect(rep.fechas).toEqual(['2026-09-28', '2026-09-29', '2026-09-30', '2026-10-01', '2026-10-02']);
    expect(rep.personas[0].marcadas['2026-10-01']).toEqual([480]);
  });

  it('errores claros si no es el archivo correcto', () => {
    expect(() => interpretarFilas([['hola']])).toThrow(ErrorReporte);
    expect(() => leerReporte(new Uint8Array([1, 2, 3]))).toThrow(/reporte|abrir|hoja/i);
  });
});

describe('calendario', () => {
  const fer = new Set(feriadosPeru(2026).map((f) => f.fecha));
  it('feriados móviles 2026 y tipos de día', () => {
    expect(fer.has('2026-04-02') && fer.has('2026-04-03')).toBe(true);
    expect(dayType('2026-09-05', fer)).toBe('SAB');
    expect(dayType('2026-09-06', fer)).toBe('DOM');
    expect(dayType('2026-10-08', fer)).toBe('FER');
  });
  it('corte: todo el rango del reporte, aunque abarque varios meses', () => {
    const rep = interpretarFilas([['Periodo:', '2026-09-28 ~ 2026-10-15'], [28, 29, 30, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15], ['ID:', '5', 'Nombre:', 'A'], []]);
    expect(corteMes(rep)).toBe('2026-10-15');
    expect(nombrePeriodo('2026-09-01', '2026-09-26')).toBe('SETIEMBRE 2026');
    expect(nombrePeriodo('2026-09-28', '2026-10-15')).toBe('SETIEMBRE – OCTUBRE 2026');
    expect(nombrePeriodo('2026-12-20', '2027-01-10')).toBe('DICIEMBRE 2026 – ENERO 2027');
  });
});
