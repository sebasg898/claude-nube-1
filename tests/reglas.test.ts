import { describe, expect, it } from 'vitest';
import { configInicial, personaNueva, type Persona } from '../src/lib/config';
import { calcularDia, clasificar, deduplicar } from '../src/lib/rules';
import { parseHHMM } from '../src/lib/time';

const cfg = configInicial(2026);
const h = (s: string) => parseHHMM(s)!;
const base = (cambios: Partial<Persona> = {}): Persona => ({ ...personaNueva({ id: '1', nombre: 'Prueba' }, new Set(), true), ...cambios });
const dia = (marcas: string[], tipo: 'LV' | 'SAB' | 'DOM' | 'FER' = 'LV', p = base(), motivo = '') => {
  const c = clasificar(marcas.map(h), tipo, p, true);
  return { c, r: calcularDia({ entrada: c.entrada, salida: c.salida, motivo }, tipo, p, cfg.reglas, cfg.motivos, true) };
};

describe('marcadas', () => {
  it('quita duplicadas a menos de 3 minutos', () => {
    expect(deduplicar([h('08:29'), h('08:30'), h('18:01'), h('18:05')])).toEqual({ marcas: [h('08:29'), h('18:01'), h('18:05')], duplicadas: [h('08:30')] });
    const { c } = dia(['08:29', '08:30', '13:00', '13:45', '18:02', '18:03']);
    expect([c.entrada, c.salida, c.almuerzoSale, c.almuerzoVuelve]).toEqual([h('08:29'), h('18:02'), h('13:00'), h('13:45')]);
  });

  it('sin salida (la última es el almuerzo): queda por completar, sin inventar nada', () => {
    const { c, r } = dia(['08:10', '13:01', '13:45']);
    expect(c.entrada).toBe(h('08:10'));
    expect(c.salida).toBeNull();
    expect(c.casos.map((x) => [x.tipo, x.faltaDato])).toEqual([['sin_salida', true]]);
    expect(r.pendiente).toBe(true);
  });

  it('reparto sin salida: sugiere 18:00 pero no lo acepta solo', () => {
    const { c } = dia(['08:10', '13:01'], 'LV', base({ reparto: true }));
    expect(c.salida).toBeNull();
    expect(c.casos[0].sugerencia).toBe('18:00');
  });

  it('sin entrada y una sola marcada', () => {
    expect(dia(['15:02', '15:34', '18:04']).c.casos[0].tipo).toBe('sin_entrada');
    const una = dia(['08:25']).c;
    expect([una.entrada, una.salida, una.casos[0].tipo]).toEqual([h('08:25'), null, 'una_marcada']);
    expect(dia(['18:14']).c.salida).toBe(h('18:14'));
  });

  it('día sin marcadas: por completar (sugerencia Falta)', () => {
    const { c, r } = dia([]);
    expect(c.casos[0]).toMatchObject({ tipo: 'sin_marcadas', faltaDato: true, sugerencia: 'Falta' });
    expect(r.pendiente).toBe(true);
    const conMotivo = calcularDia({ entrada: null, salida: null, motivo: 'Falta' }, 'LV', base(), cfg.reglas, cfg.motivos, true);
    expect([conMotivo.pendiente, conMotivo.falta, conMotivo.contra]).toEqual([false, true, 0]);
  });

  it('horas raras se resaltan pero se usan', () => {
    const { c } = dia(['06:15', '22:36']);
    expect([c.entrada, c.salida]).toEqual([h('06:15'), h('22:36')]);
    expect(c.casos.some((x) => x.tipo === 'hora_rara' && !x.faltaDato)).toBe(true);
  });
});

describe('bolsa de horas', () => {
  it('salida 18:19 no da extras; 18:20 sí (desde las 18:00)', () => {
    expect(dia(['08:30', '18:19']).r.favor).toBe(0);
    expect(dia(['08:30', '18:20']).r.favor).toBe(20);
    expect(dia(['08:30', '19:05']).r.favor).toBe(65);
  });

  it('3 horas de extras', () => {
    expect(dia(['08:30', '21:00']).r.favor).toBe(180);
  });

  it('sábado normal y con extras', () => {
    expect(dia(['08:26', '12:19'], 'SAB').r).toMatchObject({ favor: 0, contra: 0, trabajado: 233 });
    expect(dia(['08:04', '15:37'], 'SAB').r.favor).toBe(217);
    // Última marcada antes de las 11:00 en sábado: no se asume salida, queda por completar
    expect(dia(['07:51', '10:42'], 'SAB').r.pendiente).toBe(true);
    const r = calcularDia({ entrada: h('07:51'), salida: h('10:42'), motivo: 'Permiso' }, 'SAB', base(), cfg.reglas, cfg.motivos, true);
    expect(r.contra).toBe(78); // salió antes de las 12:00
  });

  it('feriado / domingo trabajado: todo suma a favor', () => {
    expect(dia(['09:00', '13:00'], 'FER').r).toMatchObject({ favor: 240, enFeriado: 240, contra: 0 });
    const doble = calcularDia({ entrada: h('09:00'), salida: h('13:00'), motivo: '' }, 'DOM', base(), { ...cfg.reglas, factorFeriado: 2 }, cfg.motivos, true);
    expect(doble.favor).toBe(480);
  });

  it('tardanza solo resta desde el umbral; salida temprana siempre', () => {
    expect(dia(['08:59', '18:00']).r.contra).toBe(0);
    expect(dia(['09:41', '18:00']).r.contra).toBe(71);
    expect(dia(['08:30', '17:52']).r.contra).toBe(8);
    const sinUmbral = calcularDia({ entrada: h('08:31'), salida: h('18:00'), motivo: '' }, 'LV', base(), { ...cfg.reglas, tardanzaDesde: 0 }, cfg.motivos, true);
    expect(sinUmbral.contra).toBe(1);
  });

  it('llegar temprano solo suma a quien lo tiene configurado', () => {
    expect(dia(['06:37', '18:01']).r.favor).toBe(0);
    expect(dia(['06:37', '18:01'], 'LV', base({ tempranoAntesDe: h('07:45') })).r.favor).toBe(113);
    expect(dia(['07:50', '18:01'], 'LV', base({ tempranoAntesDe: h('07:45') })).r.favor).toBe(0);
  });

  it('permiso / estudia sin marcadas resta la jornada; no si la persona no descuenta', () => {
    const c = (p: Persona, tipo: 'LV' | 'SAB') => calcularDia({ entrada: null, salida: null, motivo: 'Estudia' }, tipo, p, cfg.reglas, cfg.motivos, true).contra;
    expect(c(base(), 'LV')).toBe(480);
    expect(c(base(), 'SAB')).toBe(210);
    expect(c(base({ horasSemanales: 24 }), 'LV')).toBe(240);
    expect(c(base({ descuenta: false }), 'LV')).toBe(0);
  });

  it('caso especial: medio tiempo con extras desde las 18:00', () => {
    const seb = base({ horasSemanales: 24, salidaLV: h('12:00'), extrasDesdeLV: h('18:00'), descuenta: false });
    expect(dia(['08:31', '12:14'], 'LV', seb).r).toMatchObject({ favor: 0, contra: 0, pendiente: false });
    expect(dia(['08:31', '18:47'], 'LV', seb).r.favor).toBe(47);
  });

  it('salida antes que la entrada (formato 12 h) da alerta', () => {
    const r = calcularDia({ entrada: h('08:30'), salida: h('06:14'), motivo: '' }, 'LV', base(), cfg.reglas, cfg.motivos, true);
    expect(r.alerta).toBe(true);
  });

  it('antes de la fecha de ingreso no cuenta nada', () => {
    const r = calcularDia({ entrada: null, salida: null, motivo: '' }, 'LV', base(), cfg.reglas, cfg.motivos, false);
    expect(r).toMatchObject({ pendiente: false, favor: 0, contra: 0 });
  });
});
