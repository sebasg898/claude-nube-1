import { describe, expect, it } from 'vitest';
import { configInicial, personaNueva, type Persona } from '../src/lib/config';
import { calcularDia, clasificar, deduplicar } from '../src/lib/rules';
import { parseHHMM } from '../src/lib/time';

const cfg = configInicial(2026);
const h = (s: string) => parseHHMM(s)!;
const base = (cambios: Partial<Persona> = {}): Persona => ({ ...personaNueva({ id: '1', nombre: 'Prueba' }, new Set(), true), ...cambios });
type Tipo = 'LV' | 'SAB' | 'DOM' | 'FER';
const calc = (e: string | null, s: string | null, motivo = '', tipo: Tipo = 'LV', p = base(), r = cfg.reglas) =>
  calcularDia({ entrada: e ? h(e) : null, salida: s ? h(s) : null, motivo }, tipo, p, r, true);
const dia = (marcas: string[], tipo: Tipo = 'LV', p = base()) => {
  const c = clasificar(marcas.map(h), tipo, p, true);
  return { c, r: calcularDia({ entrada: c.entrada, salida: c.salida, motivo: '' }, tipo, p, cfg.reglas, true) };
};

describe('marcadas', () => {
  it('quita duplicadas a menos de 3 minutos', () => {
    expect(deduplicar([h('08:29'), h('08:30'), h('18:01'), h('18:05')])).toEqual({ marcas: [h('08:29'), h('18:01'), h('18:05')], duplicadas: [h('08:30')] });
    const { c } = dia(['08:29', '08:30', '13:00', '13:45', '18:02', '18:03']);
    expect([c.entrada, c.salida, c.almuerzoSale, c.almuerzoVuelve]).toEqual([h('08:29'), h('18:02'), h('13:00'), h('13:45')]);
  });

  it('sin salida (la última es el almuerzo): queda por completar, sin inventar nada', () => {
    const { c, r } = dia(['08:10', '13:01', '13:45']);
    expect([c.entrada, c.salida]).toEqual([h('08:10'), null]);
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

  it('día sin marcadas: por completar', () => {
    const { c, r } = dia([]);
    expect(c.casos[0]).toMatchObject({ tipo: 'sin_marcadas', faltaDato: true });
    expect(r.pendiente).toBe(true);
  });

  it('horas raras se resaltan pero se usan', () => {
    const { c } = dia(['06:15', '22:36']);
    expect([c.entrada, c.salida]).toEqual([h('06:15'), h('22:36')]);
    expect(c.casos.some((x) => x.tipo === 'hora_rara' && !x.faltaDato)).toBe(true);
  });
});

describe('horas a favor / en contra', () => {
  it('salida 18:19 no da extras; 18:20 sí (desde las 18:00)', () => {
    expect(calc('08:30', '18:19').favor).toBe(0);
    expect(calc('08:30', '18:20').favor).toBe(20);
    expect(calc('08:30', '19:05').favor).toBe(65);
    expect(calc('08:30', '21:00').favor).toBe(180);
  });

  it('toda tardanza (después de las 08:30) es hora en contra', () => {
    expect(calc('08:30', '18:00').contra).toBe(0);
    expect(calc('08:31', '18:00').contra).toBe(1);
    expect(calc('09:41', '18:00').contra).toBe(71);
    expect(calc('08:30', '17:52').contra).toBe(8);
  });

  it('llegar antes de las 08:00 es hora a favor (hasta las 08:30)', () => {
    expect(calc('08:05', '18:00').favor).toBe(0);
    expect(calc('07:59', '18:00').favor).toBe(31);
    expect(calc('06:37', '18:01').favor).toBe(113);
  });

  it('sábado normal, con extras y salida temprana', () => {
    expect(calc('08:26', '12:19', '', 'SAB')).toMatchObject({ favor: 0, contra: 0, trabajado: 233 });
    expect(calc('08:04', '15:37', '', 'SAB').favor).toBe(217);
    expect(calc('08:30', '10:42', '', 'SAB').contra).toBe(78);
    expect(dia(['07:51', '10:42'], 'SAB').r.pendiente).toBe(true);
  });

  it('feriado / domingo trabajado: todo suma a favor', () => {
    expect(calc('09:00', '13:00', '', 'FER')).toMatchObject({ favor: 240, enFeriado: 240, contra: 0 });
    expect(calc('09:00', '13:00', '', 'DOM', base(), { ...cfg.reglas, factorFeriado: 2 }).favor).toBe(480);
  });

  it('salida antes que la entrada (formato 12 h) da alerta', () => {
    expect(calc('08:30', '06:14').alerta).toBe(true);
  });

  it('antes de la fecha de ingreso no cuenta nada', () => {
    const r = calcularDia({ entrada: null, salida: null, motivo: '' }, 'LV', base(), cfg.reglas, false);
    expect(r).toMatchObject({ pendiente: false, favor: 0, contra: 0, compensado: 0 });
  });
});

describe('motivos', () => {
  it('Vacaciones y Salud: día completo trabajado, sin nada en contra', () => {
    for (const m of ['Vacaciones', 'Salud']) {
      expect(calc(null, null, m)).toMatchObject({ pendiente: false, contra: 0, favor: 0, compensado: 0, trabajado: 570 - 45 });
      expect(calc(null, null, m, 'SAB')).toMatchObject({ contra: 0, trabajado: 210 });
    }
  });

  it('Falta: la jornada es hora en contra y cuenta como falta', () => {
    expect(calc(null, null, 'Falta')).toMatchObject({ pendiente: false, falta: true, contra: 480, favor: 0 });
    expect(calc(null, null, 'Falta', 'SAB').contra).toBe(210);
  });

  it('Estudio: horas en contra (día completo o lo que faltó)', () => {
    expect(calc(null, null, 'Estudio')).toMatchObject({ pendiente: false, contra: 480 });
    expect(calc('08:30', '13:00', 'Estudio')).toMatchObject({ contra: 300, compensado: 0 });
    expect(calc('08:30', null, 'Estudio').pendiente).toBe(true);
  });

  it('Compensado: lo que falta va a compensadas, no a en contra', () => {
    expect(calc(null, null, 'Compensado')).toMatchObject({ pendiente: false, contra: 0, compensado: 480 });
    expect(calc('08:40', '16:00', 'Compensado')).toMatchObject({ contra: 0, compensado: 10 + 120 });
    expect(calc('08:30', null, 'Compensado').pendiente).toBe(true);
  });

  it('medio tiempo / jornada distinta y personas que no descuentan', () => {
    expect(calc(null, null, 'Falta', 'LV', base({ horasSemanales: 24 })).contra).toBe(240);
    expect(calc(null, null, 'Falta', 'LV', base({ descuenta: false }))).toMatchObject({ contra: 0, falta: true });
    const seb = base({ horasSemanales: 24, salidaLV: h('12:00'), extrasDesdeLV: h('18:00'), descuenta: false });
    expect(calc('08:31', '12:14', '', 'LV', seb)).toMatchObject({ favor: 0, contra: 0, pendiente: false });
    expect(calc('08:31', '18:47', '', 'LV', seb).favor).toBe(47);
  });
});
