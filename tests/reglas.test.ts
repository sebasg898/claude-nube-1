import { describe, expect, it } from 'vitest';
import { configInicial, personaNueva, type Persona } from '../src/lib/config';
import { calcularDia, clasificar, deduplicar, descontarCompensadas } from '../src/lib/rules';
import { montoSoles, valorHora } from '../src/lib/model';
import { parseHHMM } from '../src/lib/time';

const cfg = configInicial(2026);
const h = (s: string) => parseHHMM(s)!;
const base = (cambios: Partial<Persona> = {}): Persona => ({ ...personaNueva({ id: '1', nombre: 'Prueba' }, new Set(), true), ...cambios });
type Tipo = 'LV' | 'SAB' | 'DOM' | 'FER';
const calc = (e: string | null, s: string | null, motivo = '', tipo: Tipo = 'LV', p = base(), r = cfg.reglas) =>
  calcularDia({ entrada: e ? h(e) : null, salida: s ? h(s) : null, motivo }, tipo, p, r, true);
const ext = (r: { extra1: number; extra2: number }) => r.extra1 + r.extra2;
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

describe('horas extras y en contra', () => {
  it('salida 18:29 no da extras; 18:30 sí (desde las 18:00)', () => {
    expect(ext(calc('08:30', '18:29'))).toBe(0);
    expect(calc('08:30', '18:30')).toMatchObject({ extra1: 30, extra2: 0 });
    expect(calc('08:30', '19:05')).toMatchObject({ extra1: 65, extra2: 0 });
    expect(ext(calc('08:30', '12:29', '', 'SAB'))).toBe(0);
    expect(ext(calc('08:30', '12:30', '', 'SAB'))).toBe(30);
  });

  it('cada día: 2 primeras horas al 25%, el resto al 35%', () => {
    expect(calc('08:30', '20:00')).toMatchObject({ extra1: 120, extra2: 0 });
    expect(calc('08:30', '21:00')).toMatchObject({ extra1: 120, extra2: 60 });
    expect(calc('08:30', '21:30')).toMatchObject({ extra1: 120, extra2: 90 });
  });

  it('toda tardanza (después de las 08:30) es hora en contra', () => {
    expect(calc('08:30', '18:00').contra).toBe(0);
    expect(calc('08:31', '18:00').contra).toBe(1);
    expect(calc('09:41', '18:00').contra).toBe(71);
    expect(calc('08:30', '17:52').contra).toBe(8);
  });

  it('llegar temprano no suma horas', () => {
    expect(ext(calc('07:59', '18:00'))).toBe(0);
    expect(ext(calc('06:37', '18:01'))).toBe(0);
  });

  it('sábado normal, con extras y salida temprana', () => {
    expect(calc('08:26', '12:19', '', 'SAB')).toMatchObject({ extra1: 0, extra2: 0, contra: 0, trabajado: 233 });
    expect(calc('08:04', '15:37', '', 'SAB')).toMatchObject({ extra1: 120, extra2: 97 });
    expect(calc('08:30', '10:42', '', 'SAB').contra).toBe(78);
    expect(dia(['07:51', '10:42'], 'SAB').r.pendiente).toBe(true);
  });

  it('feriado / domingo trabajado: va aparte (100%), no como extra', () => {
    expect(calc('09:00', '13:00', '', 'FER')).toMatchObject({ enFeriado: 240, extra1: 0, extra2: 0, contra: 0 });
    expect(calc('09:00', '20:00', '', 'DOM')).toMatchObject({ enFeriado: 660, extra1: 0 });
  });

  it('salida antes que la entrada (formato 12 h) da alerta', () => {
    expect(calc('08:30', '06:14').alerta).toBe(true);
  });

  it('antes de la fecha de ingreso no cuenta nada', () => {
    const r = calcularDia({ entrada: null, salida: null, motivo: '' }, 'LV', base(), cfg.reglas, false);
    expect(r).toMatchObject({ pendiente: false, extra1: 0, contra: 0, compensado: 0 });
  });
});

describe('motivos', () => {
  it('Vacaciones y Salud: día completo trabajado, sin nada en contra', () => {
    for (const m of ['Vacaciones', 'Salud']) {
      expect(calc(null, null, m)).toMatchObject({ pendiente: false, contra: 0, extra1: 0, compensado: 0, trabajado: 570 - 45 });
      expect(calc(null, null, m, 'SAB')).toMatchObject({ contra: 0, trabajado: 210 });
    }
  });

  it('Falta: la jornada es hora en contra y cuenta como falta', () => {
    expect(calc(null, null, 'Falta')).toMatchObject({ pendiente: false, falta: true, contra: 480, extra1: 0 });
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
    expect(calc('08:31', '12:14', '', 'LV', seb)).toMatchObject({ extra1: 0, contra: 0, pendiente: false });
    expect(calc('08:31', '18:47', '', 'LV', seb).extra1).toBe(47);
  });
});

describe('compensadas y soles', () => {
  const dias = [
    { extra1: 120, extra2: 30 }, // día 1: 2:30 extras
    { extra1: 0, extra2: 0 },
    { extra1: 60, extra2: 0 }, // día 3: 1:00
    { extra1: 120, extra2: 60 }, // día 4: 3:00
  ];

  it('se descuentan desde el primer día, primero las del 35%', () => {
    const r = descontarCompensadas(dias, 100);
    expect(r.netos).toEqual([{ neto1: 50, neto2: 0 }, { neto1: 0, neto2: 0 }, { neto1: 60, neto2: 0 }, { neto1: 120, neto2: 60 }]);
    expect(r.nosDebe).toBe(0);
    const r2 = descontarCompensadas(dias, 200); // 150 del día 1, 50 del día 3
    expect(r2.netos.map((n) => n.neto1 + n.neto2)).toEqual([0, 0, 10, 180]);
  });

  it('si las compensadas superan las extras, el resto es "nos debe"', () => {
    const r = descontarCompensadas(dias, 600);
    expect(r.netos.every((n) => n.neto1 === 0 && n.neto2 === 0)).toBe(true);
    expect(r.nosDebe).toBe(600 - 390);
    expect(descontarCompensadas(dias, 0).netos[3]).toEqual({ neto1: 120, neto2: 60 });
  });

  it('monto en soles con recargos de ley', () => {
    // 2 h al 25% + 1 h al 35% + 4 h de feriado, valor hora S/ 10
    expect(montoSoles({ neto1: 120, neto2: 60, enFeriado: 240 }, 10, cfg.reglas)).toBe(25 + 13.5 + 80);
    expect(montoSoles({ neto1: 20, neto2: 0, enFeriado: 0 }, 7.37, cfg.reglas)).toBe(3.07);
  });
});

describe('turno distinto en semana y sábado (caso Victor)', () => {
  const vic = base({ entrada: h('17:00'), salidaLV: h('22:00'), extrasDesdeLV: h('22:00'), entradaSab: h('08:30'), salidaSab: h('12:00'), horasSemanales: 28.5 });

  it('clasifica las marcadas de la tarde', () => {
    const { c, r } = dia(['16:55', '22:05'], 'LV', vic);
    expect([c.entrada, c.salida, c.casos]).toEqual([h('16:55'), h('22:05'), []]);
    expect(r).toMatchObject({ extra1: 0, contra: 0, trabajado: 310 });
    expect(dia(['16:58', '22:45'], 'LV', vic).r.extra1).toBe(45);
    expect(dia(['17:20', '22:00'], 'LV', vic).r.contra).toBe(20);
    expect(dia(['17:05'], 'LV', vic).c.casos[0].tipo).toBe('una_marcada');
  });

  it('el sábado usa su propia entrada', () => {
    const { c, r } = dia(['08:25', '12:40'], 'SAB', vic);
    expect([c.entrada, c.salida]).toEqual([h('08:25'), h('12:40')]);
    expect(r).toMatchObject({ extra1: 40, contra: 0 });
    expect(dia(['08:45', '12:00'], 'SAB', vic).r.contra).toBe(15);
    expect(calcularDia({ entrada: null, salida: null, motivo: 'Falta' }, 'LV', vic, cfg.reglas, true).contra).toBe(285);
  });

  it('valor hora: 30 días × 8 h o según sus horas reales', () => {
    const p = { ...vic, sueldo: 1500 };
    expect(valorHora(p, cfg.reglas, 'ocho')).toBe(6.25); // 1500 ÷ 30 ÷ 8
    expect(valorHora(p, cfg.reglas, 'reales')).toBe(10.53); // 1500 ÷ 30 ÷ 4.75
    expect(valorHora({ ...p, asignacionFamiliar: true }, cfg.reglas, 'ocho')).toBe(6.72); // (1500 + 113) ÷ 240
    expect(valorHora({ ...p, sueldo: 2000 }, cfg.reglas, 'ocho')).toBe(8.33);
    expect(valorHora({ ...p, sueldo: 0 }, cfg.reglas, 'ocho')).toBe(0);
  });
});

describe('soles por día', () => {
  it('cada parte se redondea a céntimos y se suma', async () => {
    const { solesDia, redondear2 } = await import('../src/lib/model');
    const d = { neto: { neto1: 120, neto2: 72 }, calc: { enFeriado: 0 } } as Parameters<typeof solesDia>[0];
    expect(solesDia(d, 10.47, cfg.reglas)).toEqual({ s1: 26.18, s2: 16.96, sf: 0, total: 43.14 });
    expect(solesDia({ neto: { neto1: 0, neto2: 0 }, calc: { enFeriado: 240 } } as Parameters<typeof solesDia>[0], 10.47, cfg.reglas).total).toBe(83.76);
    expect([redondear2(1.005), redondear2(2.675), redondear2(26.175)]).toEqual([1.01, 2.68, 26.18]);
  });
});
