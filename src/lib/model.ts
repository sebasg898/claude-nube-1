import { dayType, type DayType } from './calendar';
import { personaNueva, type Config, type Persona, type Reglas } from './config';
import type { Reporte } from './parser';
import type { Ediciones } from './anterior';
import { calcularDia, clasificar, descontarCompensadas, type Caso, type Clasificacion, type DiaCalculo, type Neto } from './rules';
import { daysBetween, ddmm, fmtHHMM, weekday, type ISODate, type Min } from './time';

export interface Dia {
  fecha: ISODate;
  tipo: DayType;
  aplica: boolean;
  reloj: string;
  clas: Clasificacion;
  entrada: Min | null;
  salida: Min | null;
  motivo: string;
  calc: DiaCalculo;
  /** Extras que quedan por pagar ese día después de descontar las compensadas. */
  neto: Neto;
  /** Pago en soles de ese día (con el valor hora elegido). */
  soles: SolesDia;
  casos: Caso[];
}

export interface SolesDia {
  /** Extras al primer recargo (25%). */
  s1: number;
  /** Extras al segundo recargo (35%). */
  s2: number;
  /** Domingo / feriado (100%). */
  sf: number;
  /** Total del día. */
  total: number;
}

const SIN_SOLES: SolesDia = { s1: 0, s2: 0, sf: 0, total: 0 };

export interface FilaPersona {
  persona: Persona;
  dias: Dia[];
  tot: Totales;
}

export interface Totales {
  trabajado: Min;
  base: Min;
  extra1: Min;
  extra2: Min;
  compensado: Min;
  neto1: Min;
  neto2: Min;
  enFeriado: Min;
  /** Compensadas que no alcanzaron a cubrirse con extras: horas que la persona debe. */
  nosDebe: Min;
  contra: Min;
  pendientes: number;
  faltas: number;
  dias: number;
  /** Valor hora usado (según cómo esté configurada la persona). */
  valorHora: number;
  /** Pago de horas extras en soles: suma de los totales de cada día (extras netas y domingo/feriado). */
  monto: number;
  /** El pago de extras con cada forma de calcular el valor hora (para comparar). */
  montoOcho: number;
  montoReales: number;
  /** Sueldo + asignación familiar + pago de extras. */
  totalMes: number;
}

export interface Mes {
  inicio: ISODate;
  hasta: ISODate;
  finReporte: ISODate;
  fechas: { fecha: ISODate; tipo: DayType; semana: number }[];
  filas: FilaPersona[];
  avisos: string[];
}

/** Agrega al config las personas del reloj que aún no existen. Devuelve los avisos. */
export function sincronizarPersonas(rep: Reporte, cfg: Config): { config: Config; avisos: string[] } {
  const avisos: string[] = [];
  const usados = new Set(cfg.personas.map((p) => p.alias.toUpperCase()));
  const personas = [...cfg.personas];
  for (const r of rep.personas) {
    const ex = personas.find((p) => p.id === r.id);
    const tiene = Object.keys(r.marcadas).length > 0;
    if (!ex) {
      const p = personaNueva(r, usados, tiene);
      personas.push(p);
      avisos.push(tiene
        ? `Persona nueva en el reloj: "${r.nombre}" (ID ${r.id}). Se agregó como ${p.alias} con horario normal; revisa sus datos en "Personas y reglas".`
        : `"${r.nombre}" (ID ${r.id}) está en el reloj pero no marcó en todo el período; se dejó como inactiva.`);
    } else {
      if (ex.nombreReloj !== r.nombre) personas[personas.indexOf(ex)] = { ...ex, nombreReloj: r.nombre };
      if (!ex.activo && tiene) avisos.push(`${ex.alias} está marcada como inactiva pero tiene marcadas este mes. No aparecerá en el Excel hasta que la actives.`);
      if (ex.activo && !tiene) avisos.push(`${ex.alias} está activa pero no marcó en todo el período: sus días saldrán en amarillo.`);
    }
  }
  return { config: { ...cfg, personas }, avisos };
}

/** Corte por defecto: todo el rango que trae el reporte (aunque abarque varios meses). */
export function corteMes(rep: Reporte): ISODate {
  return rep.fin;
}

export function construirMes(rep: Reporte, cfg: Config, hasta: ISODate, ediciones: Ediciones = {}): Mes {
  const sinDatos = new Set(rep.sinDatos ?? []);
  const conflictos: string[] = [];
  const feriados = new Set(cfg.feriados.map((f) => f.fecha));
  const fin = hasta > rep.fin ? rep.fin : hasta < rep.inicio ? rep.inicio : hasta;
  const activos = cfg.personas.filter((p) => p.activo);
  const relojPorId = new Map(rep.personas.map((r) => [r.id, r]));

  // Columnas: todos los días menos los domingos en que nadie trabajó (como lo hace Yasmín).
  const fechas: Mes['fechas'] = [];
  let semana = 0;
  let prevW = 99;
  for (const f of daysBetween(rep.inicio, fin)) {
    const tipo = dayType(f, feriados);
    const w = weekday(f);
    if (w < prevW || fechas.length === 0) semana++;
    prevW = w;
    const alguien = activos.some((p) => (relojPorId.get(p.id)?.marcadas[f]?.length ?? 0) > 0);
    if (weekday(f) === 7 && !alguien) continue;
    fechas.push({ fecha: f, tipo, semana });
  }

  const filas: FilaPersona[] = activos.map((p) => {
    const r = relojPorId.get(p.id);
    const dias: Dia[] = fechas.map(({ fecha, tipo }) => {
      const aplica = !p.fechaIngreso || fecha >= p.fechaIngreso;
      const marc = r?.marcadas[fecha] ?? [];
      const clas = clasificar(marc, tipo, p, aplica);
      // Antes de la fecha de ingreso no se cuenta nada: las marcadas quedan solo como referencia.
      let entrada = aplica ? clas.entrada : null;
      let salida = aplica ? clas.salida : null;
      let motivo = '';
      const casos = [...clas.casos];
      if (sinDatos.has(fecha)) {
        for (const c of casos) if (c.tipo === 'sin_marcadas') c.mensaje = 'Este día no está en ningún reporte del reloj: saca un reporte que lo incluya o complétalo a mano.';
      }
      // Lo que Yasmín ya completó en un Excel anterior (se respeta aunque el reloj diga otra cosa)
      const ed = ediciones[p.id]?.[fecha];
      if (ed && aplica) {
        const copiado: string[] = [];
        if (ed.entrada != null) {
          if (entrada != null && entrada !== ed.entrada) conflictos.push(`${p.alias} ${ddmm(fecha)}: entrada escrita ${fmtHHMM(ed.entrada)}, el reloj tiene ${fmtHHMM(entrada)}`);
          entrada = ed.entrada;
          copiado.push(`entrada ${fmtHHMM(ed.entrada)}`);
        }
        if (ed.salida != null) {
          if (salida != null && salida !== ed.salida) conflictos.push(`${p.alias} ${ddmm(fecha)}: salida escrita ${fmtHHMM(ed.salida)}, el reloj tiene ${fmtHHMM(salida)}`);
          salida = ed.salida;
          copiado.push(`salida ${fmtHHMM(ed.salida)}`);
        }
        if (ed.motivo) {
          motivo = ed.motivo;
          copiado.push(`motivo ${ed.motivo}`);
        }
        casos.push({ tipo: 'copiado', faltaDato: false, mensaje: `Copiado del Excel anterior: ${copiado.join(', ')}.` });
      }
      const calc = calcularDia({ entrada, salida, motivo }, tipo, p, cfg.reglas, aplica);
      const ext = calc.extra1 + calc.extra2;
      if (ext > 180) casos.push({ tipo: 'extras_altas', faltaDato: false, mensaje: `Más de 3 horas extras en el día (${Math.floor(ext / 60)} h ${ext % 60} min). Solo para revisar.` });
      return { fecha, tipo, aplica, reloj: clas.marcas.map(fmtHHMM).join(' '), clas, entrada, salida, motivo, calc, neto: { neto1: 0, neto2: 0 }, soles: SIN_SOLES, casos };
    });
    return { persona: p, dias, tot: totales(p, dias, cfg.reglas) };
  });

  const avisos: string[] = [];
  if (fin < rep.fin) avisos.push(`Se calcula hasta el ${ddmm(fin)}; el período llega hasta el ${ddmm(rep.fin)}.`);
  if (conflictos.length) avisos.push(`Se respetaron ${conflictos.length} hora(s) que escribiste a mano aunque ahora el reloj tiene otra marcada: ${conflictos.join('; ')}.`);
  return { inicio: rep.inicio, hasta: fin, finReporte: rep.fin, fechas, filas, avisos };
}

/** Totales de una persona. También reparte las compensadas sobre las extras (guarda el neto de cada día). */
export function totales(p: Persona, dias: Dia[], r: Reglas): Totales {
  const t: Totales = {
    trabajado: 0, base: 0, extra1: 0, extra2: 0, compensado: 0, neto1: 0, neto2: 0, enFeriado: 0, nosDebe: 0, contra: 0,
    pendientes: 0, faltas: 0, dias: 0, valorHora: 0, monto: 0, montoOcho: 0, montoReales: 0, totalMes: 0,
  };
  for (const d of dias) {
    t.trabajado += d.calc.trabajado;
    t.extra1 += d.calc.extra1;
    t.extra2 += d.calc.extra2;
    t.compensado += d.calc.compensado;
    t.enFeriado += d.calc.enFeriado;
    t.contra += d.calc.contra;
    if (d.aplica && (d.tipo === 'LV' || d.tipo === 'SAB')) t.base += baseDiaria(p);
    if (d.calc.pendiente) t.pendientes++;
    if (d.calc.falta) t.faltas++;
    if (d.aplica && d.entrada != null) t.dias++;
  }
  const { netos, nosDebe } = descontarCompensadas(dias.map((d) => d.calc), t.compensado);
  dias.forEach((d, i) => {
    d.neto = netos[i];
    t.neto1 += netos[i].neto1;
    t.neto2 += netos[i].neto2;
  });
  t.nosDebe = nosDebe;
  // El pago se calcula día por día (igual que en el Excel) y se suma.
  const pagoCon = (vh: number) => redondear2(dias.reduce((acc, d) => acc + solesDia(d, vh, r).total, 0));
  t.valorHora = valorHora(p, r, p.valorHoraSegun);
  for (const d of dias) d.soles = solesDia(d, t.valorHora, r);
  t.montoOcho = pagoCon(valorHora(p, r, 'ocho'));
  t.montoReales = pagoCon(valorHora(p, r, 'reales'));
  t.monto = p.valorHoraSegun === 'reales' ? t.montoReales : t.montoOcho;
  t.totalMes = redondear2(p.sueldo + asignacion(p, r) + t.monto);
  return t;
}

/** Redondeo a céntimos como ROUND(x; 2) de Excel (evita errores de coma flotante, p. ej. 1.005 → 1.01). */
export function redondear2(v: number): number {
  return Math.round(Number((v * 100).toPrecision(15))) / 100;
}

/** Pago de un día: cada parte redondeada a céntimos, como las filas "S/" de la hoja Asistencia. */
export function solesDia(d: Pick<Dia, 'neto' | 'calc'>, vh: number, r: Reglas): SolesDia {
  const s1 = redondear2((d.neto.neto1 / 60) * vh * r.recargo1);
  const s2 = redondear2((d.neto.neto2 / 60) * vh * r.recargo2);
  const sf = redondear2((d.calc.enFeriado / 60) * vh * r.recargoFeriado);
  return { s1, s2, sf, total: redondear2(s1 + s2 + sf) };
}

export function asignacion(p: Persona, r: Reglas): number {
  return p.asignacionFamiliar ? r.montoAsignacion : 0;
}

/**
 * Valor hora = (sueldo + asignación familiar) ÷ 30 días ÷ horas por día, redondeado a 2 decimales.
 * Horas por día: 8 (forma "ocho"), o sus horas reales = horas semanales ÷ 6 (forma "reales").
 */
export function valorHora(p: Persona, r: Reglas, segun: Persona['valorHoraSegun']): number {
  const horasDia = segun === 'reales' ? baseDiaria(p) / 60 : 8;
  if (!(p.sueldo > 0) || !(horasDia > 0)) return 0;
  return redondear2((p.sueldo + asignacion(p, r)) / 30 / horasDia);
}

/** Pago sobre horas totales, redondeado una sola vez (referencia; el pago real se suma día por día). */
export function montoSoles(t: Pick<Totales, 'neto1' | 'neto2' | 'enFeriado'>, valorHora: number, r: Reglas): number {
  const v = (t.neto1 / 60) * valorHora * r.recargo1 + (t.neto2 / 60) * valorHora * r.recargo2 + (t.enFeriado / 60) * valorHora * r.recargoFeriado;
  return redondear2(v);
}

/** Horas base por día laborable (lunes a sábado): horas semanales ÷ 6. */
export function baseDiaria(p: Persona): Min {
  return Math.round((p.horasSemanales * 60) / 6);
}

export interface CasoListado {
  alias: string;
  fecha: ISODate;
  tipo: DayType;
  reloj: string;
  caso: Caso;
}

export function listarCasos(mes: Mes): CasoListado[] {
  const out: CasoListado[] = [];
  for (const f of mes.filas) for (const d of f.dias) for (const c of d.casos) out.push({ alias: f.persona.alias, fecha: d.fecha, tipo: d.tipo, reloj: d.reloj, caso: c });
  return out;
}
