import { feriadosPeru } from './calendar';
import type { ISODate, Min } from './time';
import type { RelojPersona } from './parser';

export interface Persona {
  /** ID del reloj (clave). */
  id: string;
  nombreReloj: string;
  /** Nombre corto que usa Yasmín (p. ej. "ANA", "SR LUIS"). Debe ser único. */
  alias: string;
  empresa: string;
  activo: boolean;
  /** Transportista / reparto: si no marca salida se le sugiere la hora de salida, pero siempre se pregunta. */
  reparto: boolean;
  horasSemanales: number;
  /** Entrada de lunes a viernes. */
  entrada: Min;
  /** Entrada del sábado (puede ser distinta, p. ej. turno de tarde en semana). */
  entradaSab: Min;
  salidaLV: Min;
  salidaSab: Min;
  /** Las extras de lunes a viernes se cuentan desde esta hora (normalmente = salida L-V). */
  extrasDesdeLV: Min;
  /** Si llega antes de esta hora, el tiempo antes de su entrada cuenta como extra. null = nunca (regla actual). */
  tempranoAntesDe: Min | null;
  /** Sueldo mensual en soles. */
  sueldo: number;
  /** Recibe asignación familiar (se suma al sueldo para el valor hora). */
  asignacionFamiliar: boolean;
  /** Cómo se calcula el valor hora: 30 días × 8 h, o 30 días × sus horas reales por día (horas semanales ÷ 6). */
  valorHoraSegun: 'ocho' | 'reales';
  /** Si se le restan tardanzas, salidas antes de hora y permisos. */
  descuenta: boolean;
  /** Días anteriores no se cuentan (persona nueva). */
  fechaIngreso: ISODate | null;
}

/**
 * Motivos de la fila "Motivo" (lista cerrada):
 *  - Vacaciones / Salud: se considera que trabajó el día completo.
 *  - Compensado: canjea horas a favor; lo que falte del día va a "Compensadas" (no a "En contra").
 *  - Estudio: lo que falte del día (o la jornada completa si no marcó) es hora en contra.
 *  - Falta: la jornada completa es hora en contra y se cuenta como falta.
 */
export const MOTIVOS = ['Vacaciones', 'Salud', 'Compensado', 'Estudio', 'Falta'] as const;

export interface Reglas {
  /** Las extras solo cuentan si sale este número de minutos o más después de su hora (30 → desde las 18:30). */
  margenExtras: Min;
  almuerzoMin: Min;
  /** La tardanza se resta de la bolsa solo si es de estos minutos o más (0 = toda tardanza resta). */
  tardanzaDesde: Min;
  /** Horas extras del día que van al primer recargo (120 = las 2 primeras horas). */
  tramo1Min: Min;
  /** Multiplicadores de pago sobre el valor hora (ley: mínimo 1.25, 1.35 y 2.00). */
  recargo1: number;
  recargo2: number;
  recargoFeriado: number;
  /** Monto de la asignación familiar (10% de la remuneración mínima vital). */
  montoAsignacion: number;
}

export interface Config {
  version: 4;
  empresaNombre: string;
  personas: Persona[];
  reglas: Reglas;
  feriados: { fecha: ISODate; nombre: string }[];
}


export function configInicial(anio = new Date().getFullYear()): Config {
  return {
    version: 4,
    empresaNombre: 'QALEN',
    personas: [],
    reglas: { montoAsignacion: 113, margenExtras: 30, almuerzoMin: 45, tardanzaDesde: 0, tramo1Min: 120, recargo1: 1.25, recargo2: 1.35, recargoFeriado: 2 },
    feriados: [...feriadosPeru(anio), ...feriadosPeru(anio + 1)],
  };
}

export function personaNueva(r: Pick<RelojPersona, 'id' | 'nombre'>, aliasUsados: Set<string>, activo: boolean): Persona {
  const base = (r.nombre.trim().split(/\s+/)[0] || `ID ${r.id}`).toUpperCase();
  let alias = base;
  for (let i = 2; aliasUsados.has(alias); i++) alias = `${base} ${i}`;
  aliasUsados.add(alias);
  return {
    id: r.id,
    nombreReloj: r.nombre,
    alias,
    empresa: '',
    activo,
    reparto: false,
    horasSemanales: 48,
    entrada: 8 * 60 + 30,
    entradaSab: 8 * 60 + 30,
    salidaLV: 18 * 60,
    salidaSab: 12 * 60,
    extrasDesdeLV: 18 * 60,
    tempranoAntesDe: null,
    sueldo: 0,
    asignacionFamiliar: false,
    valorHoraSegun: 'ocho',
    descuenta: true,
    fechaIngreso: null,
  };
}

/** Revisa la configuración y devuelve problemas que impiden generar el Excel. */
export function validarConfig(c: Config): string[] {
  const err: string[] = [];
  const alias = new Map<string, number>();
  for (const p of c.personas.filter((x) => x.activo)) {
    const a = p.alias.trim().toUpperCase();
    if (!a) err.push(`La persona con ID ${p.id} no tiene nombre corto.`);
    alias.set(a, (alias.get(a) ?? 0) + 1);
    if (!(p.horasSemanales > 0 && p.horasSemanales <= 72)) err.push(`${p.alias}: las horas semanales deben estar entre 1 y 72.`);
    if (p.salidaLV <= p.entrada) err.push(`${p.alias}: la salida L-V debe ser después de la entrada L-V.`);
    if (p.salidaSab <= p.entradaSab) err.push(`${p.alias}: la salida del sábado debe ser después de la entrada del sábado.`);
    if (p.extrasDesdeLV < p.salidaLV) err.push(`${p.alias}: "Extras desde" no puede ser antes de su salida L-V.`);
    if (!(p.sueldo >= 0 && p.sueldo < 100000)) err.push(`${p.alias}: el sueldo no es válido.`);
  }
  for (const [a, n] of alias) if (n > 1 && a) err.push(`El nombre corto "${a}" está repetido; cada persona necesita uno distinto.`);
  if (!(c.reglas.margenExtras >= 0 && c.reglas.margenExtras <= 120)) err.push('El margen de extras debe estar entre 0 y 120 minutos.');
  if (!(c.reglas.tardanzaDesde >= 0 && c.reglas.tardanzaDesde <= 600)) err.push('La tardanza mínima debe estar entre 0 y 600 minutos.');
  for (const [k, n] of [['recargo1', 'primer recargo'], ['recargo2', 'segundo recargo'], ['recargoFeriado', 'recargo de domingo/feriado']] as const) {
    if (!(c.reglas[k] >= 1 && c.reglas[k] <= 4)) err.push(`El ${n} debe estar entre 1 y 4 (por ejemplo 1.25).`);
  }
  if (!(c.reglas.tramo1Min >= 0 && c.reglas.tramo1Min <= 600)) err.push('Las horas del primer recargo deben estar entre 0 y 10.');
  if (!(c.reglas.montoAsignacion >= 0 && c.reglas.montoAsignacion < 10000)) err.push('El monto de la asignación familiar no es válido.');
  return err;
}

const CLAVE = 'qalen-asistencia-config-v1';

export function cargarConfig(): Config | null {
  try {
    const raw = localStorage.getItem(CLAVE);
    return raw ? normalizarConfig(JSON.parse(raw)) : null;
  } catch {
    return null;
  }
}

export function guardarConfig(c: Config): void {
  try {
    localStorage.setItem(CLAVE, JSON.stringify(c));
  } catch {
    /* sin almacenamiento: se trabaja solo en memoria */
  }
}

/** Acepta un respaldo JSON (posiblemente de una versión anterior) y lo completa con valores por defecto. */
export function normalizarConfig(x: unknown): Config {
  const base = configInicial();
  if (!x || typeof x !== 'object') throw new Error('El respaldo no es válido.');
  const o = x as Partial<Config>;
  if (!Array.isArray(o.personas)) throw new Error('El respaldo no tiene la lista de personas.');
  const plantilla = personaNueva({ id: '0', nombre: 'X' }, new Set(), true);
  // Migración de respaldos antiguos (versión 1 o 2).
  const ver = (o as { version?: number }).version ?? 1;
  return {
    version: 4,
    empresaNombre: o.empresaNombre ?? base.empresaNombre,
    personas: o.personas.map((p) => {
      const { saldoAnterior: _s, valorHora, ...resto } = p as Persona & { saldoAnterior?: number; valorHora?: number };
      void _s;
      const migr: Partial<Persona> = {};
      if (ver < 3) migr.tempranoAntesDe = null;
      // Versión < 4: antes se escribía el valor hora; ahora se escribe el sueldo (valor hora × 240).
      if (ver < 4 && !resto.sueldo && valorHora) migr.sueldo = Math.round(valorHora * 240 * 100) / 100;
      return { ...plantilla, ...resto, id: String(p.id), ...migr };
    }),
    reglas: (() => {
      // Versión < 3: llegar temprano deja de sumar; extras con recargo. Versión < 4: margen de extras 30 min.
      const { factorFeriado: _f, ...r } = { ...base.reglas, ...(o.reglas ?? {}), ...(ver < 3 ? { tardanzaDesde: 0 } : {}), ...(ver < 4 ? { margenExtras: 30 } : {}) } as Reglas & { factorFeriado?: number };
      void _f;
      return r;
    })(),
    feriados: Array.isArray(o.feriados) ? o.feriados : base.feriados,
  };
}
