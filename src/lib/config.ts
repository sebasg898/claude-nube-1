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
  entrada: Min;
  salidaLV: Min;
  salidaSab: Min;
  /** Las extras de lunes a viernes se cuentan desde esta hora (normalmente = salida L-V). */
  extrasDesdeLV: Min;
  /** Si llega antes de esta hora, el tiempo antes de su entrada cuenta a favor. null = nunca. */
  tempranoAntesDe: Min | null;
  /** Si se le restan tardanzas, salidas antes de hora y permisos. */
  descuenta: boolean;
  /** Días anteriores no se cuentan (persona nueva). */
  fechaIngreso: ISODate | null;
  /** Saldo de la bolsa de horas al iniciar el mes, en minutos. Positivo = le debemos; negativo = nos debe. */
  saldoAnterior: Min;
}

export interface Motivo {
  nombre: string;
  /** Si no hay marcadas ese día, se le resta la jornada de la bolsa. */
  restaJornada: boolean;
  /** Con este motivo igual hay que escribir entrada y salida (p. ej. avisó la salida por WhatsApp). */
  necesitaHoras: boolean;
}

export interface Reglas {
  /** Las extras solo cuentan si sale este número de minutos o más después de su hora (18:20 → 20). */
  margenExtras: Min;
  almuerzoMin: Min;
  /** La tardanza se resta de la bolsa solo si es de estos minutos o más (0 = toda tardanza resta). */
  tardanzaDesde: Min;
  /** Multiplicador para horas trabajadas en domingo o feriado. */
  factorFeriado: number;
}

export interface Config {
  version: 1;
  empresaNombre: string;
  personas: Persona[];
  motivos: Motivo[];
  reglas: Reglas;
  feriados: { fecha: ISODate; nombre: string }[];
}

export const MOTIVOS_BASE: Motivo[] = [
  { nombre: 'WhatsApp', restaJornada: false, necesitaHoras: true },
  { nombre: 'Comisión / reparto', restaJornada: false, necesitaHoras: true },
  { nombre: 'Permiso', restaJornada: true, necesitaHoras: false },
  { nombre: 'Estudia', restaJornada: true, necesitaHoras: false },
  { nombre: 'Falta', restaJornada: false, necesitaHoras: false },
  { nombre: 'Vacaciones', restaJornada: false, necesitaHoras: false },
  { nombre: 'Descanso médico', restaJornada: false, necesitaHoras: false },
  { nombre: 'Descanso', restaJornada: false, necesitaHoras: false },
  { nombre: 'Feriado', restaJornada: false, necesitaHoras: false },
  { nombre: 'Otro', restaJornada: false, necesitaHoras: false },
];

export function configInicial(anio = new Date().getFullYear()): Config {
  return {
    version: 1,
    empresaNombre: 'QALEN',
    personas: [],
    motivos: MOTIVOS_BASE.map((m) => ({ ...m })),
    reglas: { margenExtras: 20, almuerzoMin: 45, tardanzaDesde: 60, factorFeriado: 1 },
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
    salidaLV: 18 * 60,
    salidaSab: 12 * 60,
    extrasDesdeLV: 18 * 60,
    tempranoAntesDe: null,
    descuenta: true,
    fechaIngreso: null,
    saldoAnterior: 0,
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
    if (p.salidaLV <= p.entrada) err.push(`${p.alias}: la salida L-V debe ser después de la entrada.`);
    if (p.salidaSab <= p.entrada) err.push(`${p.alias}: la salida del sábado debe ser después de la entrada.`);
  }
  for (const [a, n] of alias) if (n > 1 && a) err.push(`El nombre corto "${a}" está repetido; cada persona necesita uno distinto.`);
  if (!(c.reglas.margenExtras >= 0 && c.reglas.margenExtras <= 120)) err.push('El margen de extras debe estar entre 0 y 120 minutos.');
  if (!(c.reglas.tardanzaDesde >= 0 && c.reglas.tardanzaDesde <= 600)) err.push('La tardanza mínima debe estar entre 0 y 600 minutos.');
  if (!(c.reglas.factorFeriado >= 0 && c.reglas.factorFeriado <= 3)) err.push('El factor de feriado debe estar entre 0 y 3.');
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
  return {
    version: 1,
    empresaNombre: o.empresaNombre ?? base.empresaNombre,
    personas: o.personas.map((p) => ({ ...plantilla, ...p, id: String(p.id) })),
    motivos: Array.isArray(o.motivos) && o.motivos.length ? o.motivos : base.motivos,
    reglas: { ...base.reglas, ...(o.reglas ?? {}) },
    feriados: Array.isArray(o.feriados) ? o.feriados : base.feriados,
  };
}
