import type { DayType } from './calendar';
import type { Persona, Reglas } from './config';
import { fmtHHMM, type Min } from './time';

const H11 = 11 * 60;
const H12 = 12 * 60;
const H1530 = 15 * 60 + 30;
const H0630 = 6 * 60 + 30;
const H22 = 22 * 60;

export type TipoCaso =
  | 'sin_marcadas'
  | 'sin_entrada'
  | 'sin_salida'
  | 'una_marcada'
  | 'hora_rara'
  | 'extras_altas'
  | 'domingo_feriado';

export interface Caso {
  tipo: TipoCaso;
  /** true = falta un dato y la celda queda amarilla. false = solo hay que mirarlo. */
  faltaDato: boolean;
  mensaje: string;
  sugerencia?: string;
}

export interface Clasificacion {
  marcas: Min[];
  duplicadas: Min[];
  entrada: Min | null;
  salida: Min | null;
  almuerzoSale: Min | null;
  almuerzoVuelve: Min | null;
  otras: Min[];
  casos: Caso[];
}

/** Quita marcadas repetidas (a menos de 3 min de la anterior que se conservó). */
export function deduplicar(ms: Min[]): { marcas: Min[]; duplicadas: Min[] } {
  const orden = [...ms].sort((a, b) => a - b);
  const marcas: Min[] = [];
  const duplicadas: Min[] = [];
  for (const m of orden) {
    if (marcas.length && m - marcas[marcas.length - 1] < 3) duplicadas.push(m);
    else marcas.push(m);
  }
  return { marcas, duplicadas };
}

export function limiteSalidaLV(p: Persona): Min {
  return Math.max(H11, p.salidaLV - 90);
}

/** Decide cuál marcada es la entrada y cuál la salida. No adivina: lo dudoso queda como caso. */
export function clasificar(raw: Min[], tipo: DayType, p: Persona, aplica: boolean): Clasificacion {
  const { marcas, duplicadas } = deduplicar(raw);
  const c: Clasificacion = { marcas, duplicadas, entrada: null, salida: null, almuerzoSale: null, almuerzoVuelve: null, otras: [], casos: [] };
  const lista = marcas.map(fmtHHMM).join(', ');
  const laborable = tipo === 'LV' || tipo === 'SAB';

  if (!marcas.length) {
    if (laborable && aplica) {
      c.casos.push({ tipo: 'sin_marcadas', faltaDato: true, mensaje: 'No marcó en todo el día.', sugerencia: 'Falta' });
    }
    return c;
  }

  if (!laborable) {
    c.entrada = marcas[0];
    c.salida = marcas.length > 1 ? marcas[marcas.length - 1] : null;
    c.otras = marcas.slice(1, -1);
    c.casos.push({
      tipo: 'domingo_feriado',
      faltaDato: c.salida == null,
      mensaje: `Trabajó en ${tipo === 'DOM' ? 'domingo' : 'feriado'} (${lista}).${c.salida == null ? ' Solo hay una marcada: falta la salida.' : ''}`,
    });
    return c;
  }

  const primera = marcas[0];
  const ultima = marcas[marcas.length - 1];
  const limite = tipo === 'SAB' ? H11 : limiteSalidaLV(p);
  const horaFin = tipo === 'SAB' ? p.salidaSab : p.salidaLV;

  if (primera < H11) c.entrada = primera;
  if (ultima >= limite && ultima > H11 && !(marcas.length === 1 && c.entrada != null)) c.salida = ultima;

  // Almuerzo (solo informativo) y otras marcadas
  const medio = marcas.filter((m) => m !== c.entrada && m !== c.salida);
  if (tipo === 'LV') {
    const enRango = medio.filter((m) => m >= H12 && m <= H1530);
    if (enRango.length) {
      c.almuerzoSale = enRango[0];
      const vuelve = enRango.find((m) => m - enRango[0] >= 20 && m - enRango[0] <= 90);
      c.almuerzoVuelve = vuelve ?? null;
    }
  }
  c.otras = medio.filter((m) => m !== c.almuerzoSale && m !== c.almuerzoVuelve);

  if (marcas.length === 1) {
    c.casos.push({
      tipo: 'una_marcada',
      faltaDato: true,
      mensaje: `Solo marcó una vez (${lista}).${c.entrada != null ? ' Se tomó como entrada; falta la salida.' : c.salida != null ? ' Se tomó como salida; falta la entrada.' : ' No se sabe si es entrada o salida.'}`,
      sugerencia: c.entrada != null && p.reparto ? fmtHHMM(horaFin) : undefined,
    });
  } else {
    if (c.entrada == null) {
      c.casos.push({ tipo: 'sin_entrada', faltaDato: true, mensaje: `No marcó entrada: la primera marcada es a las ${fmtHHMM(primera)}.` });
    }
    if (c.salida == null) {
      const msg = tipo === 'LV' && ultima >= H11
        ? `No marcó salida: la última marcada es a las ${fmtHHMM(ultima)}. ¿Es el almuerzo y luego no marcó, o salió a esa hora?`
        : `No marcó salida (marcadas: ${lista}).`;
      c.casos.push({
        tipo: 'sin_salida',
        faltaDato: true,
        mensaje: msg + (p.reparto ? ' Es de reparto: ¿avisó su salida por WhatsApp?' : ''),
        sugerencia: p.reparto ? fmtHHMM(horaFin) : undefined,
      });
    }
  }

  const raras = marcas.filter((m) => m < H0630 || m > H22);
  if (raras.length) {
    c.casos.push({ tipo: 'hora_rara', faltaDato: false, mensaje: `Marcada a una hora poco común: ${raras.map(fmtHHMM).join(', ')}. Es válida, pero conviene revisarla.` });
  }
  return c;
}

export interface DiaInput {
  entrada: Min | null;
  salida: Min | null;
  motivo: string;
}

export interface DiaCalculo {
  trabajado: Min;
  /** Horas extras del día (lunes a sábado) en el primer tramo (25%). */
  extra1: Min;
  /** Horas extras del día desde la 3.ª hora (35%). */
  extra2: Min;
  /** Horas trabajadas en domingo o feriado (100%). */
  enFeriado: Min;
  /** Horas en contra (tardanza, salida antes de hora, Estudio, Falta). Van aparte. */
  contra: Min;
  /** Horas canjeadas (motivo "Compensado"). Se descuentan de las extras de los primeros días. */
  compensado: Min;
  /** Falta un dato (celda amarilla). */
  pendiente: boolean;
  /** Salida igual o anterior a la entrada (casi siempre: hora escrita en formato de 12 h). */
  alerta: boolean;
  falta: boolean;
}

/** Jornada de un día completo (para Falta, Estudio o Compensado sin marcadas). */
export function jornada(p: Persona, tipo: DayType): Min {
  const fin = tipo === 'SAB' ? p.salidaSab : p.salidaLV;
  return Math.min(Math.round((p.horasSemanales * 60) / 6), fin - p.entrada);
}

/** Horas trabajadas de un día normal completo (Vacaciones / Salud). */
export function diaCompleto(p: Persona, tipo: DayType, r: Reglas): Min {
  const fin = tipo === 'SAB' ? p.salidaSab : p.salidaLV;
  return fin - p.entrada - (tipo === 'LV' ? r.almuerzoMin : 0);
}

/**
 * Cálculo de referencia de un día. Es exactamente lo que hacen las fórmulas del Excel
 * (ver excel.ts); la prueba de coherencia compara ambos.
 */
export function calcularDia(d: DiaInput, tipo: DayType, p: Persona, r: Reglas, aplica: boolean): DiaCalculo {
  const z: DiaCalculo = { trabajado: 0, extra1: 0, extra2: 0, enFeriado: 0, contra: 0, compensado: 0, pendiente: false, alerta: false, falta: false };
  if (!aplica) return z;
  const m = d.motivo.trim().toLowerCase();
  const e = d.entrada;
  const s = d.salida;
  const ambos = e != null && s != null;
  const xor = (e == null) !== (s == null);
  z.alerta = ambos && s! <= e!;
  z.falta = m === 'falta';
  const laborable = tipo === 'LV' || tipo === 'SAB';

  if (ambos && s! > e!) {
    const bruto = s! - e!;
    z.trabajado = bruto - (tipo === 'LV' && bruto > 300 ? r.almuerzoMin : 0);
  }
  if (!laborable) {
    z.pendiente = xor;
    if (ambos && s! > e!) z.enFeriado = s! - e!;
    return z;
  }

  if (m === '') z.pendiente = e == null || s == null;
  else if (m === 'compensado' || m === 'estudio') z.pendiente = xor;

  if (m === 'vacaciones' || m === 'salud') {
    z.trabajado = diaCompleto(p, tipo, r);
    return z;
  }

  const fin = tipo === 'SAB' ? p.salidaSab : p.salidaLV;
  const extrasDesde = tipo === 'SAB' ? p.salidaSab : p.extrasDesdeLV;
  if (m !== 'falta') {
    let ext = 0;
    if (s != null && s >= extrasDesde + r.margenExtras) ext += s - extrasDesde;
    if (e != null && p.tempranoAntesDe != null && e < p.tempranoAntesDe && e < p.entrada) ext += p.entrada - e;
    z.extra1 = Math.min(ext, r.tramo1Min);
    z.extra2 = ext - z.extra1;
  }

  if (p.descuenta) {
    let faltan = 0;
    if (e != null && e - p.entrada >= Math.max(r.tardanzaDesde, 1)) faltan += e - p.entrada;
    if (s != null && s < fin) faltan += fin - s;
    const sinMarcas = e == null && s == null;
    if (m === 'falta') z.contra = jornada(p, tipo);
    else if (m === 'compensado') z.compensado = sinMarcas ? jornada(p, tipo) : faltan;
    else if (m === 'estudio') z.contra = sinMarcas ? jornada(p, tipo) : faltan;
    else z.contra = faltan;
  }
  return z;
}

export interface Neto {
  /** Extras que quedan para pagar después de descontar las compensadas. */
  neto1: Min;
  neto2: Min;
}

/**
 * Descuenta las horas compensadas del período de las extras de los primeros días
 * (en orden de fecha y, dentro de cada día, primero las del 35%).
 * Devuelve lo que queda por día y las compensadas que no alcanzaron a cubrirse (nos debe).
 */
export function descontarCompensadas(dias: { extra1: Min; extra2: Min }[], compensadas: Min): { netos: Neto[]; nosDebe: Min } {
  let resto = compensadas;
  const netos = dias.map((d) => {
    const u2 = Math.min(d.extra2, resto);
    resto -= u2;
    const u1 = Math.min(d.extra1, resto);
    resto -= u1;
    return { neto1: d.extra1 - u1, neto2: d.extra2 - u2 };
  });
  return { netos, nosDebe: resto };
}
