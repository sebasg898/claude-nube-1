import type { DayType } from './calendar';
import type { Motivo, Persona, Reglas } from './config';
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
  favor: Min;
  contra: Min;
  /** Parte de "favor" que es por trabajar en domingo/feriado. */
  enFeriado: Min;
  /** Falta un dato (celda amarilla). */
  pendiente: boolean;
  /** Salida igual o anterior a la entrada (casi siempre: hora escrita en formato de 12 h). */
  alerta: boolean;
  falta: boolean;
}

/**
 * Cálculo de referencia de un día. Es exactamente lo que hacen las fórmulas del Excel
 * (ver excel.ts); la prueba de coherencia compara ambos.
 */
export function calcularDia(d: DiaInput, tipo: DayType, p: Persona, r: Reglas, motivos: Motivo[], aplica: boolean): DiaCalculo {
  const z: DiaCalculo = { trabajado: 0, favor: 0, contra: 0, enFeriado: 0, pendiente: false, alerta: false, falta: false };
  if (!aplica) return z;
  const m = motivos.find((x) => x.nombre.toLowerCase() === d.motivo.trim().toLowerCase());
  const e = d.entrada;
  const s = d.salida;
  const ambos = e != null && s != null;
  z.alerta = ambos && s! <= e!;
  z.falta = d.motivo.trim().toLowerCase() === 'falta';
  const laborable = tipo === 'LV' || tipo === 'SAB';

  if (ambos && s! > e!) {
    const bruto = s! - e!;
    z.trabajado = bruto - (tipo === 'LV' && bruto > 300 ? r.almuerzoMin : 0);
  }
  if (!laborable) {
    // Domingo/feriado: si solo hay una de las dos horas, falta un dato.
    z.pendiente = (e == null) !== (s == null);
    if (ambos && s! > e!) {
      z.favor = Math.round((s! - e!) * r.factorFeriado);
      z.enFeriado = z.favor;
    }
    return z;
  }

  const sinAlgo = e == null || s == null;
  z.pendiente = d.motivo.trim() === '' ? sinAlgo : !!m?.necesitaHoras && sinAlgo;

  const fin = tipo === 'SAB' ? p.salidaSab : p.salidaLV;
  const extrasDesde = tipo === 'SAB' ? p.salidaSab : p.extrasDesdeLV;
  if (s != null && s >= extrasDesde + r.margenExtras) z.favor += s - extrasDesde;
  if (e != null && p.tempranoAntesDe != null && e < p.tempranoAntesDe && e < p.entrada) z.favor += p.entrada - e;

  if (p.descuenta) {
    if (e != null && e > p.entrada && e - p.entrada >= r.tardanzaDesde) z.contra += e - p.entrada;
    if (s != null && s < fin) z.contra += fin - s;
    if (m?.restaJornada && e == null && s == null) z.contra += Math.min(Math.round((p.horasSemanales * 60) / 6), fin - p.entrada);
  }
  return z;
}
