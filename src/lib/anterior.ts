/**
 * Continuar un Excel ya trabajado: lee lo que Yasmín completó (horas escritas a mano, motivos,
 * sueldos) y lo combina con un reporte nuevo del reloj, para que no tenga que rehacer nada.
 */
import * as XLSX from 'xlsx';
import type { Config, Persona } from './config';
import type { RelojPersona, Reporte } from './parser';
import { daysBetween, ddmm, fmtHHMM, isoFromExcelSerial, parseHHMM, weekday, type ISODate, type Min } from './time';

export interface DiaAnterior {
  /** Marcadas del reloj tal como quedaron en la fila "Reloj". */
  reloj: Min[];
  entrada: Min | null;
  salida: Min | null;
  motivo: string;
}

export interface ExcelAnterior {
  inicio: ISODate;
  fin: ISODate;
  personas: { alias: string; dias: Record<ISODate, DiaAnterior> }[];
  sueldos: { alias: string; sueldo: number | null; asignacion: boolean | null; segun: Persona['valorHoraSegun'] | null }[];
}

/** Lo que se copia de un día: solo lo que ella escribió (horas a mano y motivo). */
export interface Edicion {
  entrada?: Min;
  salida?: Min;
  motivo?: string;
}
/** Ediciones por ID de persona y fecha. */
export type Ediciones = Record<string, Record<ISODate, Edicion>>;

const txt = (v: unknown) => (v == null ? '' : String(v).trim());

function aMinutos(v: unknown): Min | null {
  if (typeof v === 'number' && v >= 0 && v < 1) return Math.round(v * 1440);
  if (typeof v === 'string') return parseHHMM(v);
  return null;
}

/** Lee un Excel generado por esta página (cualquier versión) y ya completado por Yasmín. */
export function leerExcelAnterior(data: ArrayBuffer | Uint8Array): ExcelAnterior {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(data, { type: 'array' });
  } catch {
    throw new Error('No se pudo abrir el Excel anterior.');
  }
  const ws = wb.Sheets['Asistencia'];
  if (!ws) throw new Error('Ese archivo no tiene la hoja "Asistencia". Elige un Excel descargado de esta página.');
  const filas = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: null });

  // Fila de fechas: la que dice "Fecha" en la columna B
  const rFecha = filas.findIndex((f) => txt(f?.[1]) === 'Fecha');
  if (rFecha < 0) throw new Error('No encontré la fila de fechas en la hoja "Asistencia".');
  const colFecha = new Map<number, ISODate>();
  filas[rFecha].forEach((v, c) => {
    if (c >= 2 && typeof v === 'number' && v > 30000 && v < 80000) colFecha.set(c, isoFromExcelSerial(v));
  });
  if (!colFecha.size) throw new Error('La hoja "Asistencia" no tiene fechas.');
  const fechas = [...colFecha.values()].sort();

  // Bloques por persona: fila con nombre en la columna A y "Reloj" en la B
  const personas: ExcelAnterior['personas'] = [];
  for (let r = rFecha + 1; r < filas.length; r++) {
    if (txt(filas[r]?.[1]) !== 'Reloj' || !txt(filas[r]?.[0])) continue;
    const alias = txt(filas[r][0]).toUpperCase();
    const buscar = (etq: string) => {
      for (let k = r + 1; k < Math.min(filas.length, r + 20); k++) {
        const e = txt(filas[k]?.[1]);
        if (e === 'Reloj') return -1;
        if (e === etq) return k;
      }
      return -1;
    };
    const rE = buscar('Entrada'), rS = buscar('Salida'), rM = buscar('Motivo / Obs.');
    if (rE < 0 || rS < 0) continue;
    const dias: Record<ISODate, DiaAnterior> = {};
    for (const [c, fecha] of colFecha) {
      const reloj = (txt(filas[r][c]).match(/\d{2}:\d{2}/g) ?? []).map(parseHHMM).filter((x): x is Min => x != null);
      dias[fecha] = { reloj, entrada: aMinutos(filas[rE][c]), salida: aMinutos(filas[rS][c]), motivo: rM >= 0 ? txt(filas[rM][c]) : '' };
    }
    personas.push({ alias, dias });
  }
  if (!personas.length) throw new Error('No encontré personas en la hoja "Asistencia".');

  // Sueldos del Resumen (solo en versiones que los tienen)
  const sueldos: ExcelAnterior['sueldos'] = [];
  const res = wb.Sheets['Resumen'];
  if (res) {
    const fr = XLSX.utils.sheet_to_json<unknown[]>(res, { header: 1, raw: true, defval: null });
    const iCab = fr.findIndex((f) => f?.some((v) => txt(v) === 'Nombre') && f?.some((v) => txt(v) === 'Sueldo mensual'));
    if (iCab >= 0) {
      const cab = fr[iCab].map(txt);
      const cN = cab.indexOf('Nombre'), cS = cab.indexOf('Sueldo mensual'), cA = cab.indexOf('Asig. familiar'), cV = cab.indexOf('Valor hora según');
      for (let r = iCab + 1; r < fr.length; r++) {
        const alias = txt(fr[r]?.[cN]).toUpperCase();
        if (!alias || alias === 'TOTAL') break;
        const s = fr[r][cS];
        const a = cA >= 0 ? txt(fr[r][cA]) : '';
        const v = cV >= 0 ? txt(fr[r][cV]) : '';
        sueldos.push({
          alias,
          sueldo: typeof s === 'number' && s >= 0 ? s : null,
          asignacion: a === 'Sí' ? true : a === 'No' ? false : null,
          segun: v === 'Horas reales' ? 'reales' : v ? 'ocho' : null,
        });
      }
    }
  }
  return { inicio: fechas[0], fin: fechas[fechas.length - 1], personas, sueldos };
}

export interface Combinado {
  rep: Reporte;
  ediciones: Ediciones;
  avisos: string[];
  /** Cantidad de datos copiados (horas a mano + motivos). */
  copiados: number;
}

/**
 * Junta el reporte nuevo con el Excel anterior:
 *  - Rango final: desde el primer día de uno hasta el último del otro.
 *  - Marcadas: las del reporte nuevo; para días que solo están en el Excel anterior, las de su fila "Reloj".
 *  - Se copian las horas escritas a mano (las que no están en el reloj) y los motivos. Si ella escribió una hora
 *    y ahora el reloj trae otra, se respeta la suya (con aviso).
 * (Los sueldos se copian aparte con aplicarSueldos.)
 */
export function combinarConAnterior(rep: Reporte, ant: ExcelAnterior, cfg: Config): Combinado {
  const avisos: string[] = [];
  const porAlias = new Map(cfg.personas.map((p) => [p.alias.trim().toUpperCase(), p]));
  const inicio = rep.inicio < ant.inicio ? rep.inicio : ant.inicio;
  const fin = rep.fin > ant.fin ? rep.fin : ant.fin;
  const enReporte = new Set(rep.fechas);

  // Personas del reloj (copia) + marcadas de los días que solo están en el Excel anterior
  const personas: RelojPersona[] = rep.personas.map((p) => ({ ...p, marcadas: { ...p.marcadas }, textoOriginal: { ...p.textoOriginal } }));
  const ediciones: Ediciones = {};
  let copiados = 0;
  const noEncontrados: string[] = [];
  for (const pa of ant.personas) {
    const persona = porAlias.get(pa.alias);
    if (!persona) {
      noEncontrados.push(pa.alias);
      continue;
    }
    let rp = personas.find((x) => x.id === persona.id);
    if (!rp) {
      rp = { id: persona.id, nombre: persona.nombreReloj, departamento: '', marcadas: {}, textoOriginal: {} };
      personas.push(rp);
    }
    for (const [fecha, d] of Object.entries(pa.dias)) {
      if (!enReporte.has(fecha) && d.reloj.length) {
        rp.marcadas[fecha] = [...d.reloj];
        rp.textoOriginal[fecha] = d.reloj.map(fmtHHMM).join('');
      }
      const ed: Edicion = {};
      if (d.entrada != null && !d.reloj.includes(d.entrada)) ed.entrada = d.entrada;
      if (d.salida != null && !d.reloj.includes(d.salida)) ed.salida = d.salida;
      if (d.motivo) ed.motivo = d.motivo;
      const n = Object.keys(ed).length;
      if (n) {
        (ediciones[persona.id] ??= {})[fecha] = ed;
        copiados += n;
      }
    }
  }
  if (noEncontrados.length) avisos.push(`En el Excel anterior hay personas que no están en la configuración (${noEncontrados.join(', ')}): sus datos no se copiaron. ¿Les cambiaste el nombre corto?`);

  // Días del rango que no están ni en el reporte ni en el Excel anterior (los domingos no hacen falta)
  const fechas = daysBetween(inicio, fin);
  const cubierto = (f: ISODate) => (f >= rep.inicio && f <= rep.fin) || (f >= ant.inicio && f <= ant.fin);
  const sinDatos = fechas.filter((f) => !cubierto(f) && weekday(f) !== 7);
  if (sinDatos.length) {
    avisos.push(`Del ${ddmm(sinDatos[0])} al ${ddmm(sinDatos[sinDatos.length - 1])} hay ${sinDatos.length} día(s) que no están en el reporte nuevo ni en el Excel anterior. Saldrán en amarillo: saca un reporte del reloj que los incluya o complétalos a mano.`);
  }

  const combinado: Reporte = {
    inicio,
    fin,
    fechas,
    personas,
    avisos: rep.avisos,
    sinDatos,
  };
  return { rep: combinado, ediciones, avisos, copiados };
}

/** Copia a la configuración el sueldo, la asignación y la forma de calcular el valor hora del Resumen anterior. */
export function aplicarSueldos(ant: ExcelAnterior, cfg: Config): { config: Config; avisos: string[] } {
  const porAlias = new Map(cfg.personas.map((p) => [p.alias.trim().toUpperCase(), p]));
  let config = cfg;
  const cambiados: string[] = [];
  for (const s of ant.sueldos) {
    const p = porAlias.get(s.alias);
    if (!p) continue;
    const cambio: Partial<Persona> = {};
    if (s.sueldo != null && s.sueldo > 0 && s.sueldo !== p.sueldo) cambio.sueldo = s.sueldo;
    if (s.asignacion != null && s.asignacion !== p.asignacionFamiliar) cambio.asignacionFamiliar = s.asignacion;
    if (s.segun != null && s.segun !== p.valorHoraSegun) cambio.valorHoraSegun = s.segun;
    if (Object.keys(cambio).length) {
      config = { ...config, personas: config.personas.map((x) => (x.id === p.id ? { ...x, ...cambio } : x)) };
      cambiados.push(p.alias);
    }
  }
  const avisos = cambiados.length ? [`Se tomaron del Excel anterior el sueldo / asignación de: ${cambiados.join(', ')}. Quedan guardados en "Personas y reglas".`] : [];
  return { config, avisos };
}
