import * as XLSX from 'xlsx';
import { addDays, isoFromParts, parseHHMM, parseISO, type ISODate, type Min } from './time';

export interface RelojPersona {
  id: string;
  nombre: string;
  departamento: string;
  /** Marcadas por fecha, tal como vienen del reloj (ordenadas, sin deduplicar). */
  marcadas: Record<ISODate, Min[]>;
  /** Texto original de cada celda (para mostrar "el reporte de la máquina"). */
  textoOriginal: Record<ISODate, string>;
}

export interface Reporte {
  inicio: ISODate;
  fin: ISODate;
  fechas: ISODate[];
  personas: RelojPersona[];
  avisos: string[];
}

export class ErrorReporte extends Error {}

const HOJA = 'reporte de asistencia';

function txt(v: unknown): string {
  return v == null ? '' : String(v).trim();
}

/** Lee el .xls del reloj facial (hoja "Reporte de Asistencia"). */
export function leerReporte(data: ArrayBuffer | Uint8Array): Reporte {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(data, { type: 'array', cellDates: false });
  } catch {
    throw new ErrorReporte('No se pudo abrir el archivo. ¿Es el reporte exportado del reloj (.xls)?');
  }
  const nombreHoja = wb.SheetNames.find((n) => n.trim().toLowerCase() === HOJA)
    ?? wb.SheetNames.find((n) => n.toLowerCase().includes('asistencia') && !n.toLowerCase().includes('turno'));
  if (!nombreHoja) {
    throw new ErrorReporte(`El archivo no tiene la hoja "Reporte de Asistencia". Hojas encontradas: ${wb.SheetNames.join(', ')}.`);
  }
  const filas = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[nombreHoja], { header: 1, raw: true, defval: '' });
  return interpretarFilas(filas);
}

export function interpretarFilas(filas: unknown[][]): Reporte {
  const avisos: string[] = [];

  // 1) Período
  let inicio: ISODate | null = null;
  let fin: ISODate | null = null;
  let filaPeriodo = -1;
  for (let r = 0; r < Math.min(filas.length, 15) && !inicio; r++) {
    for (const v of filas[r]) {
      const m = /(\d{4})-(\d{2})-(\d{2})\s*~\s*(\d{4})-(\d{2})-(\d{2})/.exec(txt(v));
      if (m) {
        inicio = isoFromParts(+m[1], +m[2], +m[3]);
        fin = isoFromParts(+m[4], +m[5], +m[6]);
        filaPeriodo = r;
        break;
      }
    }
  }
  if (!inicio || !fin) throw new ErrorReporte('No encontré la línea "Periodo: AAAA-MM-DD ~ AAAA-MM-DD" en el reporte.');
  if (fin < inicio) throw new ErrorReporte('El período del reporte está al revés (la fecha final es anterior a la inicial).');

  // 2) Fila con los números de día
  let filaDias = -1;
  for (let r = filaPeriodo + 1; r < Math.min(filas.length, filaPeriodo + 6); r++) {
    const nums = filas[r].filter((v) => txt(v) !== '' && /^\d{1,2}(\.0)?$/.test(txt(v)));
    if (nums.length >= 1 && nums.length === filas[r].filter((v) => txt(v) !== '').length) {
      filaDias = r;
      break;
    }
  }
  if (filaDias < 0) throw new ErrorReporte('No encontré la fila con los números de día (1, 2, 3…).');

  // Columna -> fecha. Recorre el período en orden; admite períodos que cruzan de mes.
  const colFecha = new Map<number, ISODate>();
  const fechas: ISODate[] = [];
  {
    let cursor = inicio;
    filas[filaDias].forEach((v, c) => {
      const t = txt(v);
      if (!t) return;
      const d = Math.round(Number(t));
      // avanzar el cursor hasta la fecha con ese número de día
      let guard = 0;
      while (parseISO(cursor).d !== d && guard < 40) {
        cursor = addDays(cursor, 1);
        guard++;
      }
      if (guard >= 40 || cursor > fin!) {
        avisos.push(`La columna del día ${d} no calza con el período; se ignoró.`);
        return;
      }
      colFecha.set(c, cursor);
      fechas.push(cursor);
      cursor = addDays(cursor, 1);
    });
  }

  // 3) Personas
  const personas: RelojPersona[] = [];
  const idsVistos = new Set<string>();
  for (let r = filaDias + 1; r < filas.length; r++) {
    const fila = filas[r];
    const idx = fila.findIndex((v) => /^ID\s*:?$/i.test(txt(v)));
    if (idx < 0) continue;
    const valorTras = (etiqueta: RegExp) => {
      const i = fila.findIndex((v) => etiqueta.test(txt(v)));
      if (i < 0) return '';
      for (let j = i + 1; j < fila.length; j++) {
        const t = txt(fila[j]);
        if (t && !/:$/.test(t)) return t;
        if (/:$/.test(t)) return '';
      }
      return '';
    };
    const id = valorTras(/^ID\s*:?$/i).replace(/\.0$/, '');
    const nombre = valorTras(/^Nombre\s*:?$/i);
    const departamento = valorTras(/^Departamento\s*:?$/i);
    if (!id) {
      avisos.push(`Fila ${r + 1}: persona sin ID; se ignoró.`);
      continue;
    }
    if (idsVistos.has(id)) avisos.push(`El ID ${id} aparece dos veces en el reporte; se juntaron sus marcadas.`);
    idsVistos.add(id);

    const filaM = filas[r + 1] ?? [];
    const marcadas: Record<ISODate, Min[]> = {};
    const textoOriginal: Record<ISODate, string> = {};
    for (const [c, fecha] of colFecha) {
      const t = txt(filaM[c]);
      if (!t) continue;
      textoOriginal[fecha] = t;
      const horas = t.match(/\d{2}:\d{2}/g) ?? [];
      const sobrante = t.replace(/\d{2}:\d{2}/g, '').replace(/\s/g, '');
      if (sobrante) avisos.push(`${nombre || id}, ${fecha}: texto raro en la celda ("${t}"); se tomaron solo las horas válidas.`);
      const mins = horas.map(parseHHMM).filter((x): x is Min => x != null);
      if (mins.length !== horas.length) avisos.push(`${nombre || id}, ${fecha}: hora imposible en "${t}"; se ignoró.`);
      if (mins.length) marcadas[fecha] = mins.sort((a, b) => a - b);
    }

    const existente = personas.find((p) => p.id === id);
    if (existente) {
      for (const [f, ms] of Object.entries(marcadas)) existente.marcadas[f] = [...(existente.marcadas[f] ?? []), ...ms].sort((a, b) => a - b);
      Object.assign(existente.textoOriginal, textoOriginal);
    } else {
      personas.push({ id, nombre, departamento, marcadas, textoOriginal });
    }
    r++; // la fila siguiente era la de marcadas
  }
  if (!personas.length) throw new ErrorReporte('El reporte no tiene personas (no encontré filas con "ID:").');

  return { inicio, fin, fechas, personas, avisos };
}
