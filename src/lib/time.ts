/** Utilidades de hora. Internamente todo se maneja en minutos desde las 00:00. */

export type Min = number;

export function parseHHMM(s: string): Min | null {
  const m = /^\s*(\d{1,2})\s*[:.]\s*(\d{2})\s*$/.exec(s);
  if (!m) return null;
  const h = Number(m[1]);
  const mm = Number(m[2]);
  if (h > 23 || mm > 59) return null;
  return h * 60 + mm;
}

export function fmtHHMM(m: Min | null | undefined): string {
  if (m == null) return '';
  const h = Math.floor(m / 60);
  const mm = m % 60;
  return `${String(h).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

/** Duración en h:mm (admite negativos y más de 24 h). */
export function fmtDur(m: Min): string {
  const sign = m < 0 ? '-' : '';
  const a = Math.abs(Math.round(m));
  return `${sign}${Math.floor(a / 60)}:${String(a % 60).padStart(2, '0')}`;
}

/** Fecha ISO (YYYY-MM-DD) sin zona horaria. */
export type ISODate = string;

export function isoFromParts(y: number, mo: number, d: number): ISODate {
  return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

export function parseISO(s: ISODate): { y: number; mo: number; d: number } {
  const [y, mo, d] = s.split('-').map(Number);
  return { y, mo, d };
}

export function addDays(s: ISODate, n: number): ISODate {
  const { y, mo, d } = parseISO(s);
  const t = new Date(Date.UTC(y, mo - 1, d + n));
  return isoFromParts(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
}

/** 1 = lunes … 7 = domingo. */
export function weekday(s: ISODate): number {
  const { y, mo, d } = parseISO(s);
  const w = new Date(Date.UTC(y, mo - 1, d)).getUTCDay();
  return w === 0 ? 7 : w;
}

/** Número de serie de Excel (sistema 1900) para una fecha. */
export function excelSerial(s: ISODate): number {
  const { y, mo, d } = parseISO(s);
  return Math.round((Date.UTC(y, mo - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}

export function isoFromExcelSerial(n: number): ISODate {
  const t = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  return isoFromParts(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
}

export function daysBetween(a: ISODate, b: ISODate): ISODate[] {
  const out: ISODate[] = [];
  for (let d = a; d <= b; d = addDays(d, 1)) out.push(d);
  return out;
}

export const DIAS = ['', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
export const DIAS_CORTO = ['', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];
export const MESES = ['', 'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SETIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];

export function ddmm(s: ISODate): string {
  const { mo, d } = parseISO(s);
  return `${String(d).padStart(2, '0')}/${String(mo).padStart(2, '0')}`;
}
