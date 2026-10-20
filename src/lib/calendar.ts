import { addDays, daysBetween, isoFromParts, weekday, type ISODate } from './time';

export type DayType = 'LV' | 'SAB' | 'DOM' | 'FER';

export function dayType(date: ISODate, feriados: Set<ISODate>): DayType {
  if (feriados.has(date)) return 'FER';
  const w = weekday(date);
  if (w === 7) return 'DOM';
  if (w === 6) return 'SAB';
  return 'LV';
}

/** Domingo de Pascua (algoritmo de Meeus/Butcher). */
function easter(y: number): ISODate {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mo = Math.floor((h + l - 7 * m + 114) / 31);
  const da = ((h + l - 7 * m + 114) % 31) + 1;
  return isoFromParts(y, mo, da);
}

/**
 * Feriados nacionales de Perú. Lista de partida: hay que verificarla cada año
 * con el calendario oficial (el gobierno puede agregar feriados o días no laborables).
 */
export function feriadosPeru(y: number): { fecha: ISODate; nombre: string }[] {
  const p = easter(y);
  const f = (mo: number, d: number, nombre: string) => ({ fecha: isoFromParts(y, mo, d), nombre });
  return [
    f(1, 1, 'Año Nuevo'),
    { fecha: addDays(p, -3), nombre: 'Jueves Santo' },
    { fecha: addDays(p, -2), nombre: 'Viernes Santo' },
    f(5, 1, 'Día del Trabajo'),
    f(6, 7, 'Batalla de Arica y Día de la Bandera'),
    f(6, 29, 'San Pedro y San Pablo'),
    f(7, 23, 'Día de la Fuerza Aérea del Perú'),
    f(7, 28, 'Fiestas Patrias'),
    f(7, 29, 'Fiestas Patrias'),
    f(8, 6, 'Batalla de Junín'),
    f(8, 30, 'Santa Rosa de Lima'),
    f(10, 8, 'Combate de Angamos'),
    f(11, 1, 'Todos los Santos'),
    f(12, 8, 'Inmaculada Concepción'),
    f(12, 9, 'Batalla de Ayacucho'),
    f(12, 25, 'Navidad'),
  ];
}

export { daysBetween };
