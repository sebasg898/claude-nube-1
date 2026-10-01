import { corteSugerido, dayType, type DayType } from './calendar';
import { personaNueva, type Config, type Persona } from './config';
import type { Reporte } from './parser';
import { calcularDia, clasificar, type Caso, type Clasificacion, type DiaCalculo } from './rules';
import { daysBetween, fmtHHMM, weekday, type ISODate, type Min } from './time';

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
  casos: Caso[];
}

export interface FilaPersona {
  persona: Persona;
  dias: Dia[];
  tot: { trabajado: Min; favor: Min; contra: Min; enFeriado: Min; pendientes: number; faltas: number; dias: number; saldoFinal: Min };
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

export function corteMes(rep: Reporte, cfg: Config): ISODate {
  return corteSugerido(rep.inicio, rep.fin, new Set(cfg.feriados.map((f) => f.fecha)));
}

export function construirMes(rep: Reporte, cfg: Config, hasta: ISODate): Mes {
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
      const entrada = aplica ? clas.entrada : null;
      const salida = aplica ? clas.salida : null;
      const calc = calcularDia({ entrada, salida, motivo: '' }, tipo, p, cfg.reglas, cfg.motivos, aplica);
      const casos = [...clas.casos];
      if (calc.favor - calc.enFeriado > 180) casos.push({ tipo: 'extras_altas', faltaDato: false, mensaje: `Más de 3 horas a favor en el día (${Math.floor(calc.favor / 60)} h ${calc.favor % 60} min). Solo para revisar.` });
      return { fecha, tipo, aplica, reloj: clas.marcas.map(fmtHHMM).join(' '), clas, entrada, salida, motivo: '', calc, casos };
    });
    return { persona: p, dias, tot: totales(p, dias) };
  });

  const avisos: string[] = [];
  if (fin < rep.fin) avisos.push(`Se calcula hasta el ${fin}; el reporte trae marcadas hasta el ${rep.fin}.`);
  return { inicio: rep.inicio, hasta: fin, finReporte: rep.fin, fechas, filas, avisos };
}

export function totales(p: Persona, dias: Dia[]): FilaPersona['tot'] {
  const t = { trabajado: 0, favor: 0, contra: 0, enFeriado: 0, pendientes: 0, faltas: 0, dias: 0, saldoFinal: 0 };
  for (const d of dias) {
    t.trabajado += d.calc.trabajado;
    t.favor += d.calc.favor;
    t.contra += d.calc.contra;
    t.enFeriado += d.calc.enFeriado;
    if (d.calc.pendiente) t.pendientes++;
    if (d.calc.falta) t.faltas++;
    if (d.aplica && d.entrada != null) t.dias++;
  }
  t.saldoFinal = p.saldoAnterior + t.favor - t.contra;
  return t;
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
