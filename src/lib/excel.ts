/**
 * Genera el Excel de trabajo de Yasmín (4 hojas: Leeme, Asistencia, Semanal, Resumen).
 *
 * Reparto del trabajo:
 *  - La página decide la ENTRADA y la SALIDA de cada día (o las deja vacías si hay duda).
 *  - Todo lo demás son fórmulas simples dentro del Excel: si Yasmín cambia una hora o
 *    elige un motivo, se recalculan solos los totales, la hoja Semanal y el Resumen.
 *
 * Los parámetros de cada persona (horario, etc.) van como valores en columnas ocultas de
 * la hoja Asistencia; no hay hoja de configuración que ella tenga que tocar. El valor hora
 * está a la vista en el Resumen y se puede corregir ahí.
 * Las fórmulas replican exactamente calcularDia(), descontarCompensadas() y montoSoles()
 * (hay una prueba que lo verifica recalculando el archivo con LibreOffice).
 */
import ExcelJS from 'exceljs';
import { MOTIVOS, type Config } from './config';
import { baseDiaria, type Mes } from './model';
import type { Reporte } from './parser';
import { DIAS_CORTO, MESES, ddmm, excelSerial, fmtHHMM, parseISO, weekday } from './time';

// ---------- estilo ----------
const ROJO = 'FFC8102E';
const GRIS = 'FFF2F2F2';
const GRIS_OSC = 'FF7F7F7F';
const GRIS_RELOJ = 'FFE7E6E6';
const AMARILLO = 'FFFFEB3B';
const AZUL = 'FFDDEBF7';
const VERDE = 'FFE2F0D9';
const VERDE_ENTRADA = 'FFE2EFDA';
const NARANJA_SALIDA = 'FFFCE4D6';
const ROSA = 'FFFFC7CE';
const MORADO = 'FF7030A0';
const VERDE_TXT = 'FF375623';
const VERDE_OSC = 'FF1E5631';
const ROJO_TXT = 'FF9C0006';
const AZUL_FER = 'FF1F4E79';

const fill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const cfFill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', bgColor: { argb } });
const thin: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const grueso: Partial<ExcelJS.Border> = { style: 'medium', color: { argb: 'FF595959' } };
const bordes: Partial<ExcelJS.Borders> = { top: thin, left: thin, bottom: thin, right: thin };
const SOLES = '"S/ "#,##0.00';

export function col(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ---------- disposición ----------
export const HOJAS = { leeme: 'Leeme', asis: 'Asistencia', semanal: 'Semanal', resumen: 'Resumen' } as const;

/** Filas dentro del bloque de cada persona en la hoja Asistencia (neto1/neto2 van ocultas). */
export const FILA = {
  reloj: 0, entrada: 1, salida: 2, motivo: 3, trabajado: 4, extra1: 5, extra2: 6, feriado: 7, contra: 8, compensado: 9, neto1: 10, neto2: 11,
} as const;
type Fila = keyof typeof FILA;
export const ALTO_BLOQUE = 12;
export const PRIMERA_FILA_BLOQUE = 10;
export const PRIMERA_COL_DIA = 3; // C
const R_DIASEM = 5, R_FECHA = 6, R_TIPO = 7, R_SEMANA = 8;
const OCULTAS: Fila[] = ['neto1', 'neto2'];

export function filaBloque(i: number): number {
  return PRIMERA_FILA_BLOQUE + i * ALTO_BLOQUE;
}

/** Columnas del Resumen. */
export const RES = {
  primeraFila: 6,
  nombre: 2, empresa: 3, faltas: 4, trabajado: 5, base: 6,
  extra1: 7, extra2: 8, compensado: 9,
  neto1: 10, neto2: 11, feriado: 12, valorHora: 13, monto: 14,
  nosDebe: 15, contra: 16, pendientes: 17,
} as const;

export interface OpcionesExcel {
  logoBase64?: string; // sin el prefijo data:
  generado?: Date;
  /** Para pruebas: no guardar los resultados precalculados, así el recálculo es obligatorio. */
  sinResultados?: boolean;
}

export async function generarExcel(mes: Mes, cfg: Config, _rep: Reporte, op: OpcionesExcel = {}): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Control de asistencia QALEN';
  wb.created = op.generado ?? new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  const logoId = op.logoBase64 ? wb.addImage({ base64: op.logoBase64, extension: 'png' }) : null;

  const { y, mo } = parseISO(mes.inicio);
  const titulo = `${MESES[mo]} ${y}`;
  const periodo = `del ${ddmm(mes.inicio)} al ${ddmm(mes.hasta)}`;
  const R = cfg.reglas;
  const MARG = R.margenExtras, ALM = R.almuerzoMin, TD1 = Math.max(R.tardanzaDesde, 1), TRAMO = R.tramo1Min;
  const pct = (f: number) => `${Math.round((f - 1) * 100)}%`;
  const P1 = pct(R.recargo1), P2 = pct(R.recargo2), PF = pct(R.recargoFeriado);
  const horasTramo = TRAMO % 60 ? `${TRAMO} min` : `${TRAMO / 60} horas`;

  const ETIQUETAS: Record<Fila, string> = {
    reloj: 'Reloj', entrada: 'Entrada', salida: 'Salida', motivo: 'Motivo / Obs.', trabajado: 'Horas trabajadas',
    extra1: `Extra ${P1}`, extra2: `Extra ${P2}`, feriado: 'Dom./Feriado', contra: 'En contra', compensado: 'Compensadas',
    neto1: `Extra ${P1} a pagar`, neto2: `Extra ${P2} a pagar`,
  };
  const COLOR_TXT: Partial<Record<Fila, string>> = { extra1: VERDE_TXT, extra2: VERDE_OSC, feriado: AZUL_FER, contra: ROJO_TXT, compensado: MORADO, neto1: VERDE_TXT, neto2: VERDE_OSC };

  const leeme = wb.addWorksheet(HOJAS.leeme, { properties: { tabColor: { argb: GRIS_OSC } } });
  const asis = wb.addWorksheet(HOJAS.asis, { properties: { tabColor: { argb: ROJO } } });
  const semanal = wb.addWorksheet(HOJAS.semanal, { properties: { tabColor: { argb: 'FF5B9BD5' } } });
  const resumen = wb.addWorksheet(HOJAS.resumen, { properties: { tabColor: { argb: 'FF70AD47' } } });

  const ponerLogo = (ws: ExcelJS.Worksheet, ancho = 150) => {
    if (logoId != null) ws.addImage(logoId, { tl: { col: 0.1, row: 0.15 }, ext: { width: ancho, height: Math.round((ancho * 208) / 660) } });
  };
  const proteger = (ws: ExcelJS.Worksheet, extra: Partial<ExcelJS.WorksheetProtection> = {}) =>
    ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true, formatCells: true, objects: true, ...extra });

  // =====================================================================
  // ASISTENCIA
  // =====================================================================
  const nD = mes.fechas.length;
  const cDia = (i: number) => PRIMERA_COL_DIA + i;
  const cUlt = cDia(nD - 1);
  const cTot = cUlt + 1;
  const cEtq2 = cTot + 1; // repite las etiquetas de la columna B
  const cPar = cTot + 3; // parámetros por persona (ocultos)
  const PAR = { ent: cPar, sLV: cPar + 1, sSab: cPar + 2, xLV: cPar + 3, temp: cPar + 4, desc: cPar + 5, jor: cPar + 6, ing: cPar + 7 };
  const L = (c: number) => col(c);
  const X0 = L(cDia(0));
  const rangoFila = (r: number) => `$${X0}$${r}:$${L(cUlt)}$${r}`;
  const inicioSemana = (j: number) => j > 0 && mes.fechas[j].semana !== mes.fechas[j - 1].semana;
  const pendTotal = (i: number) => `${L(cTot)}${filaBloque(i) + FILA.motivo}`;

  {
    const ws = asis;
    ponerLogo(ws);
    ws.getColumn(1).width = 13;
    ws.getColumn(2).width = 14;
    for (let i = 0; i < nD; i++) ws.getColumn(cDia(i)).width = 7.3;
    ws.getColumn(cTot).width = 10;
    ws.getColumn(cEtq2).width = 14;
    for (let c = cPar; c <= PAR.ing; c++) ws.getColumn(c).hidden = true;

    const t = ws.getCell(1, 4);
    t.value = `ASISTENCIA ${titulo}`;
    t.font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell(2, 4).value = `Período ${periodo}. Reporte del reloj hasta el ${ddmm(mes.finReporte)}.`;
    ws.getCell(2, 4).font = { color: { argb: GRIS_OSC } };
    const ban = ws.getCell(3, 4);
    const totPend = mes.filas.length ? mes.filas.map((_, i) => pendTotal(i)).join('+') : '0';
    const totAlert = mes.filas.length
      ? mes.filas.map((_, i) => {
          const b = filaBloque(i);
          return `SUMPRODUCT(ISNUMBER(${rangoFila(b + FILA.entrada)})*ISNUMBER(${rangoFila(b + FILA.salida)})*(${rangoFila(b + FILA.salida)}<=${rangoFila(b + FILA.entrada)}))`;
        }).join('+')
      : '0';
    ban.value = {
      formula: `IF((${totPend})+(${totAlert})=0,"✔ Todo completo.","⚠ Faltan "&(${totPend})&" datos por completar (celdas amarillas)"&IF((${totAlert})>0," y hay "&(${totAlert})&" horas en rojo por revisar","")&".")`,
    };
    ban.font = { bold: true, size: 12, color: { argb: ROJO_TXT } };

    // Cabeceras de días
    for (const [r, txt] of [[R_DIASEM, 'Día'], [R_FECHA, 'Fecha'], [R_TIPO, 'Tipo'], [R_SEMANA, 'Semana']] as const) {
      for (const c of [2, cEtq2]) {
        const cell = ws.getCell(r, c);
        cell.value = txt;
        cell.font = { bold: true, size: 9, color: { argb: GRIS_OSC } };
        cell.alignment = { horizontal: c === 2 ? 'right' : 'left' };
      }
    }
    mes.fechas.forEach((f, i) => {
      const c = cDia(i);
      const h1 = ws.getCell(R_DIASEM, c);
      h1.value = DIAS_CORTO[weekday(f.fecha)];
      const h2 = ws.getCell(R_FECHA, c);
      h2.value = excelSerial(f.fecha);
      h2.numFmt = 'dd/mm';
      const h3 = ws.getCell(R_TIPO, c);
      h3.value = f.tipo;
      const h4 = ws.getCell(R_SEMANA, c);
      h4.value = f.semana;
      h4.numFmt = '"Sem "0';
      for (const h of [h1, h2, h3, h4]) {
        h.alignment = { horizontal: 'center' };
        h.font = { bold: h === h1 || h === h2, size: h === h2 ? 11 : 9, color: { argb: h === h2 ? 'FFFFFFFF' : GRIS_OSC } };
      }
      h2.fill = fill(f.tipo === 'LV' ? ROJO : f.tipo === 'SAB' ? 'FFE0707F' : GRIS_OSC);
      if (f.tipo === 'FER' || f.tipo === 'DOM') for (const h of [h1, h3]) h.fill = fill(ROSA);
    });
    const tc = ws.getCell(R_FECHA, cTot);
    tc.value = 'TOTAL';
    estiloCabecera(tc);

    // Bloques por persona
    mes.filas.forEach((fp, i) => {
      const r0 = filaBloque(i);
      const p = fp.persona;
      const RF = (k: Fila) => r0 + FILA[k];
      ws.mergeCells(r0, 1, r0 + ALTO_BLOQUE - 1, 1);
      const nc = ws.getCell(r0, 1);
      nc.value = p.alias;
      nc.font = { bold: true, size: 11 };
      nc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      nc.fill = fill(GRIS);

      const colorFila = (k: Fila) => (k === 'entrada' ? VERDE_ENTRADA : k === 'salida' ? NARANJA_SALIDA : k === 'motivo' ? 'FFFFFFFF' : GRIS);
      for (const k of Object.keys(FILA) as Fila[]) {
        for (const c of [2, cEtq2]) {
          const cell = ws.getCell(RF(k), c);
          cell.value = ETIQUETAS[k];
          cell.font = { bold: k !== 'reloj' && k !== 'trabajado', size: 9, color: { argb: COLOR_TXT[k] ?? 'FF000000' } };
          cell.fill = fill(colorFila(k));
        }
      }
      ws.getRow(RF('reloj')).height = 24;
      for (const k of OCULTAS) ws.getRow(RF(k)).hidden = true;

      // Parámetros de la persona (valores fijos, columnas ocultas)
      const pars: [number, number][] = [
        [PAR.ent, p.entrada], [PAR.sLV, p.salidaLV], [PAR.sSab, p.salidaSab], [PAR.xLV, p.extrasDesdeLV],
        [PAR.temp, p.tempranoAntesDe ?? 0], [PAR.desc, p.descuenta ? 1 : 0], [PAR.jor, baseDiaria(p)],
        [PAR.ing, p.fechaIngreso ? excelSerial(p.fechaIngreso) : 0],
      ];
      for (const [c, v] of pars) ws.getCell(r0, c).value = v;
      const P = (c: number) => `$${L(c)}$${r0}`;
      const IN = P(PAR.ent), OLV = P(PAR.sLV), OSAB = P(PAR.sSab), XLV = P(PAR.xLV), TEMP = P(PAR.temp), DESC = P(PAR.desc), JOR = P(PAR.jor), ING = P(PAR.ing);
      const COMP_TOT = `ROUND(SUM(${rangoFila(RF('compensado'))})*1440,0)`;

      fp.dias.forEach((d, j) => {
        const c = cDia(j);
        const X = L(c);
        const E = `${X}${RF('entrada')}`, S = `${X}${RF('salida')}`, M = `${X}${RF('motivo')}`;
        const T = `${X}$${R_TIPO}`, F = `${X}$${R_FECHA}`;
        const e = `ROUND(${E}*1440,0)`, s = `ROUND(${S}*1440,0)`;
        const FIN = `IF(${T}="SAB",${OSAB},${OLV})`;
        const XD = `IF(${T}="SAB",${OSAB},${XLV})`;
        const JORD = `MIN(${JOR},${FIN}-${IN})`;
        const PAG = `OR(${M}="Vacaciones",${M}="Salud")`;
        const LAB = `OR(${T}="LV",${T}="SAB")`;
        const SINM = `AND(NOT(ISNUMBER(${E})),NOT(ISNUMBER(${S})))`;
        const FALTAN = `IF(ISNUMBER(${E}),IF(${e}-${IN}>=${TD1},${e}-${IN},0),0)+IF(ISNUMBER(${S}),IF(${s}<${FIN},${FIN}-${s},0),0)`;
        // Minutos extra del día (lunes a sábado)
        const EXT =
          `IF(${F}<${ING},0,IF(${LAB},IF(OR(${PAG},${M}="Falta"),0,` +
          `IF(ISNUMBER(${S}),IF(${s}>=${XD}+${MARG},${s}-${XD},0),0)` +
          `+IF(ISNUMBER(${E}),IF(AND(${TEMP}>0,${e}<${TEMP},${e}<${IN}),${IN}-${e},0),0)),0))`;

        const cr = ws.getCell(RF('reloj'), c);
        cr.value = d.reloj || null;
        cr.font = { size: 8, color: { argb: 'FF404040' } };
        cr.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        cr.fill = fill(GRIS_RELOJ);
        const notas = d.casos.map((x) => `• ${x.mensaje}${x.sugerencia ? ` Sugerencia: ${x.sugerencia}.` : ''}`);
        if (d.clas.duplicadas.length) notas.push(`• Se quitaron marcadas repetidas: ${d.clas.duplicadas.map(fmtHHMM).join(', ')}.`);
        if (d.clas.almuerzoSale != null) notas.push(`Almuerzo: ${fmtHHMM(d.clas.almuerzoSale)}${d.clas.almuerzoVuelve != null ? ` – ${fmtHHMM(d.clas.almuerzoVuelve)}` : ''}.`);
        if (d.clas.otras.length) notas.push(`Otras marcadas: ${d.clas.otras.map(fmtHHMM).join(', ')}.`);
        if (notas.length && d.aplica) cr.note = notas.join('\n');

        const ce = ws.getCell(RF('entrada'), c);
        const cs = ws.getCell(RF('salida'), c);
        const cm = ws.getCell(RF('motivo'), c);
        ce.value = d.entrada != null ? d.entrada / 1440 : null;
        cs.value = d.salida != null ? d.salida / 1440 : null;
        cm.value = d.motivo || null;
        for (const x of [ce, cs, cm]) {
          x.protection = { locked: false };
          x.alignment = { horizontal: 'center' };
          x.border = bordes;
        }
        ce.numFmt = 'hh:mm';
        cs.numFmt = 'hh:mm';
        ce.fill = fill(VERDE_ENTRADA);
        cs.fill = fill(NARANJA_SALIDA);
        cm.font = { size: 8 };
        cm.alignment = { horizontal: 'center', shrinkToFit: true };

        const trab = `IF(${F}<${ING},0,IF(AND(${LAB},${PAG}),${FIN}-${IN}-IF(${T}="LV",${ALM},0),IF(ISNUMBER(${E}),IF(ISNUMBER(${S}),IF(${s}>${e},${s}-${e}-IF(AND(${T}="LV",${s}-${e}>300),${ALM},0),0),0),0)))/1440`;
        const ext1 = `MIN(${EXT},${TRAMO})/1440`;
        const ext2 = `MAX(${EXT}-${TRAMO},0)/1440`;
        const fer = `IF(${F}<${ING},0,IF(${LAB},0,IF(ISNUMBER(${E}),IF(ISNUMBER(${S}),IF(${s}>${e},${s}-${e},0),0),0)))/1440`;
        const contra =
          `IF(OR(${F}<${ING},NOT(${LAB}),${DESC}=0,${PAG},${M}="Compensado"),0,` +
          `IF(${M}="Falta",${JORD},IF(AND(${M}="Estudio",${SINM}),${JORD},${FALTAN})))/1440`;
        const comp = `IF(OR(${F}<${ING},NOT(${LAB}),${DESC}=0),0,IF(${M}="Compensado",IF(${SINM},${JORD},${FALTAN}),0))/1440`;
        // Compensadas que quedan por descontar al empezar este día (se descuentan desde el primer día)
        const ANT = j === 0 ? '0' : `ROUND((SUM($${X0}${RF('extra1')}:${L(c - 1)}${RF('extra1')})+SUM($${X0}${RF('extra2')}:${L(c - 1)}${RF('extra2')}))*1440,0)`;
        const RESTO = `MAX(${COMP_TOT}-${ANT},0)`;
        const E1 = `ROUND(${X}${RF('extra1')}*1440,0)`, E2 = `ROUND(${X}${RF('extra2')}*1440,0)`;
        const net2 = `(${E2}-MIN(${E2},${RESTO}))/1440`;
        const net1 = `(${E1}-MIN(${E1},MAX(${RESTO}-${E2},0)))/1440`;

        const filas: [Fila, string, number][] = [
          ['trabajado', trab, d.calc.trabajado],
          ['extra1', ext1, d.calc.extra1],
          ['extra2', ext2, d.calc.extra2],
          ['feriado', fer, d.calc.enFeriado],
          ['contra', contra, d.calc.contra],
          ['compensado', comp, d.calc.compensado],
          ['neto1', net1, d.neto.neto1],
          ['neto2', net2, d.neto.neto2],
        ];
        for (const [k, f, res] of filas) {
          const cell = ws.getCell(RF(k), c);
          cell.value = { formula: f, result: res / 1440 };
          cell.numFmt = '[h]:mm;;""';
          cell.fill = fill(GRIS);
          cell.alignment = { horizontal: 'center' };
          cell.font = { size: 9, bold: k !== 'trabajado', color: { argb: COLOR_TXT[k] ?? 'FF000000' } };
        }
      });

      // Totales de la fila
      const tot = (k: Fila, f: string, res: number, fmt: string) => {
        const c = ws.getCell(RF(k), cTot);
        c.value = { formula: f, result: res };
        c.numFmt = fmt;
        c.font = { bold: true, color: { argb: COLOR_TXT[k] ?? (k === 'motivo' ? ROJO_TXT : 'FF000000') } };
        c.alignment = { horizontal: 'center' };
        c.border = bordes;
        c.fill = fill(GRIS);
      };
      const T = fp.tot;
      const totales: [Fila, number][] = [
        ['trabajado', T.trabajado], ['extra1', T.extra1], ['extra2', T.extra2], ['feriado', T.enFeriado],
        ['contra', T.contra], ['compensado', T.compensado], ['neto1', T.neto1], ['neto2', T.neto2],
      ];
      for (const [k, v] of totales) tot(k, `SUM(${rangoFila(RF(k))})`, v / 1440, '[h]:mm');
      // Datos por completar (conteo de celdas amarillas)
      const Er = rangoFila(RF('entrada')), Sr = rangoFila(RF('salida')), Mr = rangoFila(RF('motivo'));
      const Tr = rangoFila(R_TIPO), Fr = rangoFila(R_FECHA);
      const xorR = `(ISNUMBER(${Er})<>ISNUMBER(${Sr}))`;
      tot(
        'motivo',
        `SUMPRODUCT((${Fr}>=${ING})*(((${Tr}="LV")+(${Tr}="SAB"))*((${Mr}="")*(1-ISNUMBER(${Er})*ISNUMBER(${Sr}))+((${Mr}="Compensado")+(${Mr}="Estudio"))*${xorR})+((${Tr}="DOM")+(${Tr}="FER"))*${xorR}))`,
        T.pendientes,
        '"Faltan "0;;"✔"',
      );

      // Formatos condicionales del bloque
      const Ex = `${X0}${RF('entrada')}`, Sx = `${X0}${RF('salida')}`, Mx = `${X0}${RF('motivo')}`;
      const Tx = `${X0}$${R_TIPO}`, Fx = `${X0}$${R_FECHA}`;
      const PEND = `AND(${Fx}>=${ING},IF(OR(${Tx}="LV",${Tx}="SAB"),IF(${Mx}="",NOT(AND(ISNUMBER(${Ex}),ISNUMBER(${Sx}))),AND(OR(${Mx}="Compensado",${Mx}="Estudio"),ISNUMBER(${Ex})<>ISNUMBER(${Sx}))),ISNUMBER(${Ex})<>ISNUMBER(${Sx})))`;
      const rango = (k: Fila) => `${X0}${RF(k)}:${L(cUlt)}${RF(k)}`;
      for (const k of ['entrada', 'salida'] as const) {
        const ref = k === 'entrada' ? Ex : Sx;
        ws.addConditionalFormatting({
          ref: rango(k),
          rules: [
            { type: 'expression', priority: 1, formulae: [`${Fx}<${ING}`], style: { fill: cfFill('FFD9D9D9'), font: { color: { argb: GRIS_OSC } } } },
            { type: 'expression', priority: 2, formulae: [`AND(ISNUMBER(${Ex}),ISNUMBER(${Sx}),${Sx}<=${Ex})`], style: { fill: cfFill('FFFF7C80'), font: { bold: true } } },
            { type: 'expression', priority: 3, formulae: [`AND(${PEND},NOT(ISNUMBER(${ref})))`], style: { fill: cfFill(AMARILLO) } },
            { type: 'expression', priority: 4, formulae: [`AND(ISNUMBER(${ref}),ISERROR(SEARCH(TEXT(${ref},"hh:mm"),${X0}${RF('reloj')})))`], style: { fill: cfFill(AZUL), font: { italic: true, color: { argb: AZUL_FER } } } },
          ],
        });
      }
      ws.addConditionalFormatting({
        ref: rango('motivo'),
        rules: [
          { type: 'expression', priority: 1, formulae: [`${Fx}<${ING}`], style: { fill: cfFill('FFD9D9D9') } },
          { type: 'expression', priority: 2, formulae: [`AND(${PEND},${Mx}="",OR(${Tx}="LV",${Tx}="SAB"))`], style: { fill: cfFill(AMARILLO) } },
          { type: 'expression', priority: 3, formulae: [`${Mx}<>""`], style: { fill: cfFill(VERDE) } },
        ],
      });

      // Validaciones (aviso con confirmación, no bloqueo)
      dv(ws).add(rango('entrada'), {
        type: 'custom', allowBlank: true, formulae: [`OR(${Ex}="",AND(ISNUMBER(${Ex}),${Ex}<1))`],
        showErrorMessage: true, errorStyle: 'warning', errorTitle: '¿Hora correcta?',
        error: 'Escribe la hora así: 8:30 o 18:05 (formato de 24 horas). ¿Quieres dejarlo igual?',
      });
      dv(ws).add(rango('salida'), {
        type: 'custom', allowBlank: true,
        formulae: [`OR(${Sx}="",AND(ISNUMBER(${Sx}),${Sx}<1,OR(NOT(ISNUMBER(${Ex})),${Sx}>${Ex})))`],
        showErrorMessage: true, errorStyle: 'warning', errorTitle: '¿Hora de salida correcta?',
        error: 'La salida debe ser DESPUÉS de la entrada y en formato de 24 horas.\nEjemplo: 6:14 de la tarde se escribe 18:14.\n¿Quieres dejarlo igual?',
      });
      dv(ws).add(rango('motivo'), {
        type: 'list', allowBlank: true, formulae: [`"${MOTIVOS.join(',')}"`],
        showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Motivo', error: `Elige uno de la lista: ${MOTIVOS.join(', ')}.`,
        showInputMessage: true, promptTitle: 'Motivo',
        prompt: 'Vacaciones / Salud: día completo trabajado. Compensado: se descuenta de las horas extras. Estudio y Falta: horas en contra.',
      });
    });

    // Líneas: semanas (vertical) y separación Salida | Motivo (horizontal)
    const ultFila = mes.filas.length ? filaBloque(mes.filas.length - 1) + ALTO_BLOQUE - 1 : R_SEMANA;
    for (let r = R_DIASEM; r <= ultFila; r++) {
      for (let j = 0; j < nD; j++) {
        if (j === 0 || inicioSemana(j)) {
          const cell = ws.getCell(r, cDia(j));
          cell.border = { ...(cell.border ?? {}), left: grueso };
        }
      }
      const cell = ws.getCell(r, cTot);
      cell.border = { ...(cell.border ?? {}), left: grueso };
    }
    mes.filas.forEach((_, i) => {
      const r0 = filaBloque(i);
      for (let c = 1; c <= cEtq2; c++) {
        const top = ws.getCell(r0, c);
        top.border = { ...(top.border ?? {}), top: { style: 'medium', color: { argb: ROJO } } };
        if (c === 1) continue;
        const sal = ws.getCell(r0 + FILA.salida, c);
        sal.border = { ...(sal.border ?? {}), bottom: grueso };
      }
    });

    // Leyenda
    const rLey = ultFila + 2;
    const ley: [string, string][] = [
      [AMARILLO, 'Falta un dato: escribe la hora o elige un motivo.'],
      [AZUL, 'Hora escrita a mano (no está en el reloj).'],
      ['FFFF7C80', 'Revisar: la salida es antes que la entrada (¿formato de 12 h?).'],
      [VERDE, 'Día con motivo.'],
      [GRIS, 'Calculado con fórmulas: no se escribe aquí.'],
    ];
    ley.forEach(([c, t2], i) => {
      ws.getCell(rLey + i, 2).fill = fill(c);
      ws.getCell(rLey + i, 2).border = bordes;
      ws.getCell(rLey + i, 3).value = t2;
      ws.getCell(rLey + i, 3).font = { size: 9 };
    });

    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: R_SEMANA, activeCell: `${X0}${PRIMERA_FILA_BLOQUE + 1}`, showGridLines: false }];
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    await proteger(ws);
  }

  const refRango = (r: number) => `'${HOJAS.asis}'!$${X0}$${r}:$${L(cUlt)}$${r}`;
  const refFila = (k: Fila, i: number) => refRango(filaBloque(i) + FILA[k]);
  const refTot = (k: Fila, i: number) => `'${HOJAS.asis}'!$${L(cTot)}$${filaBloque(i) + FILA[k]}`;

  // =====================================================================
  // RESUMEN
  // =====================================================================
  {
    const ws = resumen;
    ponerLogo(ws, 130);
    ws.getCell('C1').value = `RESUMEN ${titulo}`;
    ws.getCell('C1').font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell('C2').value = `Período ${periodo}. Todo se calcula solo desde la hoja Asistencia. Solo el valor hora se puede corregir aquí.`;
    ws.getCell('C2').font = { color: { argb: GRIS_OSC } };
    const r1 = RES.primeraFila;
    const rT = r1 + mes.filas.length;
    const ban = ws.getCell('C3');
    const sinVH = mes.filas.length ? `COUNTIF(${col(RES.valorHora)}${r1}:${col(RES.valorHora)}${rT - 1},"<=0")+COUNTBLANK(${col(RES.valorHora)}${r1}:${col(RES.valorHora)}${rT - 1})` : '0';
    ban.value = {
      formula:
        `IF(AND(${col(RES.pendientes)}${rT}=0,${sinVH}=0),"✔ Sin datos pendientes.",` +
        `IF(${col(RES.pendientes)}${rT}>0,"⚠ Todavía faltan "&${col(RES.pendientes)}${rT}&" datos en la hoja Asistencia: el resumen puede cambiar. ","")` +
        `&IF(${sinVH}>0,"⚠ Falta el valor hora de "&${sinVH}&" persona(s) (celdas amarillas).",""))`,
    };
    ban.font = { bold: true, size: 12, color: { argb: ROJO_TXT } };

    const grupos: [number, number, string, string][] = [
      [RES.extra1, RES.compensado, 'HORAS EXTRAS DEL MES', 'FF548235'],
      [RES.neto1, RES.monto, 'A PAGAR (después de descontar compensadas)', 'FF1E5631'],
      [RES.nosDebe, RES.pendientes, 'APARTE', 'FF7F7F7F'],
    ];
    for (const [a, b, txt, color] of grupos) {
      ws.mergeCells(r1 - 2, a, r1 - 2, b);
      const c = ws.getCell(r1 - 2, a);
      c.value = txt;
      c.alignment = { horizontal: 'center' };
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = fill(color);
    }
    const cab: [number, string, number, string?][] = [
      [1, 'N°', 5], [RES.nombre, 'Nombre', 14], [RES.empresa, 'Empresa', 11], [RES.faltas, 'Faltas', 7],
      [RES.trabajado, 'Horas trabajadas', 10], [RES.base, 'Horas base del mes', 10],
      [RES.extra1, `Extras ${P1}`, 10, 'FF548235'], [RES.extra2, `Extras ${P2}`, 10, 'FF548235'], [RES.compensado, 'Compensadas', 11, MORADO],
      [RES.neto1, `Extras ${P1}`, 10, 'FF1E5631'], [RES.neto2, `Extras ${P2}`, 10, 'FF1E5631'], [RES.feriado, `Dom./feriado ${PF}`, 10, AZUL_FER],
      [RES.valorHora, 'Valor hora', 10, 'FF1E5631'], [RES.monto, 'TOTAL A PAGAR', 13, 'FF1E5631'],
      [RES.nosDebe, 'Nos debe (compensadas sin cubrir)', 12, ROJO_TXT], [RES.contra, 'Horas en contra', 10, 'FFC00000'], [RES.pendientes, 'Datos por completar', 10],
    ];
    for (const [c, t, w, color] of cab) {
      const cell = ws.getCell(r1 - 1, c);
      cell.value = t;
      estiloCabecera(cell);
      if (color) cell.fill = fill(color);
      ws.getColumn(c).width = w;
    }
    ws.getRow(r1 - 1).height = 45;

    mes.filas.forEach((fp, i) => {
      const r = r1 + i;
      const p = fp.persona;
      const b = filaBloque(i);
      const S = (c: number) => `${col(c)}${r}`;
      const Tr = refRango(R_TIPO), Fr = refRango(R_FECHA);
      const ING = `'${HOJAS.asis}'!$${L(PAR.ing)}$${b}`, JOR = `'${HOJAS.asis}'!$${L(PAR.jor)}$${b}`;
      const t = fp.tot;
      const vals: [number, ExcelJS.CellValue, string?][] = [
        [1, i + 1],
        [RES.nombre, { text: p.alias, hyperlink: `#'${HOJAS.asis}'!A${b}` }],
        [RES.empresa, p.empresa || null],
        [RES.faltas, { formula: `COUNTIF(${refFila('motivo', i)},"Falta")`, result: t.faltas }, '0;;""'],
        [RES.trabajado, { formula: refTot('trabajado', i), result: t.trabajado / 1440 }, '[h]:mm'],
        [RES.base, { formula: `(COUNTIFS(${Tr},"LV",${Fr},">="&${ING})+COUNTIFS(${Tr},"SAB",${Fr},">="&${ING}))*${JOR}/1440`, result: t.base / 1440 }, '[h]:mm'],
        [RES.extra1, { formula: refTot('extra1', i), result: t.extra1 / 1440 }, '[h]:mm'],
        [RES.extra2, { formula: refTot('extra2', i), result: t.extra2 / 1440 }, '[h]:mm'],
        [RES.compensado, { formula: refTot('compensado', i), result: t.compensado / 1440 }, '[h]:mm'],
        [RES.neto1, { formula: refTot('neto1', i), result: t.neto1 / 1440 }, '[h]:mm'],
        [RES.neto2, { formula: refTot('neto2', i), result: t.neto2 / 1440 }, '[h]:mm'],
        [RES.feriado, { formula: refTot('feriado', i), result: t.enFeriado / 1440 }, '[h]:mm'],
        [RES.valorHora, p.valorHora || null, SOLES],
        [RES.monto, { formula: `ROUND(${S(RES.neto1)}*24*${S(RES.valorHora)}*${R.recargo1}+${S(RES.neto2)}*24*${S(RES.valorHora)}*${R.recargo2}+${S(RES.feriado)}*24*${S(RES.valorHora)}*${R.recargoFeriado},2)`, result: t.monto }, SOLES],
        [RES.nosDebe, { formula: `MAX(${S(RES.compensado)}-${S(RES.extra1)}-${S(RES.extra2)},0)`, result: t.nosDebe / 1440 }, '[h]:mm;;""'],
        [RES.contra, { formula: refTot('contra', i), result: t.contra / 1440 }, '[h]:mm'],
        [RES.pendientes, { formula: `'${HOJAS.asis}'!${pendTotal(i)}`, result: t.pendientes }, '0;;"✔"'],
      ];
      for (const [c, v, fmt] of vals) {
        const cell = ws.getCell(r, c);
        cell.value = v;
        if (fmt) cell.numFmt = fmt;
        cell.border = bordes;
        cell.alignment = { horizontal: c === RES.nombre ? 'left' : 'center' };
      }
      ws.getCell(r, RES.nombre).font = { bold: true, color: { argb: AZUL_FER }, underline: true };
      ws.getCell(r, RES.compensado).font = { bold: true, color: { argb: MORADO } };
      ws.getCell(r, RES.feriado).font = { bold: true, color: { argb: AZUL_FER } };
      ws.getCell(r, RES.monto).font = { bold: true, size: 12, color: { argb: VERDE_OSC } };
      ws.getCell(r, RES.monto).fill = fill('FFC6EFCE');
      const vh = ws.getCell(r, RES.valorHora);
      vh.protection = { locked: false };
      vh.dataValidation = { type: 'decimal', operator: 'between', formulae: [0, 999], allowBlank: true, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Valor hora', error: 'Escribe el valor de una hora en soles, por ejemplo 6.25 (sueldo ÷ 240).' };
    });
    // Totales
    ws.getCell(rT, RES.nombre).value = 'TOTAL';
    ws.getCell(rT, RES.nombre).font = { bold: true };
    for (let c: number = RES.faltas; c <= RES.pendientes; c++) {
      if (c === RES.valorHora) continue;
      const cell = ws.getCell(rT, c);
      cell.value = { formula: mes.filas.length ? `SUM(${col(c)}${r1}:${col(c)}${rT - 1})` : '0' };
      cell.numFmt = ws.getCell(r1, c).numFmt;
      cell.font = { bold: true, color: { argb: c === RES.monto ? VERDE_OSC : 'FF000000' } };
      cell.alignment = { horizontal: 'center' };
      cell.border = { top: { style: 'double' }, bottom: thin };
    }
    if (mes.filas.length) {
      const rr = (c: number) => `${col(c)}${r1}:${col(c)}${rT}`;
      const mayor0 = (c: number, style: Partial<ExcelJS.Style>) =>
        ws.addConditionalFormatting({ ref: rr(c), rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style }] });
      mayor0(RES.faltas, { font: { bold: true, color: { argb: ROJO_TXT } } });
      mayor0(RES.extra1, { font: { bold: true, color: { argb: VERDE_TXT } }, fill: cfFill(VERDE) });
      mayor0(RES.extra2, { font: { bold: true, color: { argb: VERDE_TXT } }, fill: cfFill(VERDE) });
      mayor0(RES.compensado, { fill: cfFill('FFE4DFEC') });
      mayor0(RES.neto1, { font: { bold: true, color: { argb: VERDE_OSC } } });
      mayor0(RES.neto2, { font: { bold: true, color: { argb: VERDE_OSC } } });
      mayor0(RES.nosDebe, { font: { bold: true, color: { argb: ROJO_TXT } }, fill: cfFill(ROSA) });
      mayor0(RES.contra, { font: { bold: true, color: { argb: ROJO_TXT } }, fill: cfFill(ROSA) });
      mayor0(RES.pendientes, { fill: cfFill(AMARILLO), font: { bold: true } });
      ws.addConditionalFormatting({
        ref: `${col(RES.valorHora)}${r1}:${col(RES.valorHora)}${rT - 1}`,
        rules: [{ type: 'expression', priority: 1, formulae: [`OR(${col(RES.valorHora)}${r1}="",${col(RES.valorHora)}${r1}<=0)`], style: { fill: cfFill(AMARILLO) } }],
      });
      ws.autoFilter = { from: { row: r1 - 1, column: 1 }, to: { row: rT - 1, column: RES.pendientes } };
    }
    const notas = [
      `Extras: las ${horasTramo} primeras de cada día van al ${P1} y el resto al ${P2}. Domingo/feriado trabajado: ${PF}.`,
      'Las horas compensadas se descuentan de las extras de los primeros días del período (en cada día, primero las del ' + P2 + ').',
      'Si las compensadas son más que las extras, la diferencia aparece en "Nos debe". Las horas en contra van aparte y no se descuentan de las extras.',
      `TOTAL A PAGAR = (extras ${P1} × ${R.recargo1} + extras ${P2} × ${R.recargo2} + dom./feriado × ${R.recargoFeriado}) × valor hora.`,
    ];
    notas.forEach((txt, k) => {
      const c = ws.getCell(rT + 2 + k, RES.nombre);
      c.value = txt;
      c.font = { italic: true, size: 9, color: { argb: GRIS_OSC } };
    });
    ws.views = [{ state: 'frozen', ySplit: r1 - 1, xSplit: 2, showGridLines: false }];
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    await proteger(ws, { autoFilter: true, sort: true });
  }

  // =====================================================================
  // SEMANAL
  // =====================================================================
  {
    const ws = semanal;
    ponerLogo(ws, 130);
    ws.getCell('C1').value = `HORAS POR SEMANA ${titulo}`;
    ws.getCell('C1').font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell('C2').value = 'Semanas de lunes a sábado. Se calcula solo desde la hoja Asistencia. Extras antes de descontar compensadas.';
    ws.getCell('C2').font = { color: { argb: GRIS_OSC } };
    const semanas = [...new Set(mes.fechas.map((f) => f.semana))];
    const r1 = 7;
    ws.getColumn(1).width = 14;
    ws.getColumn(2).width = 9;
    for (const [c, t] of [[1, 'Nombre'], [2, 'Base semanal']] as const) {
      const cell = ws.getCell(r1 - 1, c);
      cell.value = t;
      estiloCabecera(cell);
    }
    ws.getRow(r1 - 1).height = 30;
    const semRango = refRango(R_SEMANA);
    const COLS: [Fila, string, string, (d: Mes['filas'][number]['dias'][number]) => number][] = [
      ['trabajado', 'Trabajadas', ROJO, (d) => d.calc.trabajado],
      ['extra1', `Extra ${P1}`, 'FF548235', (d) => d.calc.extra1],
      ['extra2', `Extra ${P2}`, 'FF548235', (d) => d.calc.extra2],
      ['feriado', 'Dom./Fer.', AZUL_FER, (d) => d.calc.enFeriado],
      ['contra', 'En contra', 'FFC00000', (d) => d.calc.contra],
      ['compensado', 'Compen- sadas', MORADO, (d) => d.calc.compensado],
    ];
    const ANCHO = COLS.length;
    semanas.forEach((k, si) => {
      const dias = mes.fechas.filter((f) => f.semana === k).map((f) => f.fecha);
      const parcial = weekday(dias[0]) !== 1 || dias.filter((f) => weekday(f) <= 6).length < 6;
      const a = 3 + si * ANCHO;
      ws.mergeCells(r1 - 2, a, r1 - 2, a + ANCHO - 1);
      const h = ws.getCell(r1 - 2, a);
      h.value = `Semana ${k}: ${ddmm(dias[0])} – ${ddmm(dias[dias.length - 1])}${parcial ? ' (parcial)' : ''}`;
      h.alignment = { horizontal: 'center' };
      h.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      h.fill = fill(si % 2 ? 'FF2F5597' : 'FF5B9BD5');
      COLS.forEach(([, t, color], j) => {
        const c = ws.getCell(r1 - 1, a + j);
        c.value = t;
        estiloCabecera(c);
        c.fill = fill(color);
        ws.getColumn(a + j).width = 9;
      });
      mes.filas.forEach((fp, i) => {
        const r = r1 + i;
        const dd = fp.dias.filter((_, j) => mes.fechas[j].semana === k);
        COLS.forEach(([kk, , , f], j) => {
          const c = ws.getCell(r, a + j);
          c.value = { formula: `SUMIF(${semRango},${k},${refFila(kk, i)})`, result: dd.reduce((acc, x) => acc + f(x), 0) / 1440 };
          c.numFmt = j === 0 ? '[h]:mm' : '[h]:mm;;""';
          c.border = j === ANCHO - 1 ? { ...bordes, right: grueso } : bordes;
          c.alignment = { horizontal: 'center' };
          c.font = { bold: j > 0, color: { argb: COLOR_TXT[kk] ?? 'FF000000' } };
        });
      });
    });
    mes.filas.forEach((fp, i) => {
      const r = r1 + i;
      const n = ws.getCell(r, 1);
      n.value = fp.persona.alias;
      n.font = { bold: true };
      n.border = bordes;
      const b = ws.getCell(r, 2);
      b.value = fp.persona.horasSemanales;
      b.numFmt = '0" h"';
      b.border = bordes;
      b.alignment = { horizontal: 'center' };
    });
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: r1 - 1, showGridLines: false }];
    await proteger(ws);
  }

  // =====================================================================
  // LEEME
  // =====================================================================
  {
    const ws = leeme;
    ponerLogo(ws, 170);
    ws.getColumn(1).width = 4;
    ws.getColumn(2).width = 5;
    ws.getColumn(3).width = 100;
    const lineas: [string, ('t' | 'h' | 'c')?, string?][] = [
      [`CONTROL DE ASISTENCIA ${titulo}`, 't'],
      [`Período ${periodo}.`],
      [''],
      ['Cómo trabajar', 'h'],
      ['1. En la hoja "Asistencia" completa las celdas AMARILLAS: escribe la hora (18:05) o elige un motivo.'],
      ['2. Arriba de la hoja dice cuántas faltan. Cuando diga "✔ Todo completo", revisa la hoja "Resumen".'],
      ['3. En "Resumen" revisa el valor hora de cada persona (si está en amarillo, escríbelo): el total a pagar se calcula solo.'],
      ['4. Si pasas el mouse por una celda gris con triángulo rojo (fila Reloj), verás qué pasó ese día.'],
      [''],
      ['Motivos', 'h'],
      ['Vacaciones / Salud: se cuenta como día completo trabajado.'],
      ['Compensado: la persona usa horas extras (sale antes o falta el día). Se descuenta de las extras de los primeros días del mes.'],
      ['Estudio: lo que no trabajó se cuenta como horas en contra.'],
      ['Falta: el día completo es hora en contra y se cuenta como falta.'],
      [''],
      ['Horas extras', 'h'],
      [`• Cuentan si sale a las 18:${String(MARG).padStart(2, '0')} o después, desde las 18:00 (sábado desde las 12:00 si sale 12:${String(MARG).padStart(2, '0')} o después).`],
      [`• Cada día: las ${horasTramo} primeras van al ${P1}; las siguientes, al ${P2}. Domingo o feriado trabajado: ${PF}.`],
      [`• Las compensadas se descuentan de las extras desde el primer día (primero las del ${P2}). Lo que no alcance a cubrirse aparece en "Nos debe".`],
      ['• Llegar temprano no suma horas.'],
      ['• Tardanza (después de las 08:30) y salida antes de hora: horas en contra. Van aparte, no se descuentan de las extras.'],
      [''],
      ['Colores', 'h'],
      ['Falta un dato.', 'c', AMARILLO],
      ['Entrada.', 'c', VERDE_ENTRADA],
      ['Salida.', 'c', NARANJA_SALIDA],
      ['Hora escrita a mano (no es la del reloj).', 'c', AZUL],
      ['Revisar: la salida es antes que la entrada. ¿Escribiste 6:14 en vez de 18:14?', 'c', 'FFFF7C80'],
      ['Calculado solo. No se escribe.', 'c', GRIS],
      [''],
      ['Las hojas están protegidas SIN contraseña solo para no borrar fórmulas por accidente.'],
    ];
    let r = 5;
    for (const [txt, tipo, color] of lineas) {
      const c = ws.getCell(r, 3);
      c.value = txt;
      c.alignment = { wrapText: true, vertical: 'middle' };
      if (tipo === 't') c.font = { bold: true, size: 16, color: { argb: ROJO } };
      else if (tipo === 'h') c.font = { bold: true, size: 12, color: { argb: ROJO } };
      if (tipo === 'c' && color) {
        ws.getCell(r, 2).fill = fill(color);
        ws.getCell(r, 2).border = bordes;
      }
      r++;
    }
    ws.views = [{ showGridLines: false }];
    ws.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  }

  if (op.sinResultados) {
    wb.eachSheet((ws) => ws.eachRow((row) => row.eachCell((c) => {
      if (c.formula) c.value = { formula: c.formula } as ExcelJS.CellFormulaValue;
    })));
  }
  wb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: 1, visibility: 'visible' }];
  return wb;
}

/** ExcelJS sí soporta validaciones por rango, pero no está en sus tipos. */
function dv(ws: ExcelJS.Worksheet): { add: (rango: string, v: ExcelJS.DataValidation) => void } {
  return (ws as unknown as { dataValidations: { add: (rango: string, v: ExcelJS.DataValidation) => void } }).dataValidations;
}

function estiloCabecera(c: ExcelJS.Cell) {
  c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
  c.fill = fill(ROJO);
  c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  c.border = bordes;
}

export async function excelABuffer(wb: ExcelJS.Workbook): Promise<ArrayBuffer> {
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}
