/**
 * Genera el Excel de trabajo de Yasmín (4 hojas: Leeme, Asistencia, Semanal, Resumen).
 *
 * Reparto del trabajo:
 *  - La página decide la ENTRADA y la SALIDA de cada día (o las deja vacías si hay duda).
 *  - Todo lo demás son fórmulas simples dentro del Excel: si Yasmín cambia una hora o
 *    elige un motivo, se recalculan solos los totales, la hoja Semanal y el Resumen.
 *
 * Los parámetros de cada persona (horario, etc.) van como valores en columnas ocultas de
 * la hoja Asistencia; no hay hoja de configuración que ella tenga que tocar.
 * Las fórmulas replican exactamente calcularDia() de rules.ts (hay una prueba que lo verifica
 * recalculando el archivo con LibreOffice).
 */
import ExcelJS from 'exceljs';
import { MOTIVOS, type Config } from './config';
import { baseDiaria, type Mes } from './model';
import type { Reporte } from './parser';
import { DIAS_CORTO, MESES, ddmm, excelSerial, fmtDur, fmtHHMM, parseISO, weekday } from './time';

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
const ROJO_TXT = 'FF9C0006';

const fill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const cfFill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', bgColor: { argb } });
const thin: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const grueso: Partial<ExcelJS.Border> = { style: 'medium', color: { argb: 'FF595959' } };
const bordes: Partial<ExcelJS.Borders> = { top: thin, left: thin, bottom: thin, right: thin };

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

/** Filas dentro del bloque de cada persona en la hoja Asistencia. */
export const FILA = { reloj: 0, entrada: 1, salida: 2, motivo: 3, trabajado: 4, favor: 5, contra: 6, compensado: 7 } as const;
export const ALTO_BLOQUE = 8;
export const PRIMERA_FILA_BLOQUE = 10;
export const PRIMERA_COL_DIA = 3; // C
const R_DIASEM = 5, R_FECHA = 6, R_TIPO = 7, R_SEMANA = 8;

export function filaBloque(i: number): number {
  return PRIMERA_FILA_BLOQUE + i * ALTO_BLOQUE;
}

/** Columnas del Resumen. */
export const RES = {
  primeraFila: 6,
  nombre: 2, empresa: 3, faltas: 4, trabajado: 5, base: 6, favor: 7, contra: 8, compensado: 9, saldo: 10, pendientes: 11,
} as const;

const ETIQUETAS: Record<keyof typeof FILA, string> = {
  reloj: 'Reloj', entrada: 'Entrada', salida: 'Salida', motivo: 'Motivo / Obs.', trabajado: 'Horas trabajadas', favor: 'A favor', contra: 'En contra', compensado: 'Compensadas',
};

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
  const MARG = R.margenExtras, ALM = R.almuerzoMin, TD1 = Math.max(R.tardanzaDesde, 1), FACT = R.factorFeriado;

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
      const RF = (k: keyof typeof FILA) => r0 + FILA[k];
      ws.mergeCells(r0, 1, r0 + ALTO_BLOQUE - 1, 1);
      const nc = ws.getCell(r0, 1);
      nc.value = p.alias;
      nc.font = { bold: true, size: 11 };
      nc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      nc.fill = fill(GRIS);

      const colorFila = (k: keyof typeof FILA) => (k === 'entrada' ? VERDE_ENTRADA : k === 'salida' ? NARANJA_SALIDA : k === 'motivo' ? 'FFFFFFFF' : GRIS);
      for (const k of Object.keys(FILA) as (keyof typeof FILA)[]) {
        for (const c of [2, cEtq2]) {
          const cell = ws.getCell(RF(k), c);
          cell.value = ETIQUETAS[k];
          cell.font = {
            bold: k !== 'reloj' && k !== 'trabajado',
            size: 9,
            color: { argb: k === 'favor' ? VERDE_TXT : k === 'contra' ? ROJO_TXT : k === 'compensado' ? MORADO : 'FF000000' },
          };
          cell.fill = fill(colorFila(k));
        }
      }
      ws.getRow(RF('reloj')).height = 24;

      // Parámetros de la persona (valores fijos, columnas ocultas)
      const pars: [number, number][] = [
        [PAR.ent, p.entrada], [PAR.sLV, p.salidaLV], [PAR.sSab, p.salidaSab], [PAR.xLV, p.extrasDesdeLV],
        [PAR.temp, p.tempranoAntesDe ?? 0], [PAR.desc, p.descuenta ? 1 : 0], [PAR.jor, baseDiaria(p)],
        [PAR.ing, p.fechaIngreso ? excelSerial(p.fechaIngreso) : 0],
      ];
      for (const [c, v] of pars) ws.getCell(r0, c).value = v;
      const P = (c: number) => `$${L(c)}$${r0}`;
      const IN = P(PAR.ent), OLV = P(PAR.sLV), OSAB = P(PAR.sSab), XLV = P(PAR.xLV), TEMP = P(PAR.temp), DESC = P(PAR.desc), JOR = P(PAR.jor), ING = P(PAR.ing);

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
        const favor =
          `IF(${F}<${ING},0,IF(${LAB},IF(OR(${PAG},${M}="Falta"),0,` +
          `IF(ISNUMBER(${S}),IF(${s}>=${XD}+${MARG},${s}-${XD},0),0)` +
          `+IF(ISNUMBER(${E}),IF(AND(${TEMP}>0,${e}<${TEMP},${e}<${IN}),${IN}-${e},0),0)),` +
          `IF(ISNUMBER(${E}),IF(ISNUMBER(${S}),IF(${s}>${e},ROUND((${s}-${e})*${FACT},0),0),0),0)))/1440`;
        const contra =
          `IF(OR(${F}<${ING},NOT(${LAB}),${DESC}=0,${PAG},${M}="Compensado"),0,` +
          `IF(${M}="Falta",${JORD},IF(AND(${M}="Estudio",${SINM}),${JORD},${FALTAN})))/1440`;
        const comp = `IF(OR(${F}<${ING},NOT(${LAB}),${DESC}=0),0,IF(${M}="Compensado",IF(${SINM},${JORD},${FALTAN}),0))/1440`;

        const filas: [keyof typeof FILA, string, number, string][] = [
          ['trabajado', trab, d.calc.trabajado, 'FF000000'],
          ['favor', favor, d.calc.favor, VERDE_TXT],
          ['contra', contra, d.calc.contra, ROJO_TXT],
          ['compensado', comp, d.calc.compensado, MORADO],
        ];
        for (const [k, f, res, color] of filas) {
          const cell = ws.getCell(RF(k), c);
          cell.value = { formula: f, result: res / 1440 };
          cell.numFmt = '[h]:mm;;""';
          cell.fill = fill(GRIS);
          cell.alignment = { horizontal: 'center' };
          cell.font = { size: 9, bold: k !== 'trabajado', color: { argb: color } };
        }
      });

      // Totales de la fila
      const tot = (k: keyof typeof FILA, f: string, res: number, fmt: string, color = 'FF000000') => {
        const c = ws.getCell(RF(k), cTot);
        c.value = { formula: f, result: res };
        c.numFmt = fmt;
        c.font = { bold: true, color: { argb: color } };
        c.alignment = { horizontal: 'center' };
        c.border = bordes;
        c.fill = fill(GRIS);
      };
      tot('trabajado', `SUM(${rangoFila(RF('trabajado'))})`, fp.tot.trabajado / 1440, '[h]:mm');
      tot('favor', `SUM(${rangoFila(RF('favor'))})`, fp.tot.favor / 1440, '[h]:mm', VERDE_TXT);
      tot('contra', `SUM(${rangoFila(RF('contra'))})`, fp.tot.contra / 1440, '[h]:mm', ROJO_TXT);
      tot('compensado', `SUM(${rangoFila(RF('compensado'))})`, fp.tot.compensado / 1440, '[h]:mm', MORADO);
      // Datos por completar (conteo de celdas amarillas)
      const Er = rangoFila(RF('entrada')), Sr = rangoFila(RF('salida')), Mr = rangoFila(RF('motivo'));
      const Tr = rangoFila(R_TIPO), Fr = rangoFila(R_FECHA);
      const xorR = `(ISNUMBER(${Er})<>ISNUMBER(${Sr}))`;
      tot(
        'motivo',
        `SUMPRODUCT((${Fr}>=${ING})*(((${Tr}="LV")+(${Tr}="SAB"))*((${Mr}="")*(1-ISNUMBER(${Er})*ISNUMBER(${Sr}))+((${Mr}="Compensado")+(${Mr}="Estudio"))*${xorR})+((${Tr}="DOM")+(${Tr}="FER"))*${xorR}))`,
        fp.tot.pendientes,
        '"Faltan "0;;"✔"',
        ROJO_TXT,
      );

      // Formatos condicionales del bloque
      const Ex = `${X0}${RF('entrada')}`, Sx = `${X0}${RF('salida')}`, Mx = `${X0}${RF('motivo')}`;
      const Tx = `${X0}$${R_TIPO}`, Fx = `${X0}$${R_FECHA}`;
      const PEND = `AND(${Fx}>=${ING},IF(OR(${Tx}="LV",${Tx}="SAB"),IF(${Mx}="",NOT(AND(ISNUMBER(${Ex}),ISNUMBER(${Sx}))),AND(OR(${Mx}="Compensado",${Mx}="Estudio"),ISNUMBER(${Ex})<>ISNUMBER(${Sx}))),ISNUMBER(${Ex})<>ISNUMBER(${Sx})))`;
      const rango = (k: keyof typeof FILA) => `${X0}${RF(k)}:${L(cUlt)}${RF(k)}`;
      for (const k of ['entrada', 'salida'] as const) {
        const ref = k === 'entrada' ? Ex : Sx;
        ws.addConditionalFormatting({
          ref: rango(k),
          rules: [
            { type: 'expression', priority: 1, formulae: [`${Fx}<${ING}`], style: { fill: cfFill('FFD9D9D9'), font: { color: { argb: GRIS_OSC } } } },
            { type: 'expression', priority: 2, formulae: [`AND(ISNUMBER(${Ex}),ISNUMBER(${Sx}),${Sx}<=${Ex})`], style: { fill: cfFill('FFFF7C80'), font: { bold: true } } },
            { type: 'expression', priority: 3, formulae: [`AND(${PEND},NOT(ISNUMBER(${ref})))`], style: { fill: cfFill(AMARILLO) } },
            { type: 'expression', priority: 4, formulae: [`AND(ISNUMBER(${ref}),ISERROR(SEARCH(TEXT(${ref},"hh:mm"),${X0}${RF('reloj')})))`], style: { fill: cfFill(AZUL), font: { italic: true, color: { argb: 'FF1F4E79' } } } },
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
        prompt: 'Vacaciones / Salud: día completo trabajado. Compensado: usa horas a favor. Estudio: horas en contra. Falta: horas en contra y cuenta como falta.',
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
  const refFila = (k: keyof typeof FILA, i: number) => refRango(filaBloque(i) + FILA[k]);
  const saldoTxt = (expr: string) => `IF(${expr}<0,"-","+")&TEXT(ABS(${expr}),"[h]:mm")`;
  const saldoRes = (m: number) => `${m < 0 ? '-' : '+'}${fmtDur(Math.abs(m))}`;

  // =====================================================================
  // RESUMEN
  // =====================================================================
  {
    const ws = resumen;
    ponerLogo(ws, 130);
    ws.getCell('C1').value = `RESUMEN ${titulo}`;
    ws.getCell('C1').font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell('C2').value = `Período ${periodo}. Todo se calcula solo desde la hoja Asistencia.`;
    ws.getCell('C2').font = { color: { argb: GRIS_OSC } };
    const r1 = RES.primeraFila;
    const rT = r1 + mes.filas.length;
    const ban = ws.getCell('C3');
    ban.value = { formula: `IF(${col(RES.pendientes)}${rT}=0,"✔ Sin datos pendientes.","⚠ Todavía faltan "&${col(RES.pendientes)}${rT}&" datos en la hoja Asistencia: el resumen puede cambiar.")` };
    ban.font = { bold: true, size: 12, color: { argb: ROJO_TXT } };

    const cab: [number, string, number, string?][] = [
      [1, 'N°', 5], [RES.nombre, 'Nombre', 14], [RES.empresa, 'Empresa', 11], [RES.faltas, 'Faltas', 8],
      [RES.trabajado, 'Horas trabajadas', 11], [RES.base, 'Horas base del mes', 11],
      [RES.favor, 'Horas a favor', 11, 'FF548235'], [RES.contra, 'Horas en contra', 11, 'FFC00000'], [RES.compensado, 'Horas compensadas', 12, MORADO],
      [RES.saldo, 'SALDO del mes', 12, 'FF404040'], [RES.pendientes, 'Datos por completar', 11],
    ];
    for (const [c, t, w, color] of cab) {
      const cell = ws.getCell(r1 - 1, c);
      cell.value = t;
      estiloCabecera(cell);
      if (color) cell.fill = fill(color);
      ws.getColumn(c).width = w;
    }
    ws.getRow(r1 - 1).height = 36;

    mes.filas.forEach((fp, i) => {
      const r = r1 + i;
      const p = fp.persona;
      const b = filaBloque(i);
      const S = (c: number) => `${col(c)}${r}`;
      const neto = `(${S(RES.favor)}-${S(RES.contra)}-${S(RES.compensado)})`;
      const Tr = refRango(R_TIPO), Fr = refRango(R_FECHA);
      const ING = `'${HOJAS.asis}'!$${L(PAR.ing)}$${b}`, JOR = `'${HOJAS.asis}'!$${L(PAR.jor)}$${b}`;
      const vals: [number, ExcelJS.CellValue, string?][] = [
        [1, i + 1],
        [RES.nombre, { text: p.alias, hyperlink: `#'${HOJAS.asis}'!A${b}` }],
        [RES.empresa, p.empresa || null],
        [RES.faltas, { formula: `COUNTIF(${refFila('motivo', i)},"Falta")`, result: fp.tot.faltas }, '0;;""'],
        [RES.trabajado, { formula: `SUM(${refFila('trabajado', i)})`, result: fp.tot.trabajado / 1440 }, '[h]:mm'],
        [RES.base, { formula: `(COUNTIFS(${Tr},"LV",${Fr},">="&${ING})+COUNTIFS(${Tr},"SAB",${Fr},">="&${ING}))*${JOR}/1440`, result: fp.tot.base / 1440 }, '[h]:mm'],
        [RES.favor, { formula: `SUM(${refFila('favor', i)})`, result: fp.tot.favor / 1440 }, '[h]:mm'],
        [RES.contra, { formula: `SUM(${refFila('contra', i)})`, result: fp.tot.contra / 1440 }, '[h]:mm'],
        [RES.compensado, { formula: `SUM(${refFila('compensado', i)})`, result: fp.tot.compensado / 1440 }, '[h]:mm'],
        [RES.saldo, { formula: saldoTxt(neto), result: saldoRes(fp.tot.saldo) }],
        [RES.pendientes, { formula: `'${HOJAS.asis}'!${pendTotal(i)}`, result: fp.tot.pendientes }, '0;;"✔"'],
      ];
      for (const [c, v, fmt] of vals) {
        const cell = ws.getCell(r, c);
        cell.value = v;
        if (fmt) cell.numFmt = fmt;
        cell.border = bordes;
        cell.alignment = { horizontal: c === RES.nombre ? 'left' : 'center' };
      }
      ws.getCell(r, RES.nombre).font = { bold: true, color: { argb: 'FF1F4E79' }, underline: true };
      ws.getCell(r, RES.compensado).font = { bold: true, color: { argb: MORADO } };
      ws.getCell(r, RES.saldo).font = { bold: true, size: 12 };
    });
    // Totales
    ws.getCell(rT, RES.nombre).value = 'TOTAL';
    ws.getCell(rT, RES.nombre).font = { bold: true };
    for (let c: number = RES.faltas; c <= RES.pendientes; c++) {
      const cell = ws.getCell(rT, c);
      if (c === RES.saldo) {
        cell.value = { formula: saldoTxt(`(${col(RES.favor)}${rT}-${col(RES.contra)}${rT}-${col(RES.compensado)}${rT})`) };
      } else {
        cell.value = { formula: mes.filas.length ? `SUM(${col(c)}${r1}:${col(c)}${rT - 1})` : '0' };
        cell.numFmt = ws.getCell(r1, c).numFmt;
      }
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
      cell.border = { top: { style: 'double' }, bottom: thin };
    }
    if (mes.filas.length) {
      const rr = (c: number) => `${col(c)}${r1}:${col(c)}${rT}`;
      const fav = `$${col(RES.favor)}${r1}`, con = `$${col(RES.contra)}${r1}`, cmp = `$${col(RES.compensado)}${r1}`;
      ws.addConditionalFormatting({ ref: rr(RES.faltas), rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style: { font: { bold: true, color: { argb: ROJO_TXT } } } }] });
      ws.addConditionalFormatting({ ref: rr(RES.favor), rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style: { font: { bold: true, color: { argb: VERDE_TXT } }, fill: cfFill(VERDE) } }] });
      ws.addConditionalFormatting({ ref: rr(RES.contra), rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style: { font: { bold: true, color: { argb: ROJO_TXT } }, fill: cfFill(ROSA) } }] });
      ws.addConditionalFormatting({ ref: rr(RES.compensado), rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style: { fill: cfFill('FFE4DFEC') } }] });
      ws.addConditionalFormatting({
        ref: rr(RES.saldo),
        rules: [
          { type: 'expression', priority: 1, formulae: [`${fav}-${con}-${cmp}>0`], style: { fill: cfFill('FFC6EFCE'), font: { bold: true, color: { argb: 'FF006100' } } } },
          { type: 'expression', priority: 2, formulae: [`${fav}-${con}-${cmp}<0`], style: { fill: cfFill(ROSA), font: { bold: true, color: { argb: ROJO_TXT } } } },
        ],
      });
      ws.addConditionalFormatting({ ref: rr(RES.pendientes), rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style: { fill: cfFill(AMARILLO), font: { bold: true } } }] });
      ws.autoFilter = { from: { row: r1 - 1, column: 1 }, to: { row: rT - 1, column: RES.pendientes } };
    }
    ws.getCell(rT + 2, RES.nombre).value = 'SALDO del mes = horas a favor − horas en contra − horas compensadas. Verde (+) = le debemos horas; rojo (−) = nos debe horas.';
    ws.getCell(rT + 2, RES.nombre).font = { italic: true, size: 9, color: { argb: GRIS_OSC } };
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
    ws.getCell('C2').value = 'Semanas de lunes a sábado. Se calcula solo desde la hoja Asistencia.';
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
    const ANCHO = 5;
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
      (['Trabajadas', 'A favor', 'En contra', 'Compensadas', 'Saldo'] as const).forEach((t, j) => {
        const c = ws.getCell(r1 - 1, a + j);
        c.value = t;
        estiloCabecera(c);
        if (j === 1) c.fill = fill('FF548235');
        if (j === 2) c.fill = fill('FFC00000');
        if (j === 3) c.fill = fill(MORADO);
        ws.getColumn(a + j).width = 10;
      });
      mes.filas.forEach((fp, i) => {
        const r = r1 + i;
        const dd = fp.dias.filter((_, j) => mes.fechas[j].semana === k);
        const sum = (f: (x: (typeof dd)[number]) => number) => dd.reduce((acc, x) => acc + f(x), 0);
        const vals: [keyof typeof FILA, number, string][] = [
          ['trabajado', sum((x) => x.calc.trabajado), 'FF000000'],
          ['favor', sum((x) => x.calc.favor), VERDE_TXT],
          ['contra', sum((x) => x.calc.contra), ROJO_TXT],
          ['compensado', sum((x) => x.calc.compensado), MORADO],
        ];
        vals.forEach(([kk, res, color], j) => {
          const c = ws.getCell(r, a + j);
          c.value = { formula: `SUMIF(${semRango},${k},${refFila(kk, i)})`, result: res / 1440 };
          c.numFmt = j === 0 ? '[h]:mm' : '[h]:mm;;""';
          c.border = bordes;
          c.alignment = { horizontal: 'center' };
          c.font = { bold: j > 0, color: { argb: color } };
        });
        const net = vals[1][1] - vals[2][1] - vals[3][1];
        const cn = ws.getCell(r, a + 4);
        cn.value = { formula: saldoTxt(`(${col(a + 1)}${r}-${col(a + 2)}${r}-${col(a + 3)}${r})`), result: saldoRes(net) };
        cn.font = { bold: true };
        cn.border = { ...bordes, right: grueso };
        cn.alignment = { horizontal: 'center' };
      });
      const exp = (r: number) => `(${col(a + 1)}${r}-${col(a + 2)}${r}-${col(a + 3)}${r})`;
      if (mes.filas.length) {
        ws.addConditionalFormatting({
          ref: `${col(a + 4)}${r1}:${col(a + 4)}${r1 + mes.filas.length - 1}`,
          rules: [
            { type: 'expression', priority: 1, formulae: [`${exp(r1)}>0`], style: { fill: cfFill('FFC6EFCE'), font: { color: { argb: 'FF006100' } } } },
            { type: 'expression', priority: 2, formulae: [`${exp(r1)}<0`], style: { fill: cfFill(ROSA), font: { color: { argb: ROJO_TXT } } } },
          ],
        });
      }
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
    ws.getColumn(3).width = 95;
    const lineas: [string, ('t' | 'h' | 'c')?, string?][] = [
      [`CONTROL DE ASISTENCIA ${titulo}`, 't'],
      [`Período ${periodo}.`],
      [''],
      ['Cómo trabajar', 'h'],
      ['1. En la hoja "Asistencia" completa las celdas AMARILLAS: escribe la hora (18:05) o elige un motivo.'],
      ['2. Arriba de la hoja dice cuántas faltan. Cuando diga "✔ Todo completo", revisa la hoja "Resumen".'],
      ['3. Si pasas el mouse por una celda gris con triángulo rojo (fila Reloj), verás qué pasó ese día.'],
      [''],
      ['Motivos', 'h'],
      ['Vacaciones / Salud: se cuenta como día completo trabajado.'],
      ['Compensado: la persona usa sus horas a favor (sale antes o falta el día). Va a "Horas compensadas".'],
      ['Estudio: lo que no trabajó se cuenta como horas en contra.'],
      ['Falta: el día completo es hora en contra y se cuenta como falta.'],
      [''],
      ['Colores', 'h'],
      ['Falta un dato.', 'c', AMARILLO],
      ['Entrada.', 'c', VERDE_ENTRADA],
      ['Salida.', 'c', NARANJA_SALIDA],
      ['Hora escrita a mano (no es la del reloj).', 'c', AZUL],
      ['Revisar: la salida es antes que la entrada. ¿Escribiste 6:14 en vez de 18:14?', 'c', 'FFFF7C80'],
      ['Calculado solo. No se escribe.', 'c', GRIS],
      [''],
      ['Reglas', 'h'],
      ['• Tardanza (después de las 08:30) y salida antes de hora: horas en contra.'],
      ['• Llegar antes de las 08:00: horas a favor (desde la llegada hasta las 08:30).'],
      [`• Extras: si sale a las 18:${String(MARG).padStart(2, '0')} o después, se cuentan desde las 18:00 (sábado desde las 12:00).`],
      ['• Domingo o feriado trabajado: todo el tiempo es a favor.'],
      ['• SALDO del mes = a favor − en contra − compensadas.'],
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
