/**
 * Genera el Excel de trabajo de Yasmín.
 *
 * Reparto del trabajo:
 *  - La página decide la ENTRADA y la SALIDA de cada día (o las deja vacías si hay duda).
 *  - Todo lo demás son fórmulas simples dentro del Excel: si Yasmín cambia una hora,
 *    se recalculan solos los totales, la hoja Semanal y el Resumen.
 *
 * Las fórmulas replican exactamente calcularDia() de rules.ts (hay una prueba que lo verifica
 * recalculando el archivo con LibreOffice).
 */
import ExcelJS from 'exceljs';
import type { Config } from './config';
import type { Mes } from './model';
import type { Reporte } from './parser';
import { DIAS_CORTO, MESES, ddmm, excelSerial, fmtHHMM, parseISO, weekday, type ISODate } from './time';

// ---------- estilo ----------
const ROJO = 'FFC8102E';
const ROJO_SUAVE = 'FFFBE3E6';
const GRIS = 'FFF2F2F2';
const GRIS_OSC = 'FF7F7F7F';
const GRIS_RELOJ = 'FFE7E6E6';
const AMARILLO = 'FFFFEB3B';
const AZUL = 'FFDDEBF7';
const VERDE = 'FFE2F0D9';
const NARANJA = 'FFFCE4D6';
const BLANCO = 'FFFFFFFF';

const fill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const cfFill = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', bgColor: { argb } });
const thin: Partial<ExcelJS.Border> = { style: 'thin', color: { argb: 'FFBFBFBF' } };
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
export const HOJAS = {
  leeme: 'Leeme',
  revisar: 'Por revisar',
  asis: 'Asistencia',
  semanal: 'Semanal',
  resumen: 'Resumen',
  config: 'Configuracion',
  reloj: 'Reloj original',
} as const;

/** Filas dentro del bloque de cada persona en la hoja Asistencia. */
export const FILA = { reloj: 0, entrada: 1, salida: 2, motivo: 3, trabajado: 4, favor: 5, contra: 6, estado: 7 } as const;
export const ALTO_BLOQUE = 8;
export const PRIMERA_FILA_BLOQUE = 10;
export const PRIMERA_COL_DIA = 3; // C
const R_DIASEM = 5, R_FECHA = 6, R_TIPO = 7, R_SEMANA = 8;

export function filaBloque(i: number): number {
  return PRIMERA_FILA_BLOQUE + i * ALTO_BLOQUE;
}

/** Columnas del Resumen (para pruebas y lectura de saldos). */
export const RES = {
  primeraFila: 6,
  nombre: 2, empresa: 3, dias: 4, faltas: 5, trabajado: 6, favor: 7, enFeriado: 8, contra: 9,
  antFavor: 10, antContra: 11, castigo: 12, pagadas: 13, saldoFavor: 14, saldoContra: 15, pendientes: 16, alertas: 17,
} as const;

// Columnas de la tabla de personas en Configuracion (índice para BUSCARV)
const PC = { alias: 1, reloj: 2, id: 3, empresa: 4, horas: 5, entrada: 6, salidaLV: 7, salidaSab: 8, extrasLV: 9, temprano: 10, descuenta: 11, ingreso: 12, reparto: 13 } as const;
const CFG_FILA_PERSONAS = 13;

const MOTIVO_TXT_FALTA = 'FALTA DATO';
const TXT_REVISAR = 'REVISAR HORAS';

export interface OpcionesExcel {
  logoBase64?: string; // sin el prefijo data:
  generado?: Date;
  /** Para pruebas: no guardar los resultados precalculados, así el recálculo es obligatorio. */
  sinResultados?: boolean;
}

export async function generarExcel(mes: Mes, cfg: Config, rep: Reporte, op: OpcionesExcel = {}): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Control de asistencia QALEN';
  wb.created = op.generado ?? new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  const logoId = op.logoBase64 ? wb.addImage({ base64: op.logoBase64, extension: 'png' }) : null;

  const { y, mo } = parseISO(mes.inicio);
  const titulo = `${MESES[mo]} ${y}`;
  const periodo = `del ${ddmm(mes.inicio)} al ${ddmm(mes.hasta)}`;

  const leeme = wb.addWorksheet(HOJAS.leeme, { properties: { tabColor: { argb: GRIS_OSC } } });
  const revisar = wb.addWorksheet(HOJAS.revisar, { properties: { tabColor: { argb: AMARILLO } } });
  const asis = wb.addWorksheet(HOJAS.asis, { properties: { tabColor: { argb: ROJO } } });
  const semanal = wb.addWorksheet(HOJAS.semanal, { properties: { tabColor: { argb: 'FF5B9BD5' } } });
  const resumen = wb.addWorksheet(HOJAS.resumen, { properties: { tabColor: { argb: 'FF70AD47' } } });
  const conf = wb.addWorksheet(HOJAS.config, { properties: { tabColor: { argb: GRIS_OSC } } });
  const reloj = wb.addWorksheet(HOJAS.reloj, { properties: { tabColor: { argb: GRIS_OSC } } });

  const ponerLogo = (ws: ExcelJS.Worksheet, ancho = 150) => {
    if (logoId != null) ws.addImage(logoId, { tl: { col: 0.1, row: 0.15 }, ext: { width: ancho, height: Math.round((ancho * 208) / 660) } });
  };

  // =====================================================================
  // CONFIGURACION
  // =====================================================================
  {
    const ws = conf;
    ws.getColumn(1).width = 16;
    ws.getColumn(2).width = 22;
    for (let c = 3; c <= 13; c++) ws.getColumn(c).width = 14;
    ws.getColumn(15).width = 13;
    ws.getColumn(16).width = 34;
    ws.getCell('A1').value = 'CONFIGURACIÓN';
    ws.getCell('A1').font = { bold: true, size: 14, color: { argb: ROJO } };
    ws.getCell('A2').value = 'Lo que cambies aquí se aplica en todo el archivo (Asistencia, Semanal y Resumen). Las celdas blancas con borde se pueden editar.';
    ws.getCell('A2').font = { italic: true, color: { argb: GRIS_OSC } };

    const reglas: [string, number, string, string][] = [
      ['MargenExtras', cfg.reglas.margenExtras, 'Margen para extras (minutos)', 'Las extras cuentan solo si sale este tiempo o más después de su hora (20 → desde las 18:20). Se cuentan desde la hora de salida.'],
      ['Almuerzo', cfg.reglas.almuerzoMin, 'Almuerzo L-V (minutos)', 'Se descuenta de las horas trabajadas de lunes a viernes. No afecta la bolsa de horas.'],
      ['TardanzaDesde', cfg.reglas.tardanzaDesde, 'Tardanza que resta (minutos)', 'La tardanza se resta de la bolsa solo si es de este tiempo o más. 0 = toda tardanza resta.'],
      ['FactorFeriado', cfg.reglas.factorFeriado, 'Factor domingo / feriado', '1 = cada hora trabajada en domingo o feriado suma 1 hora a favor; 2 = suma doble.'],
    ];
    ws.getCell('A4').value = 'REGLAS GENERALES';
    ws.getCell('A4').font = { bold: true, color: { argb: ROJO } };
    reglas.forEach(([nombre, valor, etiqueta, ayuda], i) => {
      const r = 5 + i;
      ws.getCell(r, 1).value = etiqueta;
      ws.mergeCells(r, 1, r, 2);
      const c = ws.getCell(r, 3);
      c.value = valor;
      c.border = bordes;
      c.protection = { locked: false };
      c.numFmt = '0.##';
      ws.getCell(r, 4).value = ayuda;
      ws.getCell(r, 4).font = { italic: true, size: 9, color: { argb: GRIS_OSC } };
      wb.definedNames.add(`${HOJAS.config}!$C$${r}`, nombre);
      c.dataValidation = { type: 'decimal', operator: 'between', formulae: [0, 600], allowBlank: false, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Valor no válido', error: 'Escribe un número (sin letras).' };
    });

    // Personas
    const rh = CFG_FILA_PERSONAS - 1;
    ws.getCell(rh - 1, 1).value = 'PERSONAS';
    ws.getCell(rh - 1, 1).font = { bold: true, color: { argb: ROJO } };
    const cab = ['Nombre corto', 'Nombre en el reloj', 'ID del reloj', 'Empresa', 'Horas semanales', 'Entrada', 'Salida L-V', 'Salida sábado', 'Extras desde (L-V)', 'Llegar temprano cuenta si llega antes de', 'Descuenta tardanzas, salidas y permisos', 'Fecha de ingreso', 'Reparto'];
    cab.forEach((t, i) => {
      const c = ws.getCell(rh, i + 1);
      c.value = t;
      estiloCabecera(c);
    });
    ws.getRow(rh).height = 42;
    const activos = mes.filas.map((f) => f.persona);
    activos.forEach((p, i) => {
      const r = CFG_FILA_PERSONAS + i;
      const vals: ExcelJS.CellValue[] = [
        p.alias, p.nombreReloj, p.id, p.empresa, p.horasSemanales,
        p.entrada / 1440, p.salidaLV / 1440, p.salidaSab / 1440, p.extrasDesdeLV / 1440,
        p.tempranoAntesDe != null ? p.tempranoAntesDe / 1440 : null,
        p.descuenta ? 'Sí' : 'No',
        p.fechaIngreso ? excelSerial(p.fechaIngreso) : null,
        p.reparto ? 'Sí' : 'No',
      ];
      vals.forEach((v, j) => {
        const c = ws.getCell(r, j + 1);
        c.value = v;
        c.border = bordes;
        if (j >= 5 && j <= 9) c.numFmt = 'hh:mm';
        if (j === 11) c.numFmt = 'dd/mm/yyyy';
        if (j === 0) {
          c.font = { bold: true };
          c.fill = fill(GRIS);
          c.note = 'El nombre corto une esta fila con la hoja Asistencia. No lo cambies aquí: cámbialo en la página y vuelve a generar el archivo.';
        } else if (j === 1 || j === 2) {
          c.fill = fill(GRIS);
        } else {
          c.protection = { locked: false };
        }
        if (j >= 5 && j <= 9) c.dataValidation = { type: 'decimal', operator: 'between', formulae: [0, 0.99999], allowBlank: j === 9, showErrorMessage: true, errorStyle: 'stop', error: 'Escribe una hora, por ejemplo 08:30 o 18:00.' };
        if (j === 10 || j === 12) c.dataValidation = { type: 'list', formulae: ['"Sí,No"'], allowBlank: false, showErrorMessage: true, errorStyle: 'stop', error: 'Elige Sí o No.' };
        if (j === 4) c.dataValidation = { type: 'decimal', operator: 'between', formulae: [1, 72], showErrorMessage: true, errorStyle: 'stop', error: 'Horas por semana: entre 1 y 72.' };
      });
    });
    const finPersonas = CFG_FILA_PERSONAS + Math.max(activos.length, 1) - 1;
    wb.definedNames.add(`${HOJAS.config}!$A$${CFG_FILA_PERSONAS}:$M$${finPersonas}`, 'Personas');

    // Motivos
    const rm = finPersonas + 3;
    ws.getCell(rm - 1, 1).value = 'MOTIVOS (lista desplegable de la fila "Motivo")';
    ws.getCell(rm - 1, 1).font = { bold: true, color: { argb: ROJO } };
    ['Motivo', 'Si no marcó, resta la jornada', 'Igual hay que escribir entrada y salida'].forEach((t, i) => {
      const c = ws.getCell(rm, i + 1);
      c.value = t;
      estiloCabecera(c);
    });
    ws.getRow(rm).height = 32;
    const nMot = cfg.motivos.length + 4; // filas libres para agregar motivos
    for (let i = 0; i < nMot; i++) {
      const m = cfg.motivos[i];
      const r = rm + 1 + i;
      const vals = m ? [m.nombre, m.restaJornada ? 'Sí' : 'No', m.necesitaHoras ? 'Sí' : 'No'] : [null, null, null];
      vals.forEach((v, j) => {
        const c = ws.getCell(r, j + 1);
        c.value = v;
        c.border = bordes;
        c.protection = { locked: false };
        if (j > 0) c.dataValidation = { type: 'list', formulae: ['"Sí,No"'], allowBlank: true, showErrorMessage: true, errorStyle: 'stop', error: 'Elige Sí o No.' };
      });
    }
    wb.definedNames.add(`${HOJAS.config}!$A$${rm + 1}:$C$${rm + nMot}`, 'Motivos');
    wb.definedNames.add(`${HOJAS.config}!$A$${rm + 1}:$A$${rm + nMot}`, 'ListaMotivos');

    // Feriados
    ws.getCell(rh - 1, 15).value = 'FERIADOS';
    ws.getCell(rh - 1, 15).font = { bold: true, color: { argb: ROJO } };
    ws.getCell(rh - 2, 15).value = 'Verifica cada año el calendario oficial.';
    ws.getCell(rh - 2, 15).font = { italic: true, size: 9, color: { argb: GRIS_OSC } };
    ['Fecha', 'Feriado'].forEach((t, i) => {
      const c = ws.getCell(rh, 15 + i);
      c.value = t;
      estiloCabecera(c);
    });
    const fer = [...cfg.feriados].sort((a, b) => a.fecha.localeCompare(b.fecha)).filter((f) => f.fecha.slice(0, 4) === String(y));
    const nFer = Math.max(fer.length + 10, 30);
    for (let i = 0; i < nFer; i++) {
      const r = rh + 1 + i;
      const f = fer[i];
      const c1 = ws.getCell(r, 15);
      const c2 = ws.getCell(r, 16);
      c1.value = f ? excelSerial(f.fecha) : null;
      c1.numFmt = 'dd/mm/yyyy';
      c2.value = f?.nombre ?? null;
      for (const c of [c1, c2]) {
        c.border = bordes;
        c.protection = { locked: false };
      }
      c1.dataValidation = { type: 'date', operator: 'greaterThan', formulae: [new Date(Date.UTC(2000, 0, 1))], allowBlank: true, showErrorMessage: true, errorStyle: 'stop', error: 'Escribe una fecha, por ejemplo 08/10/2026.' };
    }
    wb.definedNames.add(`${HOJAS.config}!$O$${rh + 1}:$O$${rh + nFer}`, 'Feriados');
    ws.views = [{ state: 'frozen', ySplit: rh, xSplit: 1 }];
    await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true, formatCells: true, objects: true });
  }

  // =====================================================================
  // ASISTENCIA
  // =====================================================================
  const nD = mes.fechas.length;
  const cDia = (i: number) => PRIMERA_COL_DIA + i;
  const cUlt = cDia(nD - 1);
  const cTot = cUlt + 1;
  const cPar = cTot + 2; // parámetros por persona
  const PAR = { ent: cPar, sLV: cPar + 1, sSab: cPar + 2, xLV: cPar + 3, temp: cPar + 4, desc: cPar + 5, jor: cPar + 6, ing: cPar + 7 };
  const L = (c: number) => col(c);
  const rangoDias = (r: number) => `$${L(cDia(0))}$${r}:$${L(cUlt)}$${r}`;

  {
    const ws = asis;
    ponerLogo(ws);
    ws.getColumn(1).width = 13;
    ws.getColumn(2).width = 13;
    for (let i = 0; i < nD; i++) ws.getColumn(cDia(i)).width = 7.3;
    ws.getColumn(cTot).width = 10;
    ws.getColumn(cTot + 1).width = 3;
    for (let c = cPar; c <= PAR.ing; c++) {
      ws.getColumn(c).width = 9;
      ws.getColumn(c).outlineLevel = 1;
      ws.getColumn(c).hidden = true;
    }

    const t = ws.getCell(1, 4);
    t.value = `ASISTENCIA ${titulo}`;
    t.font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell(2, 4).value = `Período ${periodo}. Reporte del reloj hasta el ${ddmm(mes.finReporte)}.`;
    ws.getCell(2, 4).font = { color: { argb: GRIS_OSC } };
    const ban = ws.getCell(3, 4);
    const totPend = `COUNTIF(${rangoBloques('estado')},"${MOTIVO_TXT_FALTA}")`;
    const totAlert = `COUNTIF(${rangoBloques('estado')},"${TXT_REVISAR}")`;
    ban.value = { formula: `IF(${totPend}+${totAlert}=0,"✔ Todo completo.","⚠ Faltan "&${totPend}&" datos (celdas amarillas)"&IF(${totAlert}>0," y hay "&${totAlert}&" horas por revisar (en rojo)","")&". La lista está en la hoja 'Por revisar'.")` };
    ban.font = { bold: true, size: 12, color: { argb: 'FF9C0006' } };

    // Cabeceras de días
    for (const [r, txt] of [[R_DIASEM, 'Día'], [R_FECHA, 'Fecha'], [R_TIPO, 'Tipo'], [R_SEMANA, 'Semana']] as const) {
      const c = ws.getCell(r, 2);
      c.value = txt;
      c.font = { bold: true, size: 9, color: { argb: GRIS_OSC } };
      c.alignment = { horizontal: 'right' };
    }
    mes.fechas.forEach((f, i) => {
      const c = cDia(i);
      const L1 = L(c);
      const w = weekday(f.fecha);
      const h1 = ws.getCell(R_DIASEM, c);
      h1.value = DIAS_CORTO[w];
      const h2 = ws.getCell(R_FECHA, c);
      h2.value = excelSerial(f.fecha);
      h2.numFmt = 'dd/mm';
      const h3 = ws.getCell(R_TIPO, c);
      h3.value = { formula: `IF(COUNTIF(Feriados,${L1}$${R_FECHA})>0,"FER",IF(WEEKDAY(${L1}$${R_FECHA},2)=7,"DOM",IF(WEEKDAY(${L1}$${R_FECHA},2)=6,"SAB","LV")))`, result: f.tipo };
      const h4 = ws.getCell(R_SEMANA, c);
      h4.value = f.semana;
      h4.numFmt = '"Sem "0';
      for (const h of [h1, h2, h3, h4]) {
        h.alignment = { horizontal: 'center' };
        h.font = { bold: h !== h4 && h !== h3, size: h === h2 ? 11 : 9, color: { argb: h === h2 ? 'FFFFFFFF' : GRIS_OSC } };
      }
      h2.fill = fill(f.tipo === 'LV' ? ROJO : f.tipo === 'SAB' ? 'FFE0707F' : GRIS_OSC);
    });
    // Feriados/domingos resaltados en la cabecera (dinámico si se agrega un feriado)
    ws.addConditionalFormatting({
      ref: `${L(cDia(0))}${R_DIASEM}:${L(cUlt)}${R_TIPO}`,
      rules: [{ type: 'expression', priority: 1, formulae: [`OR(${L(cDia(0))}$${R_TIPO}="FER",${L(cDia(0))}$${R_TIPO}="DOM")`], style: { fill: cfFill('FFFFC7CE'), font: { color: { argb: 'FF9C0006' }, bold: true } } }],
    });
    const tc = ws.getCell(R_FECHA, cTot);
    tc.value = 'TOTAL';
    estiloCabecera(tc);
    const pc = ws.getCell(R_FECHA, cPar);
    pc.value = 'Datos de Configuración (no tocar)';
    pc.font = { italic: true, size: 9, color: { argb: GRIS_OSC } };
    ['Entrada', 'Salida L-V', 'Salida sáb', 'Extras L-V', 'Temprano', 'Descuenta', 'Jornada', 'Ingreso'].forEach((txt, j) => {
      const c = ws.getCell(R_SEMANA, cPar + j);
      c.value = txt;
      c.font = { size: 8, color: { argb: GRIS_OSC } };
    });

    // Bloques por persona
    mes.filas.forEach((fp, i) => {
      const r0 = filaBloque(i);
      const p = fp.persona;
      const R = (k: keyof typeof FILA) => r0 + FILA[k];
      ws.mergeCells(r0, 1, r0 + ALTO_BLOQUE - 1, 1);
      const nc = ws.getCell(r0, 1);
      nc.value = p.alias;
      nc.font = { bold: true, size: 11 };
      nc.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      nc.fill = fill(GRIS);
      if (p.empresa) nc.note = `Empresa: ${p.empresa}`;

      const etiquetas: Record<keyof typeof FILA, string> = {
        reloj: 'Reloj', entrada: 'Entrada', salida: 'Salida', motivo: 'Motivo / Obs.', trabajado: 'Horas trabajadas', favor: 'A favor', contra: 'En contra', estado: 'Estado',
      };
      for (const k of Object.keys(FILA) as (keyof typeof FILA)[]) {
        const c = ws.getCell(R(k), 2);
        c.value = etiquetas[k];
        c.font = { bold: k === 'entrada' || k === 'salida' || k === 'favor' || k === 'contra', size: 9 };
        c.fill = fill(k === 'entrada' || k === 'salida' || k === 'motivo' ? BLANCO : GRIS);
      }
      ws.getRow(R('reloj')).height = 24;

      // Parámetros (BUSCARV a Configuracion)
      const A = `$A$${r0}`;
      const bus = (k: number) => `VLOOKUP(${A},Personas,${k},FALSE)`;
      const par: [number, string, number][] = [
        [PAR.ent, `ROUND(${bus(PC.entrada)}*1440,0)`, p.entrada],
        [PAR.sLV, `ROUND(${bus(PC.salidaLV)}*1440,0)`, p.salidaLV],
        [PAR.sSab, `ROUND(${bus(PC.salidaSab)}*1440,0)`, p.salidaSab],
        [PAR.xLV, `ROUND(${bus(PC.extrasLV)}*1440,0)`, p.extrasDesdeLV],
        [PAR.temp, `IF(${bus(PC.temprano)}="",0,ROUND(${bus(PC.temprano)}*1440,0))`, p.tempranoAntesDe ?? 0],
        [PAR.desc, `IF(${bus(PC.descuenta)}="Sí",1,0)`, p.descuenta ? 1 : 0],
        [PAR.jor, `ROUND(${bus(PC.horas)}*60/6,0)`, Math.round((p.horasSemanales * 60) / 6)],
        [PAR.ing, `IF(${bus(PC.ingreso)}="",0,${bus(PC.ingreso)})`, p.fechaIngreso ? excelSerial(p.fechaIngreso) : 0],
      ];
      for (const [c, f, res] of par) {
        const cell = ws.getCell(r0, c);
        cell.value = { formula: f, result: res };
        cell.font = { size: 8, color: { argb: GRIS_OSC } };
      }
      const P = (c: number) => `$${L(c)}$${r0}`;
      const IN = P(PAR.ent), OLV = P(PAR.sLV), OSAB = P(PAR.sSab), XLV = P(PAR.xLV), TEMP = P(PAR.temp), DESC = P(PAR.desc), JOR = P(PAR.jor), ING = P(PAR.ing);

      fp.dias.forEach((d, j) => {
        const c = cDia(j);
        const X = L(c);
        const E = `${X}${R('entrada')}`, S = `${X}${R('salida')}`, M = `${X}${R('motivo')}`;
        const T = `${X}$${R_TIPO}`, F = `${X}$${R_FECHA}`;
        const e = `ROUND(${E}*1440,0)`, s = `ROUND(${S}*1440,0)`;
        const fin = `IF(${T}="SAB",${OSAB},${OLV})`;
        const xdesde = `IF(${T}="SAB",${OSAB},${XLV})`;

        const cr = ws.getCell(R('reloj'), c);
        cr.value = d.reloj || null;
        cr.font = { size: 8, color: { argb: 'FF404040' } };
        cr.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
        cr.fill = fill(GRIS_RELOJ);
        const notas = d.casos.map((x) => `• ${x.mensaje}${x.sugerencia ? ` Sugerencia: ${x.sugerencia}.` : ''}`);
        if (d.clas.duplicadas.length) notas.push(`• Se quitaron marcadas repetidas: ${d.clas.duplicadas.map(fmtHHMM).join(', ')}.`);
        if (d.clas.almuerzoSale != null) notas.push(`Almuerzo: ${fmtHHMM(d.clas.almuerzoSale)}${d.clas.almuerzoVuelve != null ? ` – ${fmtHHMM(d.clas.almuerzoVuelve)}` : ''}.`);
        if (d.clas.otras.length) notas.push(`Otras marcadas: ${d.clas.otras.map(fmtHHMM).join(', ')}.`);
        if (notas.length && d.aplica) cr.note = notas.join('\n');

        const ce = ws.getCell(R('entrada'), c);
        const cs = ws.getCell(R('salida'), c);
        const cm = ws.getCell(R('motivo'), c);
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
        cm.font = { size: 8 };
        cm.alignment = { horizontal: 'center', shrinkToFit: true };

        const trab = `IF(${F}<${ING},0,IF(ISNUMBER(${E}),IF(ISNUMBER(${S}),IF(${s}>${e},${s}-${e}-IF(AND(${T}="LV",${s}-${e}>300),Almuerzo,0),0),0),0))/1440`;
        const favor =
          `IF(${F}<${ING},0,IF(OR(${T}="DOM",${T}="FER"),` +
          `IF(ISNUMBER(${E}),IF(ISNUMBER(${S}),IF(${s}>${e},ROUND((${s}-${e})*FactorFeriado,0),0),0),0),` +
          `IF(ISNUMBER(${S}),IF(${s}>=${xdesde}+MargenExtras,${s}-${xdesde},0),0)` +
          `+IF(ISNUMBER(${E}),IF(AND(${TEMP}>0,${e}<${TEMP},${e}<${IN}),${IN}-${e},0),0)))/1440`;
        const contra =
          `IF(OR(${F}<${ING},${T}="DOM",${T}="FER",${DESC}=0),0,` +
          `IF(ISNUMBER(${E}),IF(AND(${e}>${IN},${e}-${IN}>=TardanzaDesde),${e}-${IN},0),0)` +
          `+IF(ISNUMBER(${S}),IF(${s}<${fin},${fin}-${s},0),0)` +
          `+IF(ISNUMBER(${E}),0,IF(ISNUMBER(${S}),0,IF(IFERROR(VLOOKUP(${M},Motivos,2,FALSE)="Sí",FALSE),MIN(${JOR},${fin}-${IN}),0))))/1440`;
        const estado =
          `IF(${F}<${ING},"",IF(AND(ISNUMBER(${E}),ISNUMBER(${S})),IF(${s}<=${e},"${TXT_REVISAR}",""),` +
          `IF(OR(${T}="DOM",${T}="FER"),IF(ISNUMBER(${E})<>ISNUMBER(${S}),"${MOTIVO_TXT_FALTA}",""),` +
          `IF(${M}="","${MOTIVO_TXT_FALTA}",IF(IFERROR(VLOOKUP(${M},Motivos,3,FALSE)="Sí",FALSE),"${MOTIVO_TXT_FALTA}","")))))`;

        const ct = ws.getCell(R('trabajado'), c);
        const cf = ws.getCell(R('favor'), c);
        const cc = ws.getCell(R('contra'), c);
        const cst = ws.getCell(R('estado'), c);
        ct.value = { formula: trab, result: d.calc.trabajado / 1440 };
        cf.value = { formula: favor, result: d.calc.favor / 1440 };
        cc.value = { formula: contra, result: d.calc.contra / 1440 };
        cst.value = { formula: estado, result: d.calc.alerta ? TXT_REVISAR : d.calc.pendiente ? MOTIVO_TXT_FALTA : '' };
        for (const x of [ct, cf, cc]) {
          x.numFmt = '[h]:mm;;""';
          x.fill = fill(GRIS);
          x.alignment = { horizontal: 'center' };
          x.font = { size: 9 };
        }
        cf.font = { size: 9, bold: true, color: { argb: 'FF375623' } };
        cc.font = { size: 9, bold: true, color: { argb: 'FF9C0006' } };
        cst.font = { size: 7, bold: true, color: { argb: 'FF9C0006' } };
        cst.alignment = { horizontal: 'center', shrinkToFit: true };
        cst.fill = fill(GRIS);
      });

      // Totales de la fila
      const tot = (k: keyof typeof FILA, f: string, res: number | string, fmt?: string) => {
        const c = ws.getCell(R(k), cTot);
        c.value = { formula: f, result: res };
        if (fmt) c.numFmt = fmt;
        c.font = { bold: true };
        c.alignment = { horizontal: 'center' };
        c.border = bordes;
        c.fill = fill(GRIS);
      };
      tot('entrada', `COUNT(${rangoDias(R('entrada'))})`, fp.tot.dias, '0" días"');
      tot('trabajado', `SUM(${rangoDias(R('trabajado'))})`, fp.tot.trabajado / 1440, '[h]:mm');
      tot('favor', `SUM(${rangoDias(R('favor'))})`, fp.tot.favor / 1440, '[h]:mm');
      tot('contra', `SUM(${rangoDias(R('contra'))})`, fp.tot.contra / 1440, '[h]:mm');
      tot('estado', `IF(COUNTIF(${rangoDias(R('estado'))},"${MOTIVO_TXT_FALTA}")=0,"✔","Faltan "&COUNTIF(${rangoDias(R('estado'))},"${MOTIVO_TXT_FALTA}"))`, fp.tot.pendientes ? `Faltan ${fp.tot.pendientes}` : '✔');

      // Borde grueso entre personas
      for (let c = 1; c <= cTot; c++) {
        const cell = ws.getCell(r0, c);
        cell.border = { ...(cell.border ?? {}), top: { style: 'medium', color: { argb: ROJO } } };
      }

      // Formatos condicionales del bloque
      const rangoFila = (k: keyof typeof FILA) => `${L(cDia(0))}${R(k)}:${L(cUlt)}${R(k)}`;
      const X0 = L(cDia(0));
      const est = `${X0}${R('estado')}`;
      for (const k of ['entrada', 'salida'] as const) {
        const ref = `${X0}${R(k)}`;
        ws.addConditionalFormatting({
          ref: rangoFila(k),
          rules: [
            { type: 'expression', priority: 1, formulae: [`${X0}$${R_FECHA}<${ING}`], style: { fill: cfFill('FFD9D9D9'), font: { color: { argb: GRIS_OSC } } } },
            { type: 'expression', priority: 2, formulae: [`${est}="${TXT_REVISAR}"`], style: { fill: cfFill('FFFF7C80'), font: { bold: true } } },
            { type: 'expression', priority: 3, formulae: [`AND(${est}="${MOTIVO_TXT_FALTA}",NOT(ISNUMBER(${ref})))`], style: { fill: cfFill(AMARILLO) } },
            // Hora que no está en el reloj = escrita a mano
            { type: 'expression', priority: 4, formulae: [`AND(ISNUMBER(${ref}),ISERROR(SEARCH(TEXT(${ref},"hh:mm"),${X0}${R('reloj')})))`], style: { fill: cfFill(AZUL), font: { italic: true, color: { argb: 'FF1F4E79' } } } },
          ],
        });
      }
      ws.addConditionalFormatting({
        ref: rangoFila('motivo'),
        rules: [
          { type: 'expression', priority: 1, formulae: [`${X0}$${R_FECHA}<${ING}`], style: { fill: cfFill('FFD9D9D9') } },
          { type: 'expression', priority: 2, formulae: [`AND(${est}="${MOTIVO_TXT_FALTA}",${X0}${R('motivo')}="")`], style: { fill: cfFill(AMARILLO) } },
          { type: 'expression', priority: 3, formulae: [`${X0}${R('motivo')}<>""`], style: { fill: cfFill(VERDE) } },
        ],
      });
      ws.addConditionalFormatting({
        ref: `${X0}${R('reloj')}:${L(cUlt)}${R('estado')}`,
        rules: [{ type: 'expression', priority: 10, formulae: [`AND(OR(${X0}$${R_TIPO}="DOM",${X0}$${R_TIPO}="FER"),${X0}${R('reloj')}="")`], style: { fill: cfFill('FFFFE5EA') } }],
      });

      // Validaciones de datos (aviso con confirmación, no bloqueo)
      dv(ws).add(rangoFila('entrada'), {
        type: 'custom', allowBlank: true, formulae: [`OR(${X0}${R('entrada')}="",AND(ISNUMBER(${X0}${R('entrada')}),${X0}${R('entrada')}<1))`],
        showErrorMessage: true, errorStyle: 'warning', errorTitle: '¿Hora correcta?',
        error: 'Escribe la hora así: 8:30 o 18:05 (formato de 24 horas). ¿Quieres dejarlo igual?',
        showInputMessage: false,
      });
      dv(ws).add(rangoFila('salida'), {
        type: 'custom', allowBlank: true,
        formulae: [`OR(${X0}${R('salida')}="",AND(ISNUMBER(${X0}${R('salida')}),${X0}${R('salida')}<1,OR(NOT(ISNUMBER(${X0}${R('entrada')})),${X0}${R('salida')}>${X0}${R('entrada')})))`],
        showErrorMessage: true, errorStyle: 'warning', errorTitle: '¿Hora de salida correcta?',
        error: 'La salida debe ser DESPUÉS de la entrada y en formato de 24 horas.\nEjemplo: 6:14 de la tarde se escribe 18:14.\n¿Quieres dejarlo igual?',
      });
      dv(ws).add(rangoFila('motivo'), {
        type: 'list', allowBlank: true, formulae: ['ListaMotivos'],
        showErrorMessage: false, errorStyle: 'information',
        showInputMessage: true, promptTitle: 'Motivo u observación', prompt: 'Elige de la lista o escribe una nota. Solo los motivos de la lista cambian el cálculo.',
      });
    });

    // Leyenda
    const rLey = filaBloque(mes.filas.length) + 1;
    const ley: [string, string][] = [
      [AMARILLO, 'Falta un dato: escribe la hora o elige un motivo.'],
      [AZUL, 'Hora escrita a mano (no está en el reloj).'],
      ['FFFF7C80', 'Revisar: la salida es antes que la entrada (¿formato de 12 h?).'],
      [VERDE, 'Día con motivo / observación.'],
      [GRIS, 'Calculado con fórmulas: no se escribe aquí.'],
    ];
    ley.forEach(([c, t2], i) => {
      ws.getCell(rLey + i, 2).fill = fill(c);
      ws.getCell(rLey + i, 2).border = bordes;
      ws.getCell(rLey + i, 3).value = t2;
      ws.getCell(rLey + i, 3).font = { size: 9 };
    });

    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: R_SEMANA, activeCell: `${L(cDia(0))}${PRIMERA_FILA_BLOQUE + 1}`, showGridLines: false }];
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true, formatCells: true, objects: true });
  }

  function rangoBloques(k: keyof typeof FILA): string {
    // rango continuo de todas las filas de bloques; COUNTIF solo encontrará el texto en la fila Estado
    if (!mes.filas.length) return `${HOJAS.asis}!$A$1:$A$1`;
    const a = filaBloque(0) + FILA[k];
    const b = filaBloque(mes.filas.length - 1) + FILA[k];
    return `$${L(cDia(0))}$${a}:$${L(cUlt)}$${b}`;
  }

  const ref = (c: number, r: number) => `'${HOJAS.asis}'!${L(c)}${r}`;
  const refRango = (r: number) => `'${HOJAS.asis}'!$${L(cDia(0))}$${r}:$${L(cUlt)}$${r}`;
  const refFila = (k: keyof typeof FILA, i: number) => refRango(filaBloque(i) + FILA[k]);

  // =====================================================================
  // RESUMEN
  // =====================================================================
  {
    const ws = resumen;
    ponerLogo(ws, 130);
    ws.getCell('C1').value = `RESUMEN — BOLSA DE HORAS ${titulo}`;
    ws.getCell('C1').font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell('C2').value = `Período ${periodo}. Todo se calcula desde la hoja Asistencia. Las columnas blancas con borde se pueden editar.`;
    ws.getCell('C2').font = { color: { argb: GRIS_OSC } };
    const r1 = RES.primeraFila;
    const rT = r1 + mes.filas.length;
    const ban = ws.getCell('C3');
    ban.value = { formula: `IF(${col(RES.pendientes)}${rT}+${col(RES.alertas)}${rT}=0,"✔ Sin datos pendientes.","⚠ Todavía faltan "&${col(RES.pendientes)}${rT}&" datos y hay "&${col(RES.alertas)}${rT}&" horas por revisar: el resumen puede cambiar.")` };
    ban.font = { bold: true, size: 12, color: { argb: 'FF9C0006' } };

    const grupos: [number, number, string, string][] = [
      [RES.dias, RES.contra, 'ESTE MES', ROJO],
      [RES.antFavor, RES.pagadas, 'AJUSTES (puedes editarlos)', 'FF7F6000'],
      [RES.saldoFavor, RES.saldoContra, 'SALDO FINAL', 'FF375623'],
    ];
    for (const [a, b, txt, color] of grupos) {
      ws.mergeCells(r1 - 2, a, r1 - 2, b);
      const c = ws.getCell(r1 - 2, a);
      c.value = txt;
      c.alignment = { horizontal: 'center' };
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = fill(color);
    }
    const cab: [number, string, number][] = [
      [1, 'N°', 5], [RES.nombre, 'Nombre', 14], [RES.empresa, 'Empresa', 11], [RES.dias, 'Días trabajados', 9], [RES.faltas, 'Faltas', 7],
      [RES.trabajado, 'Horas trabajadas', 10], [RES.favor, 'A favor (extras)', 10], [RES.enFeriado, 'de ello: domingo / feriado', 11], [RES.contra, 'En contra (compensadas)', 12],
      [RES.antFavor, 'Saldo anterior: le debemos', 11], [RES.antContra, 'Saldo anterior: nos debe', 11], [RES.castigo, 'Castigo', 9], [RES.pagadas, 'Horas pagadas', 9],
      [RES.saldoFavor, 'Le debemos (horas)', 12], [RES.saldoContra, 'Nos debe (horas)', 12], [RES.pendientes, 'Datos por completar', 10], [RES.alertas, 'Horas por revisar', 9],
    ];
    for (const [c, t, w] of cab) {
      const cell = ws.getCell(r1 - 1, c);
      cell.value = t;
      estiloCabecera(cell);
      ws.getColumn(c).width = w;
    }
    ws.getRow(r1 - 1).height = 45;
    const tipos = refRango(R_TIPO);

    mes.filas.forEach((fp, i) => {
      const r = r1 + i;
      const p = fp.persona;
      const ant = p.saldoAnterior;
      const fav = refFila('favor', i);
      const S = (c: number) => `${col(c)}${r}`;
      const neto = `${S(RES.antFavor)}-${S(RES.antContra)}+${S(RES.favor)}-${S(RES.contra)}-${S(RES.castigo)}-${S(RES.pagadas)}`;
      const saldo = (ant + fp.tot.favor - fp.tot.contra) / 1440;
      const vals: [number, ExcelJS.CellValue, string?][] = [
        [1, i + 1],
        [RES.nombre, { text: p.alias, hyperlink: `#'${HOJAS.asis}'!A${filaBloque(i)}` }],
        [RES.empresa, { formula: `IF(VLOOKUP(${S(RES.nombre)},Personas,${PC.empresa},FALSE)="","",VLOOKUP(${S(RES.nombre)},Personas,${PC.empresa},FALSE))`, result: p.empresa }],
        [RES.dias, { formula: `COUNT(${refFila('entrada', i)})`, result: fp.tot.dias }, '0'],
        [RES.faltas, { formula: `COUNTIF(${refFila('motivo', i)},"Falta")`, result: fp.tot.faltas }, '0'],
        [RES.trabajado, { formula: `SUM(${refFila('trabajado', i)})`, result: fp.tot.trabajado / 1440 }, '[h]:mm'],
        [RES.favor, { formula: `SUM(${fav})`, result: fp.tot.favor / 1440 }, '[h]:mm'],
        [RES.enFeriado, { formula: `SUMIF(${tipos},"DOM",${fav})+SUMIF(${tipos},"FER",${fav})`, result: fp.tot.enFeriado / 1440 }, '[h]:mm'],
        [RES.contra, { formula: `SUM(${refFila('contra', i)})`, result: fp.tot.contra / 1440 }, '[h]:mm'],
        [RES.antFavor, ant > 0 ? ant / 1440 : 0, '[h]:mm'],
        [RES.antContra, ant < 0 ? -ant / 1440 : 0, '[h]:mm'],
        [RES.castigo, 0, '[h]:mm'],
        [RES.pagadas, 0, '[h]:mm'],
        [RES.saldoFavor, { formula: `MAX(${neto},0)`, result: Math.max(saldo, 0) }, '[h]:mm'],
        [RES.saldoContra, { formula: `MAX(-(${neto}),0)`, result: Math.max(-saldo, 0) }, '[h]:mm'],
        [RES.pendientes, { formula: `COUNTIF(${refFila('estado', i)},"${MOTIVO_TXT_FALTA}")`, result: fp.tot.pendientes }, '0'],
        [RES.alertas, { formula: `COUNTIF(${refFila('estado', i)},"${TXT_REVISAR}")`, result: fp.dias.filter((d) => d.calc.alerta).length }, '0'],
      ];
      for (const [c, v, fmt] of vals) {
        const cell = ws.getCell(r, c);
        cell.value = v;
        if (fmt) cell.numFmt = fmt;
        cell.border = bordes;
        cell.alignment = { horizontal: c === RES.nombre ? 'left' : 'center' };
        const editable = c >= RES.antFavor && c <= RES.pagadas;
        if (editable) {
          cell.protection = { locked: false };
          cell.dataValidation = { type: 'decimal', operator: 'between', formulae: [0, 100], allowBlank: true, showErrorMessage: true, errorStyle: 'warning', errorTitle: '¿Horas correctas?', error: 'Escribe horas y minutos, por ejemplo 5:30 (o 0 si no hay). Para más de 24 horas escribe 30:15.' };
        } else {
          cell.fill = fill(c === RES.saldoFavor ? VERDE : c === RES.saldoContra ? ROJO_SUAVE : GRIS);
        }
      }
      ws.getCell(r, RES.nombre).font = { bold: true, color: { argb: 'FF1F4E79' }, underline: true };
      ws.getCell(r, RES.saldoFavor).font = { bold: true, color: { argb: 'FF375623' } };
      ws.getCell(r, RES.saldoContra).font = { bold: true, color: { argb: 'FF9C0006' } };
    });
    // Totales
    ws.getCell(rT, RES.nombre).value = 'TOTAL';
    ws.getCell(rT, RES.nombre).font = { bold: true };
    for (let c = RES.dias; c <= RES.alertas; c++) {
      const cell = ws.getCell(rT, c);
      const fmt = ws.getCell(r1, c).numFmt;
      cell.value = { formula: mes.filas.length ? `SUM(${col(c)}${r1}:${col(c)}${rT - 1})` : '0' };
      cell.numFmt = fmt;
      cell.font = { bold: true };
      cell.alignment = { horizontal: 'center' };
      cell.border = { top: { style: 'double' }, bottom: thin };
    }
    if (mes.filas.length) {
      ws.addConditionalFormatting({
        ref: `${col(RES.pendientes)}${r1}:${col(RES.alertas)}${rT}`,
        rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['0'], priority: 1, style: { fill: cfFill(AMARILLO), font: { bold: true } } }],
      });
      ws.autoFilter = { from: { row: r1 - 1, column: 1 }, to: { row: rT - 1, column: RES.alertas } };
    }
    ws.getCell(rT + 2, RES.nombre).value = 'Saldo final = saldo anterior + a favor − en contra − castigo − horas pagadas. "Le debemos" = horas a favor del trabajador; "Nos debe" = horas que la persona debe a la empresa.';
    ws.getCell(rT + 2, RES.nombre).font = { italic: true, size: 9, color: { argb: GRIS_OSC } };
    ws.views = [{ state: 'frozen', ySplit: r1 - 1, xSplit: 2, showGridLines: false }];
    ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true, formatCells: true, autoFilter: true, sort: true, objects: true });
  }

  // =====================================================================
  // SEMANAL
  // =====================================================================
  {
    const ws = semanal;
    ponerLogo(ws, 130);
    ws.getCell('C1').value = `HORAS POR SEMANA — ${titulo}`;
    ws.getCell('C1').font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell('C2').value = 'Semanas de lunes a sábado. Se calcula solo desde la hoja Asistencia.';
    ws.getCell('C2').font = { color: { argb: GRIS_OSC } };
    const semanas = [...new Set(mes.fechas.map((f) => f.semana))];
    const r1 = 7;
    ws.getColumn(1).width = 14;
    ws.getColumn(2).width = 9;
    const c0 = ws.getCell(r1 - 1, 1);
    c0.value = 'Nombre';
    estiloCabecera(c0);
    const c1 = ws.getCell(r1 - 1, 2);
    c1.value = 'Base semanal';
    estiloCabecera(c1);
    ws.getRow(r1 - 1).height = 30;
    const semRango = refRango(R_SEMANA);
    semanas.forEach((k, si) => {
      const dias = mes.fechas.filter((f) => f.semana === k).map((f) => f.fecha);
      const lunes = semanaLunes(dias[0]);
      const sab = semanaSabado(dias[0]);
      const parcial = lunes < mes.inicio || sab > mes.hasta;
      const a = 3 + si * 4;
      ws.mergeCells(r1 - 2, a, r1 - 2, a + 3);
      const h = ws.getCell(r1 - 2, a);
      h.value = `Semana ${k}: ${ddmm(dias[0])} – ${ddmm(dias[dias.length - 1])}${parcial ? ' (semana parcial)' : ''}`;
      h.alignment = { horizontal: 'center' };
      h.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      h.fill = fill(si % 2 ? 'FF2F5597' : 'FF5B9BD5');
      ['Trabajadas', 'A favor', 'En contra', 'Neto'].forEach((t, j) => {
        const c = ws.getCell(r1 - 1, a + j);
        c.value = t;
        estiloCabecera(c);
        ws.getColumn(a + j).width = 10;
      });
      mes.filas.forEach((fp, i) => {
        const r = r1 + i;
        const dd = fp.dias.filter((_, j) => mes.fechas[j].semana === k);
        const sum = (f: (x: (typeof dd)[number]) => number) => dd.reduce((acc, x) => acc + f(x), 0);
        const vals: [string, number][] = [
          [`SUMIF(${semRango},${k},${refFila('trabajado', i)})`, sum((x) => x.calc.trabajado)],
          [`SUMIF(${semRango},${k},${refFila('favor', i)})`, sum((x) => x.calc.favor)],
          [`SUMIF(${semRango},${k},${refFila('contra', i)})`, sum((x) => x.calc.contra)],
        ];
        vals.forEach(([f, res], j) => {
          const c = ws.getCell(r, a + j);
          c.value = { formula: f, result: res / 1440 };
          c.numFmt = '[h]:mm';
          c.border = bordes;
          c.alignment = { horizontal: 'center' };
        });
        const fa = `${col(a + 1)}${r}`, co = `${col(a + 2)}${r}`;
        const net = (sum((x) => x.calc.favor) - sum((x) => x.calc.contra));
        const cn = ws.getCell(r, a + 3);
        cn.value = { formula: `IF(${fa}-${co}<0,"-","+")&TEXT(ABS(${fa}-${co}),"[h]:mm")`, result: `${net < 0 ? '-' : '+'}${Math.floor(Math.abs(net) / 60)}:${String(Math.abs(net) % 60).padStart(2, '0')}` };
        cn.font = { bold: true };
        cn.border = bordes;
        cn.alignment = { horizontal: 'center' };
        cn.fill = fill(GRIS);
      });
    });
    mes.filas.forEach((fp, i) => {
      const r = r1 + i;
      const n = ws.getCell(r, 1);
      n.value = fp.persona.alias;
      n.font = { bold: true };
      n.border = bordes;
      const b = ws.getCell(r, 2);
      b.value = { formula: `VLOOKUP(A${r},Personas,${PC.horas},FALSE)`, result: fp.persona.horasSemanales };
      b.numFmt = '0" h"';
      b.border = bordes;
      b.alignment = { horizontal: 'center' };
    });
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: r1 - 1, showGridLines: false }];
    if (mes.filas.length) ws.autoFilter = { from: { row: r1 - 1, column: 1 }, to: { row: r1 - 1 + mes.filas.length, column: 2 } };
    await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true, formatCells: true, autoFilter: true, objects: true });
  }

  // =====================================================================
  // POR REVISAR
  // =====================================================================
  {
    const ws = revisar;
    ponerLogo(ws, 130);
    ws.getCell('C1').value = 'POR REVISAR';
    ws.getCell('C1').font = { bold: true, size: 16, color: { argb: ROJO } };
    ws.getCell('C2').value = 'Días que la página no pudo resolver sola. Haz clic en "Ir" para saltar a la celda, completa el dato en la hoja Asistencia y el estado cambia solo.';
    ws.getCell('C2').font = { color: { argb: GRIS_OSC } };
    const r1 = 6;
    const cab: [string, number][] = [['N°', 5], ['Persona', 13], ['Fecha', 11], ['Marcadas del reloj', 22], ['Qué pasó', 60], ['Sugerencia', 11], ['Ir a la celda', 9], ['Estado', 16], ['Visto (Sí)', 9]];
    cab.forEach(([t, w], i) => {
      const c = ws.getCell(r1 - 1, i + 1);
      c.value = t;
      estiloCabecera(c);
      ws.getColumn(i + 1).width = w;
    });
    let n = 0;
    mes.filas.forEach((fp, i) => {
      fp.dias.forEach((d, j) => {
        for (const caso of d.casos) {
          const r = r1 + n;
          n++;
          const X = cDia(j);
          const base = filaBloque(i);
          const destino = caso.faltaDato ? (d.entrada == null && caso.tipo !== 'sin_salida' ? FILA.entrada : FILA.salida) : FILA.reloj;
          const est = ref(X, base + FILA.estado);
          const visto = `I${r}`;
          const vals: ExcelJS.CellValue[] = [
            n,
            fp.persona.alias,
            excelSerial(d.fecha),
            d.reloj || '(sin marcadas)',
            caso.mensaje,
            caso.sugerencia ?? '',
            { text: 'Ir ›', hyperlink: `#'${HOJAS.asis}'!${L(X)}${base + destino}` },
            {
              formula: caso.faltaDato
                ? `IF(${est}="${MOTIVO_TXT_FALTA}","Falta completar",IF(${est}="${TXT_REVISAR}","Revisar horas","Listo"))`
                : `IF(${est}="${TXT_REVISAR}","Revisar horas",IF(${visto}="Sí","Listo","Mirar y marcar Visto"))`,
              result: caso.faltaDato ? 'Falta completar' : 'Mirar y marcar Visto',
            },
            null,
          ];
          vals.forEach((v, k) => {
            const c = ws.getCell(r, k + 1);
            c.value = v;
            c.border = bordes;
            c.alignment = { vertical: 'top', wrapText: k === 4 || k === 3 };
            if (k === 2) c.numFmt = '[$-es-PE]ddd dd/mm';
            if (k === 6) c.font = { color: { argb: 'FF0563C1' }, underline: true };
          });
          const v = ws.getCell(r, 9);
          v.protection = { locked: false };
          v.dataValidation = { type: 'list', formulae: ['"Sí,No"'], allowBlank: true, showErrorMessage: true, errorStyle: 'stop', error: 'Elige Sí o No.' };
          if (!caso.faltaDato) ws.getCell(r, 5).font = { color: { argb: 'FF7F6000' } };
        }
      });
    });
    const rT = r1 + n;
    const cont = ws.getCell('C3');
    cont.value = n
      ? { formula: `COUNTIF(H${r1}:H${rT - 1},"Falta completar")&" por completar · "&COUNTIF(H${r1}:H${rT - 1},"Revisar horas")&" por revisar · "&COUNTIF(H${r1}:H${rT - 1},"Mirar y marcar Visto")&" por mirar · "&COUNTIF(H${r1}:H${rT - 1},"Listo")&" listos"` }
      : '✔ No hay casos por revisar.';
    cont.font = { bold: true, size: 12, color: { argb: 'FF9C0006' } };
    if (n) {
      ws.addConditionalFormatting({
        ref: `H${r1}:H${rT - 1}`,
        rules: [
          { type: 'containsText', operator: 'containsText', text: 'Falta', priority: 1, style: { fill: cfFill(AMARILLO), font: { bold: true } } },
          { type: 'containsText', operator: 'containsText', text: 'Revisar', priority: 2, style: { fill: cfFill('FFFF7C80'), font: { bold: true } } },
          { type: 'containsText', operator: 'containsText', text: 'Mirar', priority: 3, style: { fill: cfFill(NARANJA) } },
          { type: 'containsText', operator: 'containsText', text: 'Listo', priority: 4, style: { fill: cfFill(VERDE), font: { color: { argb: 'FF375623' } } } },
        ],
      });
      ws.autoFilter = { from: { row: r1 - 1, column: 1 }, to: { row: rT - 1, column: 9 } };
    }
    ws.views = [{ state: 'frozen', ySplit: r1 - 1, showGridLines: false }];
    await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true, formatCells: true, autoFilter: true, sort: true, objects: true });
  }

  // =====================================================================
  // RELOJ ORIGINAL (copia fiel del reporte de la máquina)
  // =====================================================================
  {
    const ws = reloj;
    ws.getCell('A1').value = `Copia del reporte del reloj (${rep.inicio} ~ ${rep.fin}), tal como vino. Solo para consulta.`;
    ws.getCell('A1').font = { bold: true };
    const fechas = rep.fechas;
    ['ID', 'Nombre en el reloj'].forEach((t, i) => {
      const c = ws.getCell(3, i + 1);
      c.value = t;
      estiloCabecera(c);
    });
    ws.getColumn(1).width = 11;
    ws.getColumn(2).width = 20;
    fechas.forEach((f, j) => {
      const c = ws.getCell(3, 3 + j);
      c.value = `${DIAS_CORTO[weekday(f)]} ${ddmm(f)}`;
      estiloCabecera(c);
      ws.getColumn(3 + j).width = 12;
    });
    rep.personas.forEach((p, i) => {
      const r = 4 + i;
      ws.getCell(r, 1).value = p.id;
      ws.getCell(r, 2).value = p.nombre;
      fechas.forEach((f, j) => {
        const c = ws.getCell(r, 3 + j);
        c.value = (p.marcadas[f] ?? []).map(fmtHHMM).join(' ') || null;
        c.alignment = { wrapText: true, vertical: 'top' };
        c.font = { size: 9 };
        c.border = bordes;
      });
    });
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];
    await ws.protect('', { selectLockedCells: true, selectUnlockedCells: true, formatColumns: true, formatRows: true });
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
    const lineas: [string, ('t' | 'h' | 'p' | 'c')?, string?][] = [
      [`CONTROL DE ASISTENCIA — ${titulo}`, 't'],
      [`Período ${periodo}. Archivo generado el ${fechaHoy(op.generado ?? new Date())}.`],
      ['', 'p'],
      ['Cómo trabajar este archivo', 'h'],
      ['1. Abre la hoja "Por revisar": es la lista de días que hay que completar o mirar.'],
      ['2. Haz clic en "Ir ›" para saltar al día en la hoja "Asistencia". Ahí escribe la hora que falta (formato 24 h: 18:05) o elige un motivo en la lista.'],
      ['3. Cuando no quede nada en amarillo, revisa la hoja "Resumen". Todo se recalcula solo.'],
      ['4. Si quieres, ajusta en "Resumen" el saldo anterior, castigos u horas pagadas.'],
      ['', 'p'],
      ['Colores', 'h'],
      ['Falta un dato: escribe la hora o elige un motivo.', 'c', AMARILLO],
      ['Hora escrita a mano (no coincide con el reloj). Sirve para saber qué se corrigió.', 'c', AZUL],
      ['Revisar: la salida es antes que la entrada. Seguramente se escribió en formato de 12 h (6:14 en vez de 18:14).', 'c', 'FFFF7C80'],
      ['Día con motivo u observación.', 'c', VERDE],
      ['Calculado con fórmulas. No se escribe (la hoja avisará si lo intentas).', 'c', GRIS],
      ['Marcadas del reloj tal como vinieron. Pasa el mouse sobre las celdas con triángulo rojo para ver la explicación.', 'c', GRIS_RELOJ],
      ['', 'p'],
      ['Reglas que usa el archivo (se cambian en la hoja "Configuracion")', 'h'],
      [`• Extras: solo si sale ${cfg.reglas.margenExtras} min o más después de su hora (18:${String(cfg.reglas.margenExtras).padStart(2, '0')}); se cuentan desde su hora de salida. Sábado igual, desde las 12:00.`],
      [`• Tardanza: se resta de la bolsa solo si es de ${cfg.reglas.tardanzaDesde} min o más.`],
      ['• Salida antes de hora: se resta de la bolsa.'],
      ['• Llegar temprano: solo suma para quien lo tiene configurado (columna "Llegar temprano cuenta si llega antes de").'],
      ['• Motivos "Permiso" y "Estudia" sin marcadas: restan la jornada del día. "Falta": no toca la bolsa, se cuenta aparte.'],
      [`• Domingo o feriado trabajado: todo el tiempo suma a favor (x${cfg.reglas.factorFeriado}).`],
      ['• Saldo final = saldo anterior + a favor − en contra − castigo − horas pagadas.'],
      ['', 'p'],
      ['Seguridad', 'h'],
      ['Las hojas están protegidas SIN contraseña solo para evitar borrar fórmulas por accidente. Si de verdad necesitas cambiar algo: Revisar › Desproteger hoja.'],
      ['Si escribes una hora rara, Excel te preguntará "¿Quieres dejarlo igual?": revisa antes de aceptar.'],
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

function semanaLunes(f: ISODate): ISODate {
  const w = weekday(f);
  const { y, mo, d } = parseISO(f);
  const t = new Date(Date.UTC(y, mo - 1, d - (w - 1)));
  return t.toISOString().slice(0, 10);
}
function semanaSabado(f: ISODate): ISODate {
  const w = weekday(f);
  const { y, mo, d } = parseISO(f);
  const t = new Date(Date.UTC(y, mo - 1, d + (6 - w)));
  return t.toISOString().slice(0, 10);
}
function fechaHoy(d: Date): string {
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

export async function excelABuffer(wb: ExcelJS.Workbook): Promise<ArrayBuffer> {
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}
