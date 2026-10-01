import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { configInicial, normalizarConfig, type Config } from '../src/lib/config';
import { leerReporte, type Reporte } from '../src/lib/parser';
import { sincronizarPersonas } from '../src/lib/model';

export const SAMPLES = resolve(__dirname, '../samples');
export const hayMuestras = existsSync(resolve(SAMPLES, '1_report.xls'));

export function reporteSetiembre(): Reporte {
  return leerReporte(new Uint8Array(readFileSync(resolve(SAMPLES, '1_report.xls'))));
}

/** Config privada (no versionada) con los nombres cortos y criterios de Yasmín. */
export function configPrivada(rep: Reporte): Config {
  const f = resolve(SAMPLES, 'qalen-config.json');
  const base = existsSync(f) ? normalizarConfig(JSON.parse(readFileSync(f, 'utf8'))) : configInicial(2026);
  base.feriados = configInicial(2026).feriados;
  return sincronizarPersonas(rep, base).config;
}
