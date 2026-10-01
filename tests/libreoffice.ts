/** Recalcula un .xlsx con LibreOffice en modo headless (forzando el recálculo de todas las fórmulas). */
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';

export const hayLibreOffice = (() => {
  try {
    execFileSync('soffice', ['--version'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
})();

let perfil: string | null = null;
function perfilLO(): string {
  if (perfil) return perfil;
  perfil = mkdtempSync(join(tmpdir(), 'lo-perfil-'));
  const user = join(perfil, 'user');
  mkdirSync(user, { recursive: true });
  // OOXMLRecalcMode 0 = recalcular siempre al abrir
  writeFileSync(join(user, 'registrymodifications.xcu'), `<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>`);
  return perfil;
}

export function recalcular(archivo: string): string {
  const out = mkdtempSync(join(tmpdir(), 'lo-out-'));
  execFileSync('soffice', [`-env:UserInstallation=file://${perfilLO()}`, '--headless', '--calc', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', out, archivo], { stdio: 'pipe', timeout: 120000 });
  const r = join(out, basename(archivo));
  if (!existsSync(r)) throw new Error('LibreOffice no generó el archivo recalculado');
  return r;
}
