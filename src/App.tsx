import { useEffect, useMemo, useState } from 'react';
import logoUrl from './assets/logo-qalen.png';
import { cargarConfig, configInicial, guardarConfig, validarConfig, type Config } from './lib/config';
import { aplicarSueldos, combinarConAnterior, leerExcelAnterior, type ExcelAnterior } from './lib/anterior';
import { construirMes, listarCasos, sincronizarPersonas } from './lib/model';
import { leerReporte, type Reporte } from './lib/parser';
import { Procesar } from './ui/Procesar';
import { Personas } from './ui/Personas';

type Pestana = 'procesar' | 'personas';

export default function App() {
  const [config, setConfig] = useState<Config>(() => cargarConfig() ?? configInicial());
  const [pestana, setPestana] = useState<Pestana>('procesar');
  const [reporte, setReporte] = useState<Reporte | null>(null);
  const [nombreArchivo, setNombreArchivo] = useState('');
  const [hasta, setHasta] = useState('');
  const [avisosCarga, setAvisosCarga] = useState<string[]>([]);
  const [error, setError] = useState('');
  const [anterior, setAnterior] = useState<{ datos: ExcelAnterior; nombre: string; avisos: string[] } | null>(null);
  const [errorAnterior, setErrorAnterior] = useState('');

  useEffect(() => guardarConfig(config), [config]);

  const erroresConfig = useMemo(() => validarConfig(config), [config]);
  // Si hay un Excel anterior ya trabajado, se junta con el reporte nuevo (rango completo + lo que ella completó)
  const combinado = useMemo(() => (reporte && anterior ? combinarConAnterior(reporte, anterior.datos, config) : null), [reporte, anterior, config]);
  const repEfectivo = combinado?.rep ?? reporte;
  // "hasta" vacío = todo el rango
  const hastaEfectivo = repEfectivo ? (hasta && hasta >= repEfectivo.inicio && hasta <= repEfectivo.fin ? hasta : repEfectivo.fin) : '';
  const mes = useMemo(
    () => (repEfectivo && hastaEfectivo ? construirMes(repEfectivo, config, hastaEfectivo, combinado?.ediciones) : null),
    [repEfectivo, config, hastaEfectivo, combinado],
  );
  const casos = useMemo(() => (mes ? listarCasos(mes) : []), [mes]);

  async function cargarArchivo(f: File) {
    setError('');
    try {
      if (f.size > 20 * 1024 * 1024) throw new Error('El archivo es demasiado grande para ser el reporte del reloj.');
      const rep = leerReporte(new Uint8Array(await f.arrayBuffer()));
      const { config: nueva, avisos } = sincronizarPersonas(rep, config);
      setConfig(nueva);
      setReporte(rep);
      setNombreArchivo(f.name);
      setHasta('');
      setAvisosCarga([...rep.avisos, ...avisos]);
    } catch (e) {
      setReporte(null);
      setError(e instanceof Error ? e.message : 'No se pudo leer el archivo.');
    }
  }

  async function cargarAnterior(f: File | null) {
    setErrorAnterior('');
    if (!f) {
      setAnterior(null);
      setHasta('');
      return;
    }
    try {
      if (f.size > 30 * 1024 * 1024) throw new Error('El archivo es demasiado grande.');
      const datos = leerExcelAnterior(new Uint8Array(await f.arrayBuffer()));
      const { config: nueva, avisos } = aplicarSueldos(datos, config);
      setConfig(nueva);
      setAnterior({ datos, nombre: f.name, avisos });
      setHasta('');
    } catch (e) {
      setAnterior(null);
      setErrorAnterior(e instanceof Error ? e.message : 'No se pudo leer el Excel anterior.');
    }
  }

  return (
    <div className="min-h-screen">
      <header className="border-b-4 border-qalen bg-white shadow-sm">
        <div className="mx-auto flex max-w-6xl flex-wrap items-center gap-4 px-4 py-3">
          <img src={logoUrl} alt="QALEN · Química Allende S.A.C." className="h-12 w-auto" />
          <div className="mr-auto">
            <h1 className="text-xl font-bold text-qalen">Control de asistencia</h1>
            <p className="text-xs text-stone-500">Horas extras y asistencia desde el reporte del reloj</p>
          </div>
          <nav className="flex gap-1 rounded-lg bg-stone-100 p-1">
            {([['procesar', 'Procesar mes'], ['personas', 'Personas y reglas']] as const).map(([k, t]) => (
              <button
                key={k}
                onClick={() => setPestana(k)}
                className={`rounded-md px-4 py-2 text-sm font-semibold transition ${pestana === k ? 'bg-qalen text-white shadow' : 'text-stone-600 hover:bg-white'}`}
              >
                {t}
                {k === 'personas' && erroresConfig.length > 0 && <span className="ml-2 rounded-full bg-amber-400 px-1.5 text-xs text-stone-900">!</span>}
              </button>
            ))}
          </nav>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-4 py-6">
        {pestana === 'procesar' ? (
          <Procesar
            reporte={reporte}
            nombreArchivo={nombreArchivo}
            config={config}
            mes={mes}
            casos={casos}
            hasta={hastaEfectivo}
            setHasta={setHasta}
            rango={repEfectivo ? { inicio: repEfectivo.inicio, fin: repEfectivo.fin } : null}
            avisos={[...avisosCarga, ...(anterior?.avisos ?? []), ...(combinado?.avisos ?? []), ...(mes?.avisos ?? [])]}
            error={error}
            anterior={anterior ? { nombre: anterior.nombre, inicio: anterior.datos.inicio, fin: anterior.datos.fin, copiados: combinado?.copiados ?? null } : null}
            errorAnterior={errorAnterior}
            onAnterior={cargarAnterior}
            erroresConfig={erroresConfig}
            onArchivo={cargarArchivo}
            irAPersonas={() => setPestana('personas')}
            logoUrl={logoUrl}
          />
        ) : (
          <Personas config={config} setConfig={setConfig} errores={erroresConfig} />
        )}
      </main>
      <footer className="mx-auto max-w-6xl px-4 pb-8 text-center text-xs text-stone-400">
        Todo se procesa en esta computadora: ningún dato sale a internet. La configuración se guarda en este navegador.
      </footer>
    </div>
  );
}
