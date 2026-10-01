import { useEffect, useMemo, useState } from 'react';
import logoUrl from './assets/logo-qalen.png';
import { cargarConfig, configInicial, guardarConfig, validarConfig, type Config } from './lib/config';
import { construirMes, corteMes, listarCasos, sincronizarPersonas } from './lib/model';
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

  useEffect(() => guardarConfig(config), [config]);

  const erroresConfig = useMemo(() => validarConfig(config), [config]);
  const mes = useMemo(() => (reporte && hasta ? construirMes(reporte, config, hasta) : null), [reporte, config, hasta]);
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
      setHasta(corteMes(rep, nueva));
      setAvisosCarga([...rep.avisos, ...avisos]);
    } catch (e) {
      setReporte(null);
      setError(e instanceof Error ? e.message : 'No se pudo leer el archivo.');
    }
  }

  return (
    <div className="min-h-screen">
      <header className="border-b-4 border-qalen bg-white shadow-sm">
        <div className="mx-auto flex max-w-6xl flex-wrap items-center gap-4 px-4 py-3">
          <img src={logoUrl} alt="QALEN · Química Allende S.A.C." className="h-12 w-auto" />
          <div className="mr-auto">
            <h1 className="text-xl font-bold text-qalen">Control de asistencia</h1>
            <p className="text-xs text-stone-500">Bolsa de horas desde el reporte del reloj</p>
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
            hasta={hasta}
            setHasta={setHasta}
            avisos={[...avisosCarga, ...(mes?.avisos ?? [])]}
            error={error}
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
