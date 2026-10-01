import { useRef, useState } from 'react';
import { feriadosPeru } from '../lib/calendar';
import { normalizarConfig, type Config, type Persona } from '../lib/config';
import { leerSaldosDeExcel } from '../lib/saldos';
import { fmtDur, fmtHHMM, parseHHMM } from '../lib/time';
import { Caja } from './Procesar';

interface Props {
  config: Config;
  setConfig: (c: Config) => void;
  errores: string[];
}

export function Personas({ config, setConfig, errores }: Props) {
  const [msg, setMsg] = useState<{ tipo: 'ok' | 'error' | 'aviso'; txt: string } | null>(null);
  const [nuevoFer, setNuevoFer] = useState({ fecha: '', nombre: '' });
  const [verInactivas, setVerInactivas] = useState(false);
  const inJson = useRef<HTMLInputElement>(null);
  const inXlsx = useRef<HTMLInputElement>(null);

  const setP = (id: string, cambio: Partial<Persona>) => setConfig({ ...config, personas: config.personas.map((p) => (p.id === id ? { ...p, ...cambio } : p)) });
  const r = config.reglas;
  const setR = (cambio: Partial<Config['reglas']>) => setConfig({ ...config, reglas: { ...r, ...cambio } });
  const personas = config.personas.filter((p) => verInactivas || p.activo);

  function exportar() {
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `respaldo-asistencia-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
  }

  async function importarJson(f: File) {
    try {
      const c = normalizarConfig(JSON.parse(await f.text()));
      if (!window.confirm(`Esto reemplaza la configuración actual por la del respaldo (${c.personas.length} personas). ¿Continuar?`)) return;
      setConfig(c);
      setMsg({ tipo: 'ok', txt: 'Respaldo cargado.' });
    } catch (e) {
      setMsg({ tipo: 'error', txt: `No se pudo leer el respaldo: ${e instanceof Error ? e.message : e}` });
    }
  }

  async function importarSaldos(f: File) {
    try {
      const saldos = leerSaldosDeExcel(new Uint8Array(await f.arrayBuffer()));
      const porAlias = new Map(config.personas.map((p) => [p.alias.toUpperCase(), p]));
      const cambios = saldos.filter((s) => porAlias.has(s.alias.toUpperCase()));
      const sinPersona = saldos.filter((s) => !porAlias.has(s.alias.toUpperCase())).map((s) => s.alias);
      const lista = cambios.map((s) => `${s.alias}: ${s.minutos >= 0 ? `le debemos ${fmtDur(s.minutos)}` : `nos debe ${fmtDur(-s.minutos)}`}`).join('\n');
      if (!window.confirm(`Se usarán estos saldos finales como "saldo anterior" del próximo mes:\n\n${lista}${sinPersona.length ? `\n\nNo encontré a: ${sinPersona.join(', ')}` : ''}\n\n¿Continuar?`)) return;
      setConfig({
        ...config,
        personas: config.personas.map((p) => {
          const s = cambios.find((x) => x.alias.toUpperCase() === p.alias.toUpperCase());
          return s ? { ...p, saldoAnterior: s.minutos } : p;
        }),
      });
      setMsg({ tipo: 'ok', txt: `Saldos actualizados para ${cambios.length} personas.` });
    } catch (e) {
      setMsg({ tipo: 'error', txt: e instanceof Error ? e.message : String(e) });
    }
  }

  function agregarFeriado() {
    if (!nuevoFer.fecha) return;
    if (config.feriados.some((f) => f.fecha === nuevoFer.fecha)) {
      setMsg({ tipo: 'aviso', txt: 'Esa fecha ya está en la lista.' });
      return;
    }
    setConfig({ ...config, feriados: [...config.feriados, { fecha: nuevoFer.fecha, nombre: nuevoFer.nombre || 'Feriado' }].sort((a, b) => a.fecha.localeCompare(b.fecha)) });
    setNuevoFer({ fecha: '', nombre: '' });
  }

  const anios = [...new Set(config.feriados.map((f) => f.fecha.slice(0, 4)))].sort();

  return (
    <div className="space-y-6">
      {errores.length > 0 && (
        <Caja tipo="error">
          <p className="font-semibold">Hay que corregir:</p>
          <ul className="ml-5 list-disc">{errores.map((e, i) => <li key={i}>{e}</li>)}</ul>
        </Caja>
      )}
      {msg && <Caja tipo={msg.tipo}>{msg.txt}</Caja>}

      <section className="rounded-xl bg-white p-5 shadow-sm">
        <div className="flex flex-wrap items-center gap-3">
          <h2 className="mr-auto text-lg font-bold">Personas</h2>
          <label className="flex items-center gap-2 text-sm">
            <input type="checkbox" checked={verInactivas} onChange={(e) => setVerInactivas(e.target.checked)} /> Ver también las inactivas
          </label>
        </div>
        {config.personas.length === 0 ? (
          <p className="mt-3 text-sm text-stone-500">Todavía no hay personas. Se agregan solas al subir el primer reporte del reloj.</p>
        ) : (
          <div className="mt-3 overflow-x-auto">
            <table className="text-sm">
              <thead>
                <tr className="bg-qalen text-left text-xs text-white">
                  {['Activa', 'Nombre corto', 'Nombre en el reloj', 'Empresa', 'Horas / semana', 'Entrada', 'Salida L-V', 'Salida sáb.', 'Extras desde (L-V)', 'Llegar temprano cuenta si llega antes de', 'Descuenta tardanzas y salidas', 'Reparto', 'Fecha de ingreso', 'Saldo anterior'].map((t) => (
                    <th key={t} className="p-2 align-bottom">{t}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {personas.map((p) => (
                  <tr key={p.id} className={`border-b border-stone-200 ${p.activo ? '' : 'text-stone-400'}`}>
                    <td className="p-1 text-center"><input type="checkbox" checked={p.activo} onChange={(e) => setP(p.id, { activo: e.target.checked })} /></td>
                    <td className="p-1"><input type="text" className="w-28 font-semibold uppercase" value={p.alias} onChange={(e) => setP(p.id, { alias: e.target.value.toUpperCase() })} /></td>
                    <td className="p-1 text-xs">{p.nombreReloj}<span className="block text-stone-400">ID {p.id}</span></td>
                    <td className="p-1"><input type="text" className="w-24" value={p.empresa} onChange={(e) => setP(p.id, { empresa: e.target.value })} /></td>
                    <td className="p-1"><input type="number" className="w-16" min={1} max={72} value={p.horasSemanales} onChange={(e) => setP(p.id, { horasSemanales: Number(e.target.value) })} /></td>
                    <Hora v={p.entrada} on={(v) => v != null && setP(p.id, { entrada: v })} />
                    <Hora v={p.salidaLV} on={(v) => v != null && setP(p.id, { salidaLV: v })} />
                    <Hora v={p.salidaSab} on={(v) => v != null && setP(p.id, { salidaSab: v })} />
                    <Hora v={p.extrasDesdeLV} on={(v) => v != null && setP(p.id, { extrasDesdeLV: v })} />
                    <Hora v={p.tempranoAntesDe} vacio on={(v) => setP(p.id, { tempranoAntesDe: v })} />
                    <td className="p-1 text-center"><input type="checkbox" checked={p.descuenta} onChange={(e) => setP(p.id, { descuenta: e.target.checked })} /></td>
                    <td className="p-1 text-center"><input type="checkbox" checked={p.reparto} onChange={(e) => setP(p.id, { reparto: e.target.checked })} /></td>
                    <td className="p-1"><input type="date" value={p.fechaIngreso ?? ''} onChange={(e) => setP(p.id, { fechaIngreso: e.target.value || null })} /></td>
                    <td className="p-1"><Saldo key={p.saldoAnterior} v={p.saldoAnterior} on={(v) => setP(p.id, { saldoAnterior: v })} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <p className="mt-2 text-xs text-stone-500">
              “Llegar temprano…” vacío = llegar antes no suma. “Extras desde” sirve para casos especiales (p. ej. medio tiempo que sale 12:00 pero sus extras se cuentan desde las 18:00).
            </p>
          </div>
        )}
      </section>

      <section className="grid gap-6 md:grid-cols-2">
        <div className="rounded-xl bg-white p-5 shadow-sm">
          <h2 className="text-lg font-bold">Reglas</h2>
          <div className="mt-3 space-y-3 text-sm">
            <Regla t="Margen para extras (min)" ayuda="Las extras cuentan si sale este tiempo o más después de su hora. 20 = desde las 18:20.">
              <input type="number" className="w-20" min={0} max={120} value={r.margenExtras} onChange={(e) => setR({ margenExtras: Number(e.target.value) })} />
            </Regla>
            <Regla t="Tardanza que resta (min)" ayuda="Solo se resta de la bolsa si llega tarde por este tiempo o más. 0 = toda tardanza resta.">
              <input type="number" className="w-20" min={0} max={600} value={r.tardanzaDesde} onChange={(e) => setR({ tardanzaDesde: Number(e.target.value) })} />
            </Regla>
            <Regla t="Almuerzo L-V (min)" ayuda="Se descuenta de las horas trabajadas (no de la bolsa).">
              <input type="number" className="w-20" min={0} max={180} value={r.almuerzoMin} onChange={(e) => setR({ almuerzoMin: Number(e.target.value) })} />
            </Regla>
            <Regla t="Factor domingo / feriado" ayuda="1 = cada hora suma 1 hora; 2 = suma doble.">
              <input type="number" className="w-20" min={0} max={3} step={0.25} value={r.factorFeriado} onChange={(e) => setR({ factorFeriado: Number(e.target.value) })} />
            </Regla>
          </div>
        </div>

        <div className="rounded-xl bg-white p-5 shadow-sm">
          <h2 className="text-lg font-bold">Feriados</h2>
          <p className="text-xs text-amber-700">Verifica cada año el calendario oficial: el gobierno puede agregar feriados.</p>
          <div className="mt-2 flex flex-wrap gap-2">
            <input type="date" value={nuevoFer.fecha} onChange={(e) => setNuevoFer({ ...nuevoFer, fecha: e.target.value })} />
            <input type="text" placeholder="Nombre" value={nuevoFer.nombre} onChange={(e) => setNuevoFer({ ...nuevoFer, nombre: e.target.value })} />
            <button className="rounded bg-qalen px-3 py-1 text-sm font-semibold text-white" onClick={agregarFeriado}>Agregar</button>
            <button
              className="rounded border border-stone-300 px-3 py-1 text-sm"
              onClick={() => {
                const y = Number(window.prompt('¿Qué año quieres precargar?', String(new Date().getFullYear() + 1)));
                if (!y) return;
                const nuevos = feriadosPeru(y).filter((f) => !config.feriados.some((x) => x.fecha === f.fecha));
                setConfig({ ...config, feriados: [...config.feriados, ...nuevos].sort((a, b) => a.fecha.localeCompare(b.fecha)) });
              }}
            >
              Precargar un año
            </button>
          </div>
          <div className="mt-3 max-h-64 overflow-y-auto text-sm">
            {anios.map((a) => (
              <div key={a}>
                <p className="mt-2 font-semibold">{a}</p>
                {config.feriados.filter((f) => f.fecha.startsWith(a)).map((f) => (
                  <div key={f.fecha} className="flex items-center gap-2 border-b border-stone-100 py-0.5">
                    <span className="w-24 font-mono">{f.fecha.split('-').reverse().join('/')}</span>
                    <span className="mr-auto">{f.nombre}</span>
                    <button className="text-stone-400 hover:text-red-600" title="Quitar" onClick={() => setConfig({ ...config, feriados: config.feriados.filter((x) => x.fecha !== f.fecha) })}>✕</button>
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>
      </section>

      <section className="rounded-xl bg-white p-5 shadow-sm">
        <h2 className="text-lg font-bold">Saldos y respaldo</h2>
        <div className="mt-3 grid gap-4 md:grid-cols-2">
          <div>
            <p className="text-sm font-semibold">Pasar los saldos al mes siguiente</p>
            <p className="text-xs text-stone-500">Elige el Excel del mes que ya terminaste (guardado en Excel). Sus saldos finales quedan como “saldo anterior”.</p>
            <button className="mt-2 rounded bg-qalen px-3 py-2 text-sm font-semibold text-white" onClick={() => inXlsx.current?.click()}>Leer saldos de un Excel terminado</button>
            <input ref={inXlsx} type="file" accept=".xlsx" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) importarSaldos(f); e.target.value = ''; }} />
          </div>
          <div>
            <p className="text-sm font-semibold">Respaldo de la configuración</p>
            <p className="text-xs text-stone-500">La configuración se guarda en este navegador. Descarga un respaldo de vez en cuando, o para pasarla a otra computadora.</p>
            <div className="mt-2 flex gap-2">
              <button className="rounded border border-stone-300 px-3 py-2 text-sm font-semibold" onClick={exportar}>Descargar respaldo</button>
              <button className="rounded border border-stone-300 px-3 py-2 text-sm font-semibold" onClick={() => inJson.current?.click()}>Cargar respaldo</button>
              <input ref={inJson} type="file" accept=".json" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) importarJson(f); e.target.value = ''; }} />
            </div>
          </div>
        </div>
      </section>
    </div>
  );
}

function Hora({ v, on, vacio }: { v: number | null; on: (v: number | null) => void; vacio?: boolean }) {
  return (
    <td className="p-1">
      <input type="time" value={fmtHHMM(v)} required={!vacio} onChange={(e) => on(e.target.value ? parseHHMM(e.target.value) : vacio ? null : v)} />
    </td>
  );
}

function Saldo({ v, on }: { v: number; on: (v: number) => void }) {
  const abs = Math.abs(v);
  const signo = v < 0 ? -1 : 1;
  const [txt, setTxt] = useState(fmtDur(abs));
  return (
    <div className="flex items-center gap-1">
      <input
        type="text"
        className="w-16 text-right font-mono"
        value={txt}
        onChange={(e) => setTxt(e.target.value)}
        onBlur={() => {
          const m = /^(\d{1,4}):([0-5]\d)$/.exec(txt.trim());
          if (!m) {
            setTxt(fmtDur(abs));
            return;
          }
          on(signo * (Number(m[1]) * 60 + Number(m[2])));
        }}
        title="Horas:minutos, por ejemplo 15:07"
      />
      <select value={signo} onChange={(e) => on(Number(e.target.value) * abs)}>
        <option value={1}>le debemos</option>
        <option value={-1}>nos debe</option>
      </select>
    </div>
  );
}

function Regla({ t, ayuda, children }: { t: string; ayuda: string; children: React.ReactNode }) {
  return (
    <label className="flex items-start justify-between gap-3">
      <span>
        <span className="font-semibold">{t}</span>
        <span className="block text-xs text-stone-500">{ayuda}</span>
      </span>
      {children}
    </label>
  );
}
