import { useRef, useState } from 'react';
import { feriadosPeru } from '../lib/calendar';
import { normalizarConfig, type Config, type Persona } from '../lib/config';
import { baseDiaria, valorHora } from '../lib/model';
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
                  {['Activa', 'Nombre corto', 'Nombre en el reloj', 'Empresa', 'Horas / semana', 'Entrada L-V', 'Salida L-V', 'Entrada sáb.', 'Salida sáb.', 'Extras desde (L-V)', 'Sueldo mensual (S/)', 'Asig. familiar', 'Valor hora según', 'Valor hora', 'Descuenta tardanzas y salidas', 'Reparto', 'Fecha de ingreso'].map((t) => (
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
                    <Hora
                      v={p.salidaLV}
                      on={(v) => v != null && setP(p.id, { salidaLV: v, ...(p.extrasDesdeLV === p.salidaLV || p.extrasDesdeLV < v ? { extrasDesdeLV: v } : {}) })}
                    />
                    <Hora v={p.entradaSab} on={(v) => v != null && setP(p.id, { entradaSab: v })} />
                    <Hora v={p.salidaSab} on={(v) => v != null && setP(p.id, { salidaSab: v })} />
                    <Hora v={p.extrasDesdeLV} on={(v) => v != null && setP(p.id, { extrasDesdeLV: v })} />
                    <td className="p-1">
                      <input type="number" className={`w-24 ${p.activo && !(p.sueldo > 0) ? 'bg-yellow-200' : ''}`} min={0} step={0.01} value={p.sueldo || ''} placeholder="0.00"
                        onChange={(e) => setP(p.id, { sueldo: Number(e.target.value) || 0 })} />
                    </td>
                    <td className="p-1 text-center"><input type="checkbox" checked={p.asignacionFamiliar} onChange={(e) => setP(p.id, { asignacionFamiliar: e.target.checked })} /></td>
                    <td className="p-1">
                      <select value={p.valorHoraSegun} onChange={(e) => setP(p.id, { valorHoraSegun: e.target.value as Persona['valorHoraSegun'] })}>
                        <option value="ocho">30 días × 8 h</option>
                        <option value="reales">Horas reales ({fmtDur(baseDiaria(p))} al día)</option>
                      </select>
                    </td>
                    <td className="whitespace-nowrap p-1 text-right font-mono text-xs">
                      {p.sueldo > 0 ? `S/ ${valorHora(p, config.reglas, p.valorHoraSegun).toFixed(2)}` : '—'}
                    </td>
                    <td className="p-1 text-center"><input type="checkbox" checked={p.descuenta} onChange={(e) => setP(p.id, { descuenta: e.target.checked })} /></td>
                    <td className="p-1 text-center"><input type="checkbox" checked={p.reparto} onChange={(e) => setP(p.id, { reparto: e.target.checked })} /></td>
                    <td className="p-1"><input type="date" value={p.fechaIngreso ?? ''} onChange={(e) => setP(p.id, { fechaIngreso: e.target.value || null })} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <p className="mt-2 text-xs text-stone-500">
              Valor hora = (sueldo + asignación familiar) ÷ 30 días ÷ 8 h, redondeado a 2 decimales. Con “Horas reales” se divide entre sus horas por día
              (horas semanales ÷ 6), para quien no trabaja 8 h diarias. El sueldo también se puede escribir en el Resumen del Excel. “Extras desde” sirve para
              casos especiales (p. ej. medio tiempo que sale 12:00 pero sus extras se cuentan desde las 18:00); se ajusta solo al cambiar la salida L-V.
            </p>
          </div>
        )}
      </section>

      <section className="grid gap-6 md:grid-cols-2">
        <div className="rounded-xl bg-white p-5 shadow-sm">
          <h2 className="text-lg font-bold">Reglas</h2>
          <div className="mt-3 space-y-3 text-sm">
            <Regla t="Margen para extras (min)" ayuda="Las extras cuentan si sale este tiempo o más después de su hora de salida. 30 = desde las 18:30.">
              <input type="number" className="w-20" min={0} max={120} value={r.margenExtras} onChange={(e) => setR({ margenExtras: Number(e.target.value) })} />
            </Regla>
            <Regla t="Tardanza mínima (min)" ayuda="La tardanza cuenta como hora en contra si es de este tiempo o más. 0 = toda tardanza cuenta.">
              <input type="number" className="w-20" min={0} max={600} value={r.tardanzaDesde} onChange={(e) => setR({ tardanzaDesde: Number(e.target.value) })} />
            </Regla>
            <Regla t="Almuerzo L-V (min)" ayuda="Se descuenta de las horas trabajadas.">
              <input type="number" className="w-20" min={0} max={180} value={r.almuerzoMin} onChange={(e) => setR({ almuerzoMin: Number(e.target.value) })} />
            </Regla>
            <Regla t="Horas del primer recargo" ayuda="Cada día, estas primeras horas extras van al primer recargo; las demás, al segundo. Ley: 2 horas.">
              <input type="number" className="w-20" min={0} max={10} step={0.5} value={r.tramo1Min / 60} onChange={(e) => setR({ tramo1Min: Math.round(Number(e.target.value) * 60) })} />
            </Regla>
            <Regla t="Primer recargo" ayuda="Multiplicador del valor hora. Ley: mínimo 1.25 (25%).">
              <input type="number" className="w-20" min={1} max={4} step={0.05} value={r.recargo1} onChange={(e) => setR({ recargo1: Number(e.target.value) })} />
            </Regla>
            <Regla t="Segundo recargo" ayuda="Desde la 3.ª hora extra del día. Ley: mínimo 1.35 (35%).">
              <input type="number" className="w-20" min={1} max={4} step={0.05} value={r.recargo2} onChange={(e) => setR({ recargo2: Number(e.target.value) })} />
            </Regla>
            <Regla t="Recargo domingo / feriado" ayuda="Horas trabajadas en domingo o feriado sin descanso sustituto. Ley: 2.00 (100%).">
              <input type="number" className="w-20" min={1} max={4} step={0.05} value={r.recargoFeriado} onChange={(e) => setR({ recargoFeriado: Number(e.target.value) })} />
            </Regla>
            <Regla t="Asignación familiar (S/)" ayuda="10% de la remuneración mínima vital. Se suma al sueldo de quien la recibe para calcular el valor hora.">
              <input type="number" className="w-20" min={0} step={0.01} value={r.montoAsignacion} onChange={(e) => setR({ montoAsignacion: Number(e.target.value) })} />
            </Regla>
            <p className="text-xs text-amber-700">Los recargos son los mínimos de ley (D.S. 007-2002-TR). Confírmalos con el contador, igual que el monto de la asignación familiar.</p>
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
        <h2 className="text-lg font-bold">Respaldo</h2>
        <div className="mt-3 grid gap-4 md:grid-cols-2">
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
