import { Fragment, useRef, useState } from 'react';
import type { Config } from '../lib/config';
import { excelABuffer, generarExcel } from '../lib/excel';
import type { CasoListado, Mes } from '../lib/model';
import type { Reporte } from '../lib/parser';
import { DIAS_CORTO, ddmm, fmtDur, fmtHHMM, nombrePeriodo, weekday } from '../lib/time';

interface Props {
  reporte: Reporte | null;
  nombreArchivo: string;
  config: Config;
  mes: Mes | null;
  casos: CasoListado[];
  hasta: string;
  setHasta: (s: string) => void;
  avisos: string[];
  error: string;
  erroresConfig: string[];
  onArchivo: (f: File) => void;
  irAPersonas: () => void;
  logoUrl: string;
}

export function Procesar(p: Props) {
  const [arrastrando, setArrastrando] = useState(false);
  const [abierto, setAbierto] = useState<string | null>(null);
  const [confirmar, setConfirmar] = useState(false);
  const [generando, setGenerando] = useState(false);
  const [errorDescarga, setErrorDescarga] = useState('');
  const input = useRef<HTMLInputElement>(null);

  const pendientes = p.casos.filter((c) => c.caso.faltaDato).length;
  const pct = (f: number) => `${Math.round((f - 1) * 100)}%`;
  const pct1 = pct(p.config.reglas.recargo1), pct2 = pct(p.config.reglas.recargo2), pctF = pct(p.config.reglas.recargoFeriado);
  const sinSueldo = p.mes?.filas.filter((f) => !(f.persona.sueldo > 0)).map((f) => f.persona.alias) ?? [];
  const paraMirar = p.casos.length - pendientes;

  function soltar(e: React.DragEvent) {
    e.preventDefault();
    setArrastrando(false);
    const f = e.dataTransfer.files?.[0];
    if (f) p.onArchivo(f);
  }

  async function descargar() {
    if (!p.mes || !p.reporte) return;
    setConfirmar(false);
    setGenerando(true);
    setErrorDescarga('');
    try {
      const logo = await fetch(p.logoUrl).then((r) => r.arrayBuffer()).then(aBase64).catch(() => undefined);
      const wb = await generarExcel(p.mes, p.config, p.reporte, { logoBase64: logo });
      const buf = await excelABuffer(wb);
      const periodo = nombrePeriodo(p.mes.inicio, p.mes.hasta);
      const nombre = `Asistencia ${periodo.charAt(0) + periodo.slice(1).toLowerCase()} (del ${ddmm(p.mes.inicio).replace('/', '-')} al ${ddmm(p.mes.hasta).replace('/', '-')}).xlsx`;
      const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = nombre;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch (e) {
      setErrorDescarga(`No se pudo generar el Excel: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setGenerando(false);
    }
  }

  return (
    <div className="space-y-6">
      {/* Paso 1 */}
      <section className="rounded-xl bg-white p-5 shadow-sm">
        <Paso n={1} titulo="Sube el reporte del reloj" />
        <div
          onDragOver={(e) => { e.preventDefault(); setArrastrando(true); }}
          onDragLeave={() => setArrastrando(false)}
          onDrop={soltar}
          onClick={() => input.current?.click()}
          className={`mt-3 cursor-pointer rounded-xl border-2 border-dashed p-8 text-center transition ${arrastrando ? 'border-qalen bg-qalen-soft' : 'border-stone-300 hover:border-qalen hover:bg-stone-50'}`}
        >
          <p className="text-lg font-semibold">{p.reporte ? `📄 ${p.nombreArchivo}` : 'Arrastra aquí el archivo .xls del reloj'}</p>
          <p className="mt-1 text-sm text-stone-500">{p.reporte ? 'Para cambiarlo, arrastra otro o haz clic.' : 'o haz clic para buscarlo en tu computadora'}</p>
          <input ref={input} type="file" accept=".xls,.xlsx" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) p.onArchivo(f); e.target.value = ''; }} />
        </div>
        {p.error && <Caja tipo="error">{p.error}</Caja>}

        {p.reporte && p.mes && (
          <div className="mt-4 grid gap-3 sm:grid-cols-4">
            <Dato titulo="Período del reporte" valor={`${ddmm(p.reporte.inicio)} al ${ddmm(p.reporte.fin)}`} />
            <Dato titulo="Personas en el reloj" valor={String(p.reporte.personas.length)} sub={`${p.mes.filas.length} activas`} />
            <label className="rounded-lg bg-stone-50 p-3 sm:col-span-2">
              <span className="block text-xs font-semibold uppercase text-stone-500">Calcular hasta el día</span>
              <input type="date" value={p.hasta} min={p.reporte.inicio} max={p.reporte.fin} onChange={(e) => e.target.value && p.setHasta(e.target.value)} className="mt-1" />
              <span className="mt-1 block text-xs text-stone-500">Por defecto: todo el rango del reporte (aunque abarque varios meses). Puedes elegir un día anterior para ver cómo va.</span>
            </label>
          </div>
        )}
        {p.avisos.length > 0 && (
          <Caja tipo="aviso">
            <p className="font-semibold">Ten en cuenta:</p>
            <ul className="ml-5 list-disc">{p.avisos.map((a, i) => <li key={i}>{a}</li>)}</ul>
          </Caja>
        )}
      </section>

      {p.mes && (
        <>
          {/* Paso 2 */}
          <section className="rounded-xl bg-white p-5 shadow-sm">
            <Paso n={2} titulo="Revisa cómo quedó" />
            <div className="mt-3 grid gap-3 sm:grid-cols-3">
              <Kpi color="bg-yellow-300" valor={pendientes} texto="datos por completar" ayuda="Quedan en amarillo en el Excel." />
              <Kpi color="bg-orange-200" valor={paraMirar} texto="cosas para mirar" ayuda="Horas raras o muchas extras. No bloquean." />
              <Kpi color="bg-stone-200" valor={p.mes.filas.length} texto="personas" ayuda={`${p.mes.fechas.length} días en el período`} />
            </div>

            <div className="mt-4 overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-qalen text-left text-xs uppercase text-white">
                    <th className="p-2">Nombre</th>
                    <th className="p-2 text-right">Faltas</th>
                    <th className="p-2 text-right">Extras {pct1}</th>
                    <th className="p-2 text-right">Extras {pct2}</th>
                    <th className="p-2 text-right">Compen&shy;sadas</th>
                    <th className="p-2 text-right">A pagar {pct1}</th>
                    <th className="p-2 text-right">A pagar {pct2}</th>
                    <th className="p-2 text-right">Dom./fer. {pctF}</th>
                    <th className="p-2 text-right">Pago extras</th>
                    <th className="p-2 text-right">Sueldo + extras</th>
                    <th className="p-2 text-right">Nos debe</th>
                    <th className="p-2 text-right">En contra</th>
                    <th className="p-2 text-center">Por completar</th>
                    <th className="p-2" />
                  </tr>
                </thead>
                <tbody>
                  {p.mes.filas.map((f) => {
                    const ab = abierto === f.persona.id;
                    return (
                      <Fragment key={f.persona.id}>
                        <tr className="cursor-pointer border-b border-stone-200 hover:bg-stone-50" onClick={() => setAbierto(ab ? null : f.persona.id)}>
                          <td className="p-2 font-semibold">{f.persona.alias}<span className="ml-2 text-xs font-normal text-stone-400">{f.persona.empresa}</span></td>
                          <td className={`p-2 text-right ${f.tot.faltas ? 'font-bold text-red-700' : 'text-stone-300'}`}>{f.tot.faltas || '—'}</td>
                          <td className="p-2 text-right text-green-700">{hm(f.tot.extra1)}</td>
                          <td className="p-2 text-right text-green-700">{hm(f.tot.extra2)}</td>
                          <td className="p-2 text-right font-semibold text-purple-700">{hm(f.tot.compensado)}</td>
                          <td className="p-2 text-right font-semibold text-green-800">{hm(f.tot.neto1)}</td>
                          <td className="p-2 text-right font-semibold text-green-800">{hm(f.tot.neto2)}</td>
                          <td className="p-2 text-right font-semibold text-sky-800">{hm(f.tot.enFeriado)}</td>
                          <td className="p-2 text-right font-bold">
                            {f.persona.sueldo > 0 ? (
                              <>
                                <span className="text-green-800">{soles(f.tot.monto)}</span>
                                {f.tot.montoOcho !== f.tot.montoReales && (
                                  <span className="block whitespace-nowrap text-xs font-normal text-stone-500" title="El mismo pago con la otra forma de calcular el valor hora">
                                    {f.persona.valorHoraSegun === 'reales' ? '30 d × 8 h' : 'horas reales'}: {soles(f.persona.valorHoraSegun === 'reales' ? f.tot.montoOcho : f.tot.montoReales)}
                                  </span>
                                )}
                              </>
                            ) : (
                              <span className="rounded bg-yellow-200 px-1 text-xs font-semibold text-amber-900" title="Falta el sueldo (Personas y reglas o Resumen del Excel)">sin sueldo</span>
                            )}
                          </td>
                          <td className="whitespace-nowrap p-2 text-right">{f.persona.sueldo > 0 ? soles(f.tot.totalMes) : ''}</td>
                          <td className="p-2 text-right font-bold text-red-700">{hm(f.tot.nosDebe)}</td>
                          <td className="p-2 text-right text-red-700">{hm(f.tot.contra)}</td>
                          <td className="p-2 text-center">{f.tot.pendientes ? <span className="rounded bg-yellow-300 px-2 font-bold">{f.tot.pendientes}</span> : <span className="text-green-700">✔</span>}</td>
                          <td className="p-2 text-right text-stone-400">{ab ? '▲' : '▼ días'}</td>
                        </tr>
                        {ab && (
                          <tr>
                            <td colSpan={14} className="bg-stone-50 p-2">
                              <DetallePersona fila={f} pct1={pct1} pct2={pct2} />
                            </td>
                          </tr>
                        )}
                      </Fragment>
                    );
                  })}
                </tbody>
              </table>
              <p className="mt-2 text-xs text-stone-500">
                Cada día, las primeras {p.config.reglas.tramo1Min / 60} h extras van al {pct1} y el resto al {pct2}. Las compensadas se descuentan de las extras de los
                primeros días (primero las del {pct2}); si no alcanzan, aparece en “Nos debe”. Las horas en contra van aparte. Los días por completar cuentan como 0
                hasta que se llenen en el Excel.
              </p>
            </div>
          </section>

          {/* Paso 3 */}
          <section className="rounded-xl bg-white p-5 shadow-sm">
            <Paso n={3} titulo="Descarga el Excel y complétalo" />
            <p className="mt-2 text-sm text-stone-600">
              El Excel tiene el reporte del reloj y, debajo de cada día, la entrada y la salida. Completa las <b className="bg-yellow-300 px-1">celdas amarillas</b>{' '}
              (hora o motivo: Vacaciones, Salud, Compensado, Estudio o Falta) y el Resumen se actualiza solo.
            </p>
            {p.erroresConfig.length > 0 && (
              <Caja tipo="error">
                <p className="font-semibold">Antes de descargar, corrige esto en “Personas y reglas”:</p>
                <ul className="ml-5 list-disc">{p.erroresConfig.map((e, i) => <li key={i}>{e}</li>)}</ul>
                <button className="mt-2 font-semibold underline" onClick={p.irAPersonas}>Ir a Personas y reglas</button>
              </Caja>
            )}
            {sinSueldo.length > 0 && (
              <Caja tipo="aviso">
                Falta el sueldo de: <b>{sinSueldo.join(', ')}</b>. Sin él no se calcula el pago en soles. Puedes ponerlo en{' '}
                <button className="font-semibold underline" onClick={p.irAPersonas}>Personas y reglas</button> o después en el Resumen del Excel (celdas amarillas).
              </Caja>
            )}
            {errorDescarga && <Caja tipo="error">{errorDescarga}</Caja>}
            <button
              disabled={generando || p.erroresConfig.length > 0 || p.mes.filas.length === 0}
              onClick={() => (pendientes > 0 ? setConfirmar(true) : descargar())}
              className="mt-4 rounded-lg bg-qalen px-6 py-3 text-lg font-bold text-white shadow hover:bg-qalen-dark disabled:cursor-not-allowed disabled:opacity-50"
            >
              {generando ? 'Generando…' : '⬇ Descargar Excel'}
            </button>
          </section>
        </>
      )}

      {confirmar && (
        <div className="fixed inset-0 z-10 flex items-center justify-center bg-black/40 p-4" onClick={() => setConfirmar(false)}>
          <div className="max-w-md rounded-xl bg-white p-6 shadow-xl" onClick={(e) => e.stopPropagation()}>
            <h3 className="text-lg font-bold">Hay {pendientes} datos por completar</h3>
            <p className="mt-2 text-sm text-stone-600">
              Quedarán en <b className="bg-yellow-300 px-1">amarillo</b> en la hoja Asistencia (pasa el mouse por la fila Reloj para ver qué pasó). Mientras no los completes, el
              Resumen los cuenta como 0.
            </p>
            <div className="mt-5 flex justify-end gap-2">
              <button className="rounded-lg px-4 py-2 font-semibold text-stone-600 hover:bg-stone-100" onClick={() => setConfirmar(false)}>Cancelar</button>
              <button className="rounded-lg bg-qalen px-4 py-2 font-bold text-white hover:bg-qalen-dark" onClick={descargar}>Entendido, descargar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function DetallePersona({ fila, pct1, pct2 }: { fila: Mes['filas'][number]; pct1: string; pct2: string }) {
  return (
    <table className="w-full text-xs">
      <thead>
        <tr className="text-left text-stone-500">
          <th className="p-1">Fecha</th>
          <th className="p-1">Reloj</th>
          <th className="p-1">Entrada</th>
          <th className="p-1">Salida</th>
          <th className="p-1 text-right">Trabajado</th>
          <th className="p-1 text-right">Extra {pct1}</th>
          <th className="p-1 text-right">Extra {pct2}</th>
          <th className="p-1 text-right">Dom./fer.</th>
          <th className="p-1 text-right">En contra</th>
          <th className="p-1 text-right">Compensadas</th>
          <th className="p-1 text-right">S/ {pct1}</th>
          <th className="p-1 text-right">S/ {pct2}</th>
          <th className="p-1 text-right">S/ día</th>
          <th className="p-1">Nota</th>
        </tr>
      </thead>
      <tbody>
        {fila.dias.map((d) => {
          const falta = d.casos.some((c) => c.faltaDato);
          const fin = d.tipo === 'DOM' || d.tipo === 'FER';
          return (
            <tr key={d.fecha} className={`border-t border-stone-200 ${!d.aplica ? 'text-stone-300' : fin ? 'bg-rose-50' : ''}`}>
              <td className="whitespace-nowrap p-1">{DIAS_CORTO[weekday(d.fecha)]} {ddmm(d.fecha)}{d.tipo === 'FER' && <span className="ml-1 text-red-600">feriado</span>}</td>
              <td className="p-1 font-mono text-stone-500">{d.reloj || '—'}</td>
              <td className={`p-1 font-mono ${falta && d.entrada == null && d.aplica ? 'bg-yellow-300' : ''}`}>{fmtHHMM(d.entrada)}</td>
              <td className={`p-1 font-mono ${falta && d.salida == null && d.aplica ? 'bg-yellow-300' : ''}`}>{fmtHHMM(d.salida)}</td>
              <td className="p-1 text-right font-mono">{d.calc.trabajado ? fmtDur(d.calc.trabajado) : ''}</td>
              <td className="p-1 text-right font-mono text-green-700">{hm(d.calc.extra1)}</td>
              <td className="p-1 text-right font-mono text-green-800">{hm(d.calc.extra2)}</td>
              <td className="p-1 text-right font-mono text-sky-800">{hm(d.calc.enFeriado)}</td>
              <td className="p-1 text-right font-mono text-red-700">{hm(d.calc.contra)}</td>
              <td className="p-1 text-right font-mono text-purple-700">{hm(d.calc.compensado)}</td>
              <td className="whitespace-nowrap p-1 text-right font-mono text-amber-800">{d.soles.s1 ? d.soles.s1.toFixed(2) : ''}</td>
              <td className="whitespace-nowrap p-1 text-right font-mono text-amber-800">{d.soles.s2 ? d.soles.s2.toFixed(2) : ''}</td>
              <td className="whitespace-nowrap p-1 text-right font-mono font-bold text-amber-900">{d.soles.total ? d.soles.total.toFixed(2) : ''}</td>
              <td className="p-1 text-stone-600">{!d.aplica ? 'Antes de su fecha de ingreso' : d.casos.map((c) => c.mensaje + (c.sugerencia ? ` (sugerencia: ${c.sugerencia})` : '')).join(' ')}</td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

/** h:mm, vacío si es cero. */
function hm(m: number): string {
  return m ? fmtDur(m) : '';
}

function soles(n: number): string {
  return `S/ ${n.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function Paso({ n, titulo }: { n: number; titulo: string }) {
  return (
    <h2 className="flex items-center gap-3 text-lg font-bold">
      <span className="flex h-8 w-8 items-center justify-center rounded-full bg-qalen text-white">{n}</span>
      {titulo}
    </h2>
  );
}

function Dato({ titulo, valor, sub }: { titulo: string; valor: string; sub?: string }) {
  return (
    <div className="rounded-lg bg-stone-50 p-3">
      <span className="block text-xs font-semibold uppercase text-stone-500">{titulo}</span>
      <span className="text-lg font-bold">{valor}</span>
      {sub && <span className="ml-2 text-sm text-stone-500">{sub}</span>}
    </div>
  );
}

function Kpi({ color, valor, texto, ayuda }: { color: string; valor: number; texto: string; ayuda: string }) {
  return (
    <div className="flex items-center gap-3 rounded-lg border border-stone-200 p-3">
      <span className={`flex h-12 min-w-12 items-center justify-center rounded-lg px-2 text-2xl font-bold ${color}`}>{valor}</span>
      <div>
        <p className="font-semibold">{texto}</p>
        <p className="text-xs text-stone-500">{ayuda}</p>
      </div>
    </div>
  );
}

export function Caja({ tipo, children }: { tipo: 'error' | 'aviso' | 'ok'; children: React.ReactNode }) {
  const c = tipo === 'error' ? 'border-red-300 bg-red-50 text-red-800' : tipo === 'aviso' ? 'border-amber-300 bg-amber-50 text-amber-900' : 'border-green-300 bg-green-50 text-green-800';
  return <div className={`mt-3 rounded-lg border p-3 text-sm ${c}`}>{children}</div>;
}

function aBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
