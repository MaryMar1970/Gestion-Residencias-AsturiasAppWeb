import { useState, useEffect, useCallback, useRef } from 'react';
import './calendar.css';
import { apiFetch } from './api';

// ─── Types ─────────────────────────────────────────────────────────────────────
type Residencia = { id: string; nombre: string; totalHabitaciones: number };
type Habitacion = { id: string; residenciaId: string; residenciaNombre: string; tipoHabitacionId: string; tipoNombre: string; tipoCodigo: string; numero: string; nombre?: string };
type Tarifa    = { id: string; residenciaId: string; residenciaNombre: string; tipoHabitacionId: string; tipoHabitacionNombre: string; nombreTarifa: string; precioNoche: number; precioMes: number; porcentajeIva: number };
type Huesped   = { id: string; dni: string; nombre: string; apellidos: string; nombreCompleto: string; telefono?: string; email?: string; tipoHuesped: string; enListaNegra: boolean; totalReservas: number; direccion?: string; codigoPostal?: string; municipio?: string; provincia?: string; centroOrigen?: string; departamento?: string; empleo?: string; situacion?: string; finalidad?: string; empleoCategoria?: string; motivoListaNegra?: string; notas?: string };
type Reserva   = {
  id: string; numeroOrden: number; habitacionId: string; habitacionNumero: string; residenciaNombre: string; tipoHabitacionNombre: string;
  huespedId?: string; huespedNombreCompleto?: string; huespedDni?: string; huespedNombre?: string; huespedApellidos?: string;
  fechaEntrada: string; fechaSalida: string; totalNoches: number; numPersonas: number;
  estado: string; estadoInt: number; esBloqueo: boolean; motivoBloqueo?: string;
  tarifaNombreSnapshot?: string; precioNocheAplicado: number; importeTotal: number;
  pagado: boolean; formaPago?: string; observaciones?: string; finalidad?: string; empleo?: string; evaluacion?: string;
};
type CalHabitacion = { habitacionId: string; numero: string; residenciaNombre: string; tipoNombre: string; tipoCodigo: string; activa: boolean; orden: number; reservas: Reserva[] };
type CalData = { fechaInicio: string; fechaFin: string; habitaciones: CalHabitacion[] };
type Festivo   = { id: string; fecha: string; descripcion: string };

// ─── Helpers ───────────────────────────────────────────────────────────────────
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const DIAS_LONG = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
const ESTADO_COLOR: Record<string, string> = {
  Confirmada: 'confirmada', Pendiente: 'pendiente',
  CheckIn: 'checkin', CheckOut: 'checkout', Bloqueada: 'bloqueada'
};

function addDays(date: Date, n: number): Date {
  const d = new Date(date); d.setDate(d.getDate() + n); return d;
}
function toDateStr(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function parseLocal(s: string): Date {
  const [y,m,d] = s.split('-').map(Number);
  return new Date(y, m-1, d);
}
function diasEntre(desde: string, hasta: string): Date[] {
  const out: Date[] = [];
  const cur = parseLocal(desde);
  const end = parseLocal(hasta);
  while (cur < end) { out.push(new Date(cur)); cur.setDate(cur.getDate() + 1); }
  return out;
}

export function calcularEvaluacion(finalidad: string, empleoCategoria: string, situacion: string): string {
  const f = (finalidad || 'Otros').trim().toLowerCase();
  let finalidadMapeada = 'Otros';
  if (f.includes('comision no') || f.includes('comisión no') || f.includes('indem')) {
    finalidadMapeada = 'Comisión NO indem.';
  } else if (f.includes('comision') || f.includes('comisión')) {
    finalidadMapeada = 'Comisión';
  } else if (f.includes('destino')) {
    finalidadMapeada = 'Destino';
  } else if (f.includes('enfermedad')) {
    finalidadMapeada = 'Enfermedad';
  } else if (f.includes('urgencia')) {
    finalidadMapeada = 'Urgencia';
  } else if (f.includes('sepelio')) {
    finalidadMapeada = 'Sepelio';
  } else if (f.includes('estancia')) {
    finalidadMapeada = 'Máx. Estancia';
  }

  const e = (empleoCategoria || 'GC').trim().toLowerCase();
  let empleoCat = 'GC';
  if (e.includes('alumno')) {
    empleoCat = 'Alumno';
  } else if (e.includes('funcionario')) {
    empleoCat = 'Funcionario en GC';
  } else if (e.includes('militar en')) {
    empleoCat = 'Militar en GC';
  } else if (e.includes('militar no') || e.includes('militar')) {
    empleoCat = 'Militar no  GC';
  }

  const s = (situacion || 'Activo').trim().toLowerCase();
  let situacionMapeada = 'Activo';
  if (s.includes('viogen')) {
    situacionMapeada = 'Viogen';
  } else if (s.includes('asoc')) {
    situacionMapeada = 'Asociación';
  } else if (s.includes('reserva activo') || s.includes('reserva activa')) {
    situacionMapeada = 'Reserva activo';
  } else if (s.includes('reserva')) {
    situacionMapeada = 'Reserva';
  } else if (s.includes('excedencia')) {
    situacionMapeada = 'Excedencia';
  } else if (s.includes('especial')) {
    situacionMapeada = 'Especiales';
  } else if (s.includes('retirado') || s.includes('jubilado')) {
    situacionMapeada = 'Retirado';
  } else if (s.includes('viuda')) {
    situacionMapeada = 'Viuda';
  } else if (s.includes('huerfano') || s.includes('huérfano')) {
    situacionMapeada = 'Huerfano';
  }

  const clave = `${finalidadMapeada.toLowerCase()}|${empleoCat.toLowerCase()}|${situacionMapeada.toLowerCase()}`;

  const matriz: Record<string, string> = {
    'comisión no indem.|gc|viogen': '1, 1, 1',
    'comisión no indem.|gc|activo': '1, 1, 2',
    'comisión no indem.|gc|reserva activo': '1, 1, 2',
    'destino|gc|activo': '1, 1, 3',
    'comisión|gc|activo': '1, 1, 4',
    'enfermedad|gc|retirado': '2, 8, 1',
    'comisión|alumno|activo': '1, 2, 2',
    'comisión|militar en gc|activo': '1, 3, 2',
    'destino|militar en gc|activo': '1, 3, 2',
    'comisión|funcionario en gc|activo': '1, 4, 2',
    'destino|funcionario en gc|activo': '1, 4, 2',
    'enfermedad|gc|activo': '2, 1, 1',
    'otros|gc|viogen': '2, 1, 2',
    'urgencia|gc|activo': '2, 1, 3',
    'sepelio|gc|activo': '2, 1, 4',
    'máx. estancia|gc|activo': '2, 1, 5',
    'otros|gc|asociación': '2, 1, 6',
    'otros|gc|activo': '2, 1, 7',
    'otros|gc|reserva activo': '2, 1, 8',
    'otros|alumno|activo': '2, 2, 7',
    'otros|gc|reserva': '2, 3, 7',
    'otros|militar en gc|activo': '2, 4, 7',
    'otros|funcionario en gc|activo': '2, 5, 7',
    'otros|gc|excedencia': '2, 6, 7',
    'otros|gc|especiales': '2, 7, 7',
    'otros|gc|retirado': '2, 8, 7',
    'otros|gc|viuda': '2, 9, 7',
    'otros|gc|huerfano': '2, 9, 7',
    'otros|militar no  gc|activo': '2, 10, 7',
    'otros|militar no  gc|reserva': '2, 10, 7',
    'otros|militar no  gc|retirado': '2, 10, 7'
  };

  return matriz[clave] || '(NO VÁLIDO)';
}

// ─── CALENDARIO GANTT ─────────────────────────────────────────────────────────
export function CalendarioGantt({ onNuevaReserva }: { onNuevaReserva: (habitacionId: string, fecha: string) => void }) {
  const hoy = new Date();
  const [inicio, setInicio] = useState(() => {
    const gotoFecha = sessionStorage.getItem('gotoReservaFecha');
    if (gotoFecha) {
      const checkinDate = parseLocal(gotoFecha);
      const startView = addDays(checkinDate, -2);
      return toDateStr(startView);
    }
    const d = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    return toDateStr(d);
  });
  const [dias, setDias] = useState(() => {
    const gotoId = sessionStorage.getItem('gotoReservaId');
    return gotoId ? 30 : 10;
  });
  const [residenciaFiltro, setResidenciaFiltro] = useState('');
  const [calData, setCalData] = useState<CalData | null>(null);
  const [festivos, setFestivos] = useState<Set<string>>(new Set());
  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [loading, setLoading] = useState(false);
  const [tooltip, setTooltip] = useState<{ reserva: Reserva; x: number; y: number } | null>(null);
  const [reservaSeleccionada, setReservaSeleccionada] = useState<Reserva | null>(null);
  const [highlightReservaId, setHighlightReservaId] = useState<string | null>(null);
  const tooltipTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const fin = toDateStr(addDays(parseLocal(inicio), dias));
  const fechas = diasEntre(inicio, fin);
  const hoyStr = toDateStr(hoy);

  const cargar = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams({ fechaInicio: inicio, fechaFin: fin });
      if (residenciaFiltro) params.set('residenciaId', residenciaFiltro);
      const [cal, fest] = await Promise.all([
        apiFetch<CalData>(`/api/calendario?${params}`),
        apiFetch<Festivo[]>(`/api/festivos?anio=${new Date(inicio).getFullYear()}`),
      ]);
      setCalData(cal);
      setFestivos(new Set(fest.map(f => f.fecha)));
    } catch { } finally { setLoading(false); }
  }, [inicio, fin, residenciaFiltro]);

  useEffect(() => {
    apiFetch<Residencia[]>('/api/residencias').then(setResidencias).catch(() => {});
  }, []);

  const pendingSelectIdRef = useRef<string | null>(sessionStorage.getItem('gotoReservaId'));

  useEffect(() => {
    sessionStorage.removeItem('gotoReservaId');
    sessionStorage.removeItem('gotoReservaFecha');
  }, []);

  useEffect(() => {
    if (calData && pendingSelectIdRef.current) {
      const targetId = pendingSelectIdRef.current;
      pendingSelectIdRef.current = null;

      let found: Reserva | undefined;
      for (const hab of calData.habitaciones) {
        found = hab.reservas.find(r => r.id === targetId);
        if (found) break;
      }
      if (found) {
        setHighlightReservaId(found.id);
        setTimeout(() => {
          const el = document.getElementById(`room-label-${found.habitacionId}`);
          if (el) {
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          }
        }, 150);
        setTimeout(() => {
          setHighlightReservaId(null);
        }, 4000);
      }
    }
  }, [calData]);

  useEffect(() => { void cargar(); }, [cargar]);

  // Calcular grupos de meses para la cabecera
  const mesesHeader: { label: string; count: number }[] = [];
  fechas.forEach(d => {
    const label = `${MESES[d.getMonth()]} ${d.getFullYear()}`;
    const last = mesesHeader[mesesHeader.length - 1];
    if (last?.label === label) last.count++;
    else mesesHeader.push({ label, count: 1 });
  });

  const colDays = fechas.length;
  const colWidth = dias <= 5 ? 'minmax(140px, 1fr)' : dias <= 10 ? 'minmax(124px, 1fr)' : dias <= 15 ? 'minmax(92px, 1fr)' : 'minmax(60px, 1fr)';
  const gridCols = `96px repeat(${colDays}, ${colWidth})`;

  return (
    <div>
      {/* Toolbar */}
      <div className="cal-toolbar">
        <h4>📅 Calendario de ocupación</h4>
        <select className="form-group select" value={residenciaFiltro} onChange={e => setResidenciaFiltro(e.target.value)}
          style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
          <option value="">Todas las residencias</option>
          {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
        </select>
        <input type="date" value={inicio} onChange={e => setInicio(e.target.value)}
          style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }} />
        <select value={dias} onChange={e => setDias(+e.target.value)}
          style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
          <option value={5}>5 días</option>
          <option value={10}>10 días</option>
          <option value={15}>15 días</option>
          <option value={30}>30 días</option>
        </select>
        <button className="btn btn-ghost btn-sm" onClick={() => setInicio(toDateStr(new Date()))}>Hoy</button>
        <button className="btn btn-ghost btn-sm" onClick={() => setInicio(toDateStr(addDays(parseLocal(inicio), -1)))}>◀</button>
        <button className="btn btn-ghost btn-sm" onClick={() => setInicio(toDateStr(addDays(parseLocal(inicio), 1)))}>▶</button>
        <button className="btn btn-primary btn-sm" style={{ marginLeft: '24px' }} onClick={() => onNuevaReserva('', '')}>+ Nueva reserva</button>
        {loading && <span className="spinner"></span>}
      </div>

      {/* Grid */}
      <div className="cal-wrapper" style={{ borderRadius: 0, border: 'none' }}>
        <div className="cal-grid" style={{ display: 'grid', gridTemplateColumns: gridCols, minWidth: '100%' }}>
          {/* Fila meses */}
          <div className="cal-head-room" style={{
            gridRow: 1,
            gridColumn: 1,
            position: 'sticky',
            left: 0,
            zIndex: 10,
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center',
            alignItems: 'center',
            textAlign: 'center',
            lineHeight: '1.2',
            padding: '2px 0',
          }}>
            {(() => {
              if (!residenciaFiltro) {
                return (
                  <>
                    <span>HABITACIÓN</span>
                    <span>APARTAMENTO</span>
                  </>
                );
              }
              const name = residencias.find(r => r.id === residenciaFiltro)?.nombre ?? '';
              if (name.toLowerCase().includes('gijón')) {
                return <span>HABITACIÓN</span>;
              }
              if (name.toLowerCase().includes('soto')) {
                return <span>APARTAMENTO</span>;
              }
              return <span>HABITACIÓN</span>;
            })()}
          </div>
          {mesesHeader.map((m, i) => {
            const offset = mesesHeader.slice(0, i).reduce((s, x) => s + x.count, 0);
            return (
              <div key={m.label} className="cal-head-month"
                style={{ gridColumn: `${offset + 2} / span ${m.count}`, gridRow: 1 }}>
                {m.label}
              </div>
            );
          })}

          {/* Fila días */}
          <div style={{ gridRow: 2, gridColumn: 1, position: 'sticky', left: 0, zIndex: 10, background: 'var(--primary-dark)', borderRight: '2px solid rgba(255,255,255,0.15)' }}></div>
          {fechas.map((d, i) => {
            const ds = toDateStr(d);
            const dow = d.getDay();
            const isWeekend = dow === 0 || dow === 6;
            const isFestivo = festivos.has(ds);
            const isToday = ds === hoyStr;
            return (
              <div key={ds} className={`cal-head-day ${isWeekend ? 'weekend' : ''} ${isFestivo ? 'festivo' : ''} ${isToday ? 'today' : ''}`}
                style={{ gridColumn: i + 2, gridRow: 2, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 10, fontWeight: 700 }} title={ds}>
                {DIAS_LONG[dow]} {d.getDate()}
              </div>
            );
          })}

          {/* Filas habitaciones */}
          {(calData?.habitaciones ?? []).map((hab, rowIdx) => {
            const row = rowIdx + 3;
            return (
              <div key={hab.habitacionId} className="cal-row">
                {/* Etiqueta habitación */}
                <div id={`room-label-${hab.habitacionId}`} className="cal-room-label" style={{ gridRow: row, gridColumn: 1, position: 'sticky', left: 0, zIndex: 5 }}>
                  <div className="room-num">🔑 {hab.numero}</div>
                  <div className="room-res">{hab.residenciaNombre}</div>
                  <div className="room-tipo"><span className="badge badge-primary" style={{ fontSize: 9, padding: '1px 5px' }}>{hab.tipoCodigo}</span></div>
                </div>

                {/* Celdas vacías clicables */}
                {fechas.map((d, ci) => {
                  const ds = toDateStr(d);
                  const dow = d.getDay();
                  const isWeekend = dow === 0 || dow === 6;
                  const isFestivo = festivos.has(ds);
                  const isToday = ds === hoyStr;
                  return (
                    <div key={ds} className={`cal-cell ${isWeekend ? 'weekend' : ''} ${isFestivo ? 'festivo' : ''} ${isToday ? 'today' : ''}`}
                      style={{ gridRow: row, gridColumn: ci + 2 }}
                      onClick={() => onNuevaReserva(hab.habitacionId, ds)}
                      title={`${hab.numero} — ${ds}`}
                    />
                  );
                })}

                {/* Barras de reservas */}
                {hab.reservas.map(res => {
                  const resStart = parseLocal(res.fechaEntrada);
                  const resEnd   = parseLocal(res.fechaSalida);
                  const calStart = parseLocal(inicio);
                  const calEnd   = parseLocal(fin);

                  const startDraw = resStart < calStart ? calStart : resStart;
                  const endDraw   = resEnd   > calEnd   ? calEnd   : resEnd;

                  const colStart = fechas.findIndex(d => toDateStr(d) === toDateStr(startDraw));
                  const colEnd   = fechas.findIndex(d => toDateStr(d) === toDateStr(endDraw));
                  if (colStart === -1) return null;
                  const spanCols = colEnd === -1 ? fechas.length - colStart : (colEnd - colStart + 1);
                  if (spanCols <= 0) return null;

                  const isStartOffscreen = resStart < calStart;
                  const isEndOffscreen = resEnd > calEnd;

                  const leftPercent = (res.esBloqueo || isStartOffscreen) ? 0 : (35 / spanCols);
                  const rightPercent = (res.esBloqueo || isEndOffscreen) ? 0 : (80 / spanCols);

                  const getTitleText = (r: any) => {
                    if (r.esBloqueo) return `🔒 ${r.motivoBloqueo || 'Bloqueado'}`;
                    if (!r.huespedId) return `📋 ${r.motivoBloqueo || 'Reservado (Interno)'}`;
                    return `#${r.numeroOrden} - ${r.huespedNombre || ''} ${r.huespedApellidos || ''} (${r.numPersonas} pax)`;
                  };
                  const titleText = getTitleText(res);

                  return (
                    <div key={res.id}
                      style={{
                        gridRow: row,
                        gridColumn: `${colStart + 2} / span ${spanCols}`,
                        position: 'relative', zIndex: res.id === highlightReservaId ? 100 : 3,
                        height: '100%',
                      }}
                    >
                      <div
                        className={`cal-reserva ${ESTADO_COLOR[res.estado] || 'confirmada'}`}
                        style={{
                          position: 'absolute',
                          left: `${leftPercent}%`,
                          right: `${rightPercent}%`,
                          top: '1px',
                          bottom: '1px',
                          margin: 0,
                          display: 'flex',
                          alignItems: 'center',
                          padding: '0 6px',
                          overflow: 'hidden',
                          border: res.id === highlightReservaId ? '2.5px solid #f59e0b' : undefined,
                          boxShadow: res.id === highlightReservaId ? '0 0 15px rgba(245, 158, 11, 0.9)' : undefined,
                          transform: res.id === highlightReservaId ? 'scale(1.05)' : undefined,
                          transition: 'all 0.3s ease',
                        }}
                        title={titleText}
                        onMouseEnter={e => {
                          if (tooltipTimer.current) clearTimeout(tooltipTimer.current);
                          const rect = (e.target as HTMLElement).getBoundingClientRect();
                          setTooltip({ reserva: res, x: rect.left, y: rect.bottom + 6 });
                        }}
                        onMouseLeave={() => {
                          if (tooltipTimer.current) clearTimeout(tooltipTimer.current);
                          tooltipTimer.current = setTimeout(() => setTooltip(null), 200);
                        }}
                        onClick={e => { e.stopPropagation(); setReservaSeleccionada(res); }}
                      >
                        {res.esBloqueo ? (
                          <span style={{ fontSize: '10.5px', fontWeight: 600 }}>🔒 {res.motivoBloqueo || 'Bloqueado'}</span>
                        ) : !res.huespedId ? (
                          <span style={{ fontSize: '10.5px', fontWeight: 600, color: '#744210' }}>📋 {res.motivoBloqueo || 'Reservado (Interno)'}</span>
                        ) : (
                          <div style={{ display: 'flex', alignItems: 'center', width: '100%', height: '100%', overflow: 'hidden' }}>
                            {/* Número de orden grande centrado a la izquierda */}
                            <div style={{
                              fontSize: '16px',
                              fontWeight: 900,
                              color: '#451a03',
                              marginRight: '6px',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              borderRight: '1px solid rgba(116,66,16,0.15)',
                              paddingRight: '6px',
                              height: '100%',
                              minWidth: '32px',
                            }}>
                              {res.numeroOrden}
                            </div>
                            {/* Info de huésped a la derecha */}
                            <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', lineHeight: '1.2', overflow: 'hidden' }}>
                              <div style={{ fontSize: '11.5px', fontWeight: 800, color: '#744210', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                {res.huespedNombre ? `${res.huespedNombre} ` : ''}
                                {res.huespedApellidos ? `${res.huespedApellidos.trim().split(' ')[0]}` : ''}
                              </div>
                              <div style={{ fontSize: '10px', fontWeight: 700, color: '#744210', opacity: 0.85, marginTop: '2px' }}>
                                {res.numPersonas} pax
                              </div>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}

          {calData?.habitaciones.length === 0 && (
            <div style={{ gridColumn: `1 / span ${colDays + 1}`, gridRow: 3, padding: 30, textAlign: 'center', color: 'var(--text-muted)' }}>
              Sin habitaciones configuradas. Ve a <strong>Configuración → Habitaciones</strong> para añadirlas.
            </div>
          )}
        </div>

        {/* Leyenda */}
        <div className="cal-legend">
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: '#2d8a4e' }}></div>Confirmada</div>
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: '#d4a017' }}></div>Pendiente</div>
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: '#1a6eb5' }}></div>Check-in</div>
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: '#7c4dff' }}></div>Check-out</div>
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: '#888', backgroundImage: 'repeating-linear-gradient(45deg,#777,#777 2px,#999 2px,#999 5px)' }}></div>Bloqueado</div>
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: 'rgba(201,168,76,0.25)', border: '1px solid var(--accent)' }}></div>Hoy</div>
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: '#fff5f5', border: '1px solid #ffcccc' }}></div>Festivo</div>
        </div>
      </div>

      {/* Tooltip flotante */}
      {tooltip && (
        <div className="reserva-tooltip"
          style={{ position: 'fixed', top: tooltip.y, left: Math.min(tooltip.x, window.innerWidth - 290) }}>
          <h4>{tooltip.reserva.esBloqueo ? '🔒 Bloqueo' : `👤 ${tooltip.reserva.huespedNombreCompleto}`}</h4>
          <p>🏠 {tooltip.reserva.habitacionNumero} — {tooltip.reserva.residenciaNombre}</p>
          <p>📅 {tooltip.reserva.fechaEntrada} → {tooltip.reserva.fechaSalida} ({tooltip.reserva.totalNoches}n)</p>
          {!tooltip.reserva.esBloqueo && <p>💰 {tooltip.reserva.importeTotal.toFixed(2)} € {tooltip.reserva.pagado ? '✓ Pagado' : '⏳ Pendiente'}</p>}
          <p style={{ marginTop: 4, opacity: 0.6, fontSize: 10 }}>Clic para ver detalle</p>
        </div>
      )}

      {/* Modal detalle reserva */}
      {reservaSeleccionada && (
        <ReservaDetalleModal reserva={reservaSeleccionada} onClose={() => { setReservaSeleccionada(null); void cargar(); }} />
      )}
    </div>
  );
}

// ─── MODAL DETALLE RESERVA ─────────────────────────────────────────────────────
function ReservaDetalleModal({ reserva, onClose }: { reserva: Reserva; onClose: () => void }) {
  const [loading, setLoading] = useState(false);

  const cancelar = async () => {
    if (!confirm('¿Cancelar esta reserva? Se mantendrá en el historial.')) return;
    setLoading(true);
    try { await apiFetch(`/api/reservas/${reserva.id}`, { method: 'DELETE' }); onClose(); }
    catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  const marcarPagado = async () => {
    setLoading(true);
    const forma = prompt('Forma de pago: Efectivo / Tarjeta / Bizum / Transferencia', 'Efectivo');
    if (!forma) { setLoading(false); return; }
    try {
      await apiFetch(`/api/reservas/${reserva.id}`, {
        method: 'PUT',
        body: JSON.stringify({
          fechaEntrada: reserva.fechaEntrada, fechaSalida: reserva.fechaSalida,
          numPersonas: reserva.numPersonas, camasSupletorias: 0,
          estado: reserva.estado, esBloqueo: reserva.esBloqueo,
          motivoBloqueo: reserva.motivoBloqueo, tarifaId: null,
          observaciones: reserva.observaciones, pagado: true, formaPago: forma
        })
      });
      onClose();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="modal-header">
          <h3>{reserva.esBloqueo ? '🔒 Bloqueo' : reserva.huespedId ? '📋 Reserva' : '📋 Reserva interna'} — {reserva.habitacionNumero}</h3>
          <button className="modal-close" onClick={onClose}>×</button>
        </div>
        <div className="modal-body">
          <table style={{ width: '100%', fontSize: 13 }}>
            <tbody>
              {!reserva.esBloqueo && reserva.huespedId && <>
                <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)', width: 140 }}>Nº Orden / Registro</td><td><strong>#{reserva.numeroOrden}</strong></td></tr>
                <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Huésped</td><td><strong>{reserva.huespedNombreCompleto ?? '—'}</strong></td></tr>
                <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>DNI</td><td>{reserva.huespedDni ?? '—'}</td></tr>
                <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Finalidad / Motivo</td><td>{reserva.finalidad ?? '—'}</td></tr>
                <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Clasificación / Eval.</td><td><span className="badge badge-primary">{reserva.evaluacion ?? '—'}</span> (Mapeado: {reserva.empleo})</td></tr>
              </>}
              {reserva.esBloqueo && <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Motivo</td><td><strong>{reserva.motivoBloqueo ?? '—'}</strong></td></tr>}
              {!reserva.esBloqueo && !reserva.huespedId && <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Motivo / Causa</td><td><strong>{reserva.motivoBloqueo ?? '—'}</strong></td></tr>}
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Residencia</td><td>{reserva.residenciaNombre}</td></tr>
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Habitación</td><td>{reserva.habitacionNumero} — {reserva.tipoHabitacionNombre}</td></tr>
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Entrada</td><td>{reserva.fechaEntrada}</td></tr>
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Salida</td><td>{reserva.fechaSalida}</td></tr>
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Noches</td><td>{reserva.totalNoches}</td></tr>
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Estado</td><td><span className={`badge estado-${reserva.estado}`}>{reserva.estado}</span></td></tr>
              {!reserva.esBloqueo && reserva.huespedId && <>
                <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Importe</td><td><strong>{reserva.importeTotal.toFixed(2)} €</strong> {reserva.tarifaNombreSnapshot && <span className="text-muted">({reserva.tarifaNombreSnapshot})</span>}</td></tr>
                <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Pago</td><td>{reserva.pagado ? <span className="badge badge-success">✓ Pagado — {reserva.formaPago}</span> : <span className="badge badge-danger">⏳ Pendiente</span>}</td></tr>
              </>}
              {reserva.observaciones && <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Notas</td><td>{reserva.observaciones}</td></tr>}
            </tbody>
          </table>
        </div>
        <div className="modal-footer">
          {!reserva.pagado && !reserva.esBloqueo && reserva.huespedId && (
            <button className="btn btn-accent" onClick={marcarPagado} disabled={loading}>💰 Marcar pagado</button>
          )}
          <button className="btn btn-danger btn-sm" onClick={cancelar} disabled={loading}>✗ Cancelar reserva</button>
          <button className="btn btn-ghost" onClick={onClose}>Cerrar</button>
        </div>
      </div>
    </div>
  );
}

// ─── FORMULARIO NUEVA RESERVA ─────────────────────────────────────────────────
export function NuevaReservaModal({
  habitacionIdInicial, fechaEntradaInicial, onSaved, onCancel, isSolicitud
}: {
  habitacionIdInicial?: string; fechaEntradaInicial?: string;
  onSaved: () => void; onCancel: () => void;
  isSolicitud?: boolean;
}) {
  const [habitaciones, setHabitaciones] = useState<Habitacion[]>([]);
  const [tarifas, setTarifas]           = useState<Tarifa[]>([]);
  const [residencias, setResidencias]   = useState<Residencia[]>([]);
  const [residenciaSeleccionadaId, setResidenciaSeleccionadaId] = useState('');
  const [buscarDni, setBuscarDni]        = useState('');
  const [huesped, setHuesped]            = useState<Huesped | null>(null);
  const [huespedNuevo, setHuespedNuevo]  = useState(false);
  const [solapamiento, setSolapamiento]  = useState('');
  const [loading, setLoading]            = useState(false);
  const [tipoRegistro, setTipoRegistro]  = useState<'huesped' | 'reserva-interna' | 'bloqueo'>('huesped');

  // Formulario Huésped Nuevo
  const [huespedForm, setHuespedForm] = useState({
    dni: '', nombre: '', apellidos: '', telefono: '', email: '',
    direccion: '', codigoPostal: '', municipio: '', provincia: '',
    empleo: 'Guardia', situacion: 'Activo', tipoHuesped: 'Externo',
    centroOrigen: '', departamento: '', enListaNegra: false, motivoListaNegra: '', notas: ''
  });

  const [municipiosSugerencias, setMunicipiosSugerencias] = useState<{ codigoPostal: string; municipio: string; provincia: string }[]>([]);
  const [mostrarSugerencias, setMostrarSugerencias] = useState(false);
  const [sugerenciaOrigen, setSugerenciaOrigen] = useState<'cp' | 'municipio' | null>(null);

  const getLocalDateTimeString = (d = new Date()) => {
    const offset = d.getTimezoneOffset() * 60000;
    return new Date(d.getTime() - offset).toISOString().slice(0, 16);
  };

  const manana = toDateStr(addDays(new Date(), 1));
  const [f, setF] = useState({
    habitacionId: habitacionIdInicial ?? '',
    huespedId: '',
    fechaEntrada: fechaEntradaInicial ?? toDateStr(new Date()),
    fechaSalida: fechaEntradaInicial ? toDateStr(addDays(parseLocal(fechaEntradaInicial), 1)) : manana,
    numPersonas: 1, camasSupletorias: 0,
    esBloqueo: false, motivoBloqueo: '',
    tarifaId: '', observaciones: '',
    finalidad: 'Otros',
    fechaSolicitud: getLocalDateTimeString()
  });

  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));
  const setH = (k: string, v: any) => setHuespedForm(p => ({ ...p, [k]: v }));

  useEffect(() => {
    Promise.all([
      apiFetch<Habitacion[]>('/api/habitaciones'),
      apiFetch<Tarifa[]>('/api/tarifas'),
      apiFetch<Residencia[]>('/api/residencias')
    ]).then(([h, t, r]) => {
      setHabitaciones(h);
      setTarifas(t);
      setResidencias(r);

      let initialResId = '';
      if (habitacionIdInicial) {
        const initialHab = h.find(x => x.id === habitacionIdInicial);
        if (initialHab) {
          initialResId = initialHab.residenciaId;
        }
      }

      if (!initialResId) {
        if (r.length === 1) {
          initialResId = r[0].id;
        } else if (r.length > 1) {
          const last = localStorage.getItem('lastSelectedResidenciaId');
          if (last && r.some(x => x.id === last)) {
            initialResId = last;
          } else {
            initialResId = r[0].id;
          }
        }
      }
      setResidenciaSeleccionadaId(initialResId);
    });
  }, []);

  const handleResidenciaChange = (resId: string) => {
    setResidenciaSeleccionadaId(resId);
    localStorage.setItem('lastSelectedResidenciaId', resId);
    const currentHab = habitaciones.find(hab => hab.id === f.habitacionId);
    if (currentHab && currentHab.residenciaId !== resId) {
      set('habitacionId', '');
    }
  };

  // Verificar solapamiento cuando cambian fechas, habitación o tipo (esBloqueo)
  useEffect(() => {
    if (!f.habitacionId || !f.fechaEntrada || !f.fechaSalida) return;
    apiFetch<{ haySolapamiento: boolean; mensaje?: string }>(
      `/api/reservas/comprobar-solapamiento?habitacionId=${f.habitacionId}&fechaEntrada=${f.fechaEntrada}&fechaSalida=${f.fechaSalida}&esBloqueo=${f.esBloqueo}`
    ).then(r => {
      if (r.haySolapamiento) {
        setSolapamiento(r.mensaje ?? '⚠️ Solapamiento detectado');
        alert(r.mensaje ?? '⚠️ Solapamiento detectado');
      } else {
        setSolapamiento('');
      }
    })
     .catch(() => setSolapamiento(''));
  }, [f.habitacionId, f.fechaEntrada, f.fechaSalida, f.esBloqueo]);

  const buscarHuesped = async () => {
    if (!buscarDni.trim()) return;
    try {
      const h = await apiFetch<Huesped>(`/api/huespedes/dni/${buscarDni.trim()}`);
      if (h.enListaNegra) { alert(`⚠️ ATENCIÓN: Este huésped está en LISTA NEGRA.\nMotivo: ${h.motivoListaNegra}`); }
      setHuesped(h); set('huespedId', h.id);
      setHuespedNuevo(false);
    } catch {
      // DNI no encontrado: mostrar formulario para crear al vuelo
      setHuesped(null);
      set('huespedId', '');
      setHuespedNuevo(true);
      setHuespedForm(prev => ({
        ...prev,
        dni: buscarDni.trim().toUpperCase(),
        nombre: '', apellidos: '', telefono: '', email: '', direccion: '', codigoPostal: '', municipio: '', provincia: '', empleo: 'Guardia', situacion: 'Activo'
      }));
    }
  };

  const handleCpChange = async (cpValue: string) => {
    setH('codigoPostal', cpValue);
    setSugerenciaOrigen('cp');
    if (cpValue.length >= 4) {
      try {
        const res = await apiFetch<{ codigoPostal: string; municipio: string; provincia: string }[]>(`/api/direcciones/buscar?cp=${cpValue}`);
        if (res.length === 1) {
          setHuespedForm(p => ({
            ...p,
            codigoPostal: res[0].codigoPostal,
            municipio: res[0].municipio,
            provincia: res[0].provincia
          }));
          setMostrarSugerencias(false);
        } else if (res.length > 1) {
          setMunicipiosSugerencias(res);
          setMostrarSugerencias(true);
        }
      } catch { /* ignore */ }
    } else {
      setMostrarSugerencias(false);
    }
  };

  const handleMunicipioChange = async (mValue: string) => {
    setH('municipio', mValue);
    setSugerenciaOrigen('municipio');
    if (mValue.length >= 3) {
      try {
        const res = await apiFetch<{ codigoPostal: string; municipio: string; provincia: string }[]>(`/api/direcciones/buscar?municipio=${mValue}`);
        setMunicipiosSugerencias(res);
        setMostrarSugerencias(true);
      } catch { /* ignore */ }
    } else {
      setMostrarSugerencias(false);
    }
  };

  const seleccionarSugerencia = (sug: { codigoPostal: string; municipio: string; provincia: string }) => {
    setHuespedForm(p => ({
      ...p,
      codigoPostal: sug.codigoPostal,
      municipio: sug.municipio,
      provincia: sug.provincia
    }));
    setMostrarSugerencias(false);
  };

  const tarifasFiltradas = tarifas.filter(t => {
    if (residenciaSeleccionadaId && t.residenciaId !== residenciaSeleccionadaId) return false;
    if (f.habitacionId) {
      const hab = habitaciones.find(h => h.id === f.habitacionId);
      if (hab && t.tipoHabitacionId !== hab.tipoHabitacionId) return false;
    }
    return true;
  });

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (solapamiento) { alert(`No se puede crear: ${solapamiento}`); return; }
    setLoading(true);
    try {
      let actualHuespedId = f.huespedId;
      if (!f.esBloqueo && huespedNuevo) {
        // Primero registrar el huésped nuevo al vuelo
        const nuevoHuesped = await apiFetch<Huesped>('/api/huespedes', {
          method: 'POST',
          body: JSON.stringify(huespedForm)
        });
        actualHuespedId = nuevoHuesped.id;
      }

      await apiFetch('/api/reservas', {
        method: 'POST',
        body: JSON.stringify({
          ...f,
          habitacionId: f.habitacionId || null,
          huespedId: f.esBloqueo ? null : (actualHuespedId || null),
          tarifaId: f.tarifaId || null,
          motivoBloqueo: f.esBloqueo ? f.motivoBloqueo : null,
        })
      });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  const listaEmpleos = [
    'Guardia', 'Cabo', 'Cabo 1º', 'Cabo Mayor', 'Sargento', 'Sargento 1º', 'Brigada',
    'Subteniente', 'Suboficial Mayor', 'Alférez', 'Teniente', 'Capitán', 'Comandante',
    'Tte. Coronel', 'Coronel', 'General B', 'General D', 'Funcionario', 'Otro'
  ];

  const listaSituaciones = [
    'Viogen', 'Asociación', 'Activo', 'Reserva activo', 'Reserva', 'Excedencia',
    'Especiales', 'Retirado', 'Viuda', 'Huerfano'
  ];

  const listaFinalidades = [
    'Otros', 'Comisión', 'Comisión NO indem.', 'Destino', 'Enfermedad', 'Urgencia', 'Sepelio', 'Máx. Estancia'
  ];

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onCancel()}>
      <div className="modal" style={{ maxWidth: 700 }}>
        <div className="modal-header">
          <h3>{isSolicitud ? "➕ Nueva solicitud" : "➕ Nueva reserva / bloqueo"}</h3>
          <button className="modal-close" onClick={onCancel}>×</button>
        </div>
        <form onSubmit={submit}>
          <div className="modal-body">
            {solapamiento && <div className="alert alert-danger">⚠️ {solapamiento}</div>}

            <div className="form-grid">
              <div className="form-group">
                <label>Residencia *</label>
                <select value={residenciaSeleccionadaId} onChange={e => handleResidenciaChange(e.target.value)} required>
                  <option value="">— Seleccione Residencia —</option>
                  {residencias.map(r => (
                    <option key={r.id} value={r.id}>{r.nombre}</option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label>Habitación</label>
                <select value={f.habitacionId} onChange={e => set('habitacionId', e.target.value)}>
                  <option value="">— Asignar automáticamente —</option>
                  {habitaciones.filter(h => !residenciaSeleccionadaId || h.residenciaId === residenciaSeleccionadaId).map(h => (
                    <option key={h.id} value={h.id}>{h.numero} ({h.tipoCodigo})</option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label>Tarifa</label>
                <select value={f.tarifaId} onChange={e => set('tarifaId', e.target.value)}>
                  <option value="">Sin tarifa</option>
                  {tarifasFiltradas.map(t => (
                    <option key={t.id} value={t.id}>{t.nombreTarifa} — {t.precioNoche}€/noche</option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label>Fecha/Hora Solicitud</label>
                <input type="datetime-local" value={f.fechaSolicitud} onChange={e => set('fechaSolicitud', e.target.value)} />
              </div>
              <div className="form-group">
                <label>Fecha entrada *</label>
                <input type="date" value={f.fechaEntrada} onChange={e => set('fechaEntrada', e.target.value)} required />
              </div>
              <div className="form-group">
                <label>Fecha salida *</label>
                <input type="date" value={f.fechaSalida} min={f.fechaEntrada} onChange={e => set('fechaSalida', e.target.value)} required />
              </div>
              <div className="form-group">
                <label>Camas supletorias</label>
                <input type="number" min={0} max={5} value={f.camasSupletorias} onChange={e => set('camasSupletorias', +e.target.value)} />
              </div>
              {!f.esBloqueo && (
                <div className="form-group">
                  <label>Finalidad / Motivo *</label>
                  <select value={f.finalidad} onChange={e => set('finalidad', e.target.value)} required>
                    {listaFinalidades.map(fin => (
                      <option key={fin} value={fin}>{fin}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>

            {/* Tipo de registro */}
            <div className="form-group" style={{ margin: '14px 0 10px' }}>
              <label style={{ fontWeight: 600 }}>Tipo de registro</label>
              <select value={tipoRegistro} onChange={e => {
                const val = e.target.value as 'huesped' | 'reserva-interna' | 'bloqueo';
                setTipoRegistro(val);
                set('esBloqueo', val === 'bloqueo');
                if (val !== 'huesped') {
                  set('huespedId', '');
                }
              }} style={{ padding: '8px 12px', border: '1px solid var(--border)', borderRadius: 7, width: '100%', fontSize: 13 }}>
                <option value="huesped">👤 Reserva para Huésped</option>
                <option value="reserva-interna">📋 Reserva de Habitación / Apartamento (Sin huésped)</option>
                <option value="bloqueo">🔒 Bloqueo de Habitación / Apartamento (Sin huésped)</option>
              </select>
            </div>

            {tipoRegistro !== 'huesped' ? (
              <div className="form-group">
                <label>{tipoRegistro === 'bloqueo' ? 'Motivo del bloqueo *' : 'Motivo de la reserva interna *'}</label>
                <input value={f.motivoBloqueo} onChange={e => set('motivoBloqueo', e.target.value)} placeholder="Mantenimiento, reforma, reservado para evento..." required />
              </div>
            ) : (
              <div style={{ background: 'var(--surface-2)', padding: 14, borderRadius: 8, border: '1px solid var(--border)' }}>
                <div style={{ marginBottom: 10, fontWeight: 600, fontSize: 13, color: 'var(--primary)' }}>👤 Huésped</div>
                
                {!huespedNuevo && !huesped && (
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
                    <div className="search-box" style={{ flex: 1 }}>
                      <span className="search-icon">🔍</span>
                      <input placeholder="Buscar por DNI..." value={buscarDni}
                        onChange={e => setBuscarDni(e.target.value)}
                        onKeyDown={e => e.key === 'Enter' && (e.preventDefault(), buscarHuesped())} />
                    </div>
                    <button type="button" className="btn btn-ghost btn-sm" onClick={buscarHuesped}>Buscar</button>
                  </div>
                )}

                {huesped && (
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 12px', background: 'rgba(45,138,78,0.1)', border: '1px solid rgba(45,138,78,0.2)', borderRadius: 6 }}>
                    <div>
                      <strong>{huesped.nombreCompleto}</strong>
                      <span className="text-muted" style={{ marginLeft: 8 }}>{huesped.dni} · {huesped.empleo || 'Sin Empleo'} ({huesped.situacion || 'Sin Situación'})</span>
                    </div>
                    <button type="button" className="btn btn-ghost btn-sm" onClick={() => { setHuesped(null); set('huespedId', ''); setBuscarDni(''); }}>✕ Buscar otro</button>
                  </div>
                )}

                {huespedNuevo && (
                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                      <span style={{ fontSize: 12, fontWeight: 600, color: 'var(--warning)' }}>⚠️ Huésped no encontrado. Registrando datos:</span>
                      <button type="button" className="btn btn-ghost btn-sm" style={{ padding: '2px 8px' }} onClick={() => { setHuespedNuevo(false); setBuscarDni(''); }}>✕ Cancelar alta al vuelo</button>
                    </div>
                    
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, fontSize: 12 }}>
                      {/* Fila 1: DNI - RANGO */}
                      <div style={{ display: 'flex', gap: 12 }}>
                        <div className="form-group" style={{ flex: 1 }}>
                          <label>DNI *</label>
                          <input value={huespedForm.dni} onChange={e => setH('dni', e.target.value.toUpperCase())} required />
                        </div>
                        <div className="form-group" style={{ flex: 1 }}>
                          <label>Rango *</label>
                          <select value={huespedForm.empleo} onChange={e => setH('empleo', e.target.value)} required>
                            {listaEmpleos.map(emp => (
                              <option key={emp} value={emp}>{emp}</option>
                            ))}
                          </select>
                        </div>
                      </div>

                      {/* Fila 2: NOMBRE - APELLIDOS */}
                      <div style={{ display: 'flex', gap: 12 }}>
                        <div className="form-group" style={{ flex: 1 }}>
                          <label>Nombre *</label>
                          <input value={huespedForm.nombre} onChange={e => setH('nombre', e.target.value)} required />
                        </div>
                        <div className="form-group" style={{ flex: 1 }}>
                          <label>Apellidos *</label>
                          <input value={huespedForm.apellidos} onChange={e => setH('apellidos', e.target.value)} required />
                        </div>
                      </div>

                      {/* Fila 3: TELÉFONO - EMAIL */}
                      <div style={{ display: 'flex', gap: 12 }}>
                        <div className="form-group" style={{ flex: 1 }}>
                          <label>Teléfono</label>
                          <input value={huespedForm.telefono} onChange={e => setH('telefono', e.target.value)} />
                        </div>
                        <div className="form-group" style={{ flex: 1 }}>
                          <label>Email</label>
                          <input type="email" value={huespedForm.email} onChange={e => setH('email', e.target.value)} />
                        </div>
                      </div>

                      {/* Fila 4: DIRECCIÓN POSTAL */}
                      <div className="form-group">
                        <label>Dirección postal *</label>
                        <input value={huespedForm.direccion} onChange={e => setH('direccion', e.target.value)} placeholder="C/. LA CÁMARA 42, 1º A" required />
                      </div>

                      {/* Fila 5: CÓDIGO POSTAL - POBLACIÓN - PROVINCIA */}
                      <div style={{ display: 'flex', gap: 12 }}>
                        <div className="form-group" style={{ flex: 1, position: 'relative' }}>
                          <label>Código Postal *</label>
                          <input value={huespedForm.codigoPostal} onChange={e => handleCpChange(e.target.value)} required />
                          {mostrarSugerencias && sugerenciaOrigen === 'cp' && municipiosSugerencias.length > 0 && (
                            <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: 'var(--surface-1)', border: '1px solid var(--border)', borderRadius: 6, zIndex: 100, maxHeight: 150, overflowY: 'auto', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}>
                              {municipiosSugerencias.map((sug, si) => (
                                <div key={si} style={{ padding: '6px 10px', cursor: 'pointer', borderBottom: '1px solid var(--border-light)', fontSize: 11 }}
                                  onClick={() => seleccionarSugerencia(sug)}
                                  onMouseDown={e => e.preventDefault()}>
                                  <strong>{sug.codigoPostal}</strong> - {sug.municipio} ({sug.provincia})
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                        <div className="form-group" style={{ flex: 1.5, position: 'relative' }}>
                          <label>Población *</label>
                          <input value={huespedForm.municipio} onChange={e => handleMunicipioChange(e.target.value)} required />
                          {mostrarSugerencias && sugerenciaOrigen === 'municipio' && municipiosSugerencias.length > 0 && (
                            <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: 'var(--surface-1)', border: '1px solid var(--border)', borderRadius: 6, zIndex: 100, maxHeight: 150, overflowY: 'auto', boxShadow: '0 4px 12px rgba(0,0,0,0.1)' }}>
                              {municipiosSugerencias.map((sug, si) => (
                                <div key={si} style={{ padding: '6px 10px', cursor: 'pointer', borderBottom: '1px solid var(--border-light)', fontSize: 11 }}
                                  onClick={() => seleccionarSugerencia(sug)}
                                  onMouseDown={e => e.preventDefault()}>
                                  <strong>{sug.codigoPostal}</strong> - {sug.municipio} ({sug.provincia})
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                        <div className="form-group" style={{ flex: 1 }}>
                          <label>Provincia *</label>
                          <input value={huespedForm.provincia} onChange={e => setH('provincia', e.target.value)} required />
                        </div>
                      </div>

                      {/* Fila 6: SITUACIÓN */}
                      <div className="form-group">
                        <label>Situación *</label>
                        <select value={huespedForm.situacion} onChange={e => setH('situacion', e.target.value)} required>
                          {listaSituaciones.map(sit => (
                            <option key={sit} value={sit}>{sit}</option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </div>
                )}

                {!huespedNuevo && !huesped && (
                  <p style={{ fontSize: 12, color: 'var(--text-muted)' }}>Busca un huésped por DNI. Si no existe, se abrirá el alta al vuelo.</p>
                )}

                {/* Campo PAX */}
                <div className="form-group" style={{ marginTop: 14, maxWidth: 180 }}>
                  <label>PAX (Personas alojadas) *</label>
                  <input type="number" min={1} max={10} value={f.numPersonas} onChange={e => set('numPersonas', +e.target.value)} required />
                </div>
              </div>
            )}

            <div className="form-group" style={{ marginTop: 12 }}>
              <label>Observaciones</label>
              <textarea value={f.observaciones} onChange={e => set('observaciones', e.target.value)} rows={2} />
            </div>
          </div>
          <div className="modal-footer">
            <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
            <button type="submit" className="btn btn-primary" disabled={loading || !!solapamiento || (tipoRegistro === 'huesped' && !f.huespedId && !huespedNuevo)}>
              {loading ? 'Guardando...' : '💾 Crear reserva'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// ─── PÁGINA RESERVAS (lista) ───────────────────────────────────────────────────
export function ReservasPage({ setPage }: { setPage: (p: any) => void }) {
  const [reservas, setReservas] = useState<Reserva[]>([]);
  const [loading, setLoading] = useState(true);
  const [buscar, setBuscar] = useState('');
  const [filtroEstado, setFiltroEstado] = useState('');
  const [showNueva, setShowNueva] = useState(false);
  const [statusMsg, setStatusMsg] = useState('');

  // Filtro Residencia
  const [residencias, setResidencias] = useState<any[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>('Todas');

  const [sortField, setSortField] = useState<'numeroOrden' | 'fechaEntrada' | 'fechaSalida'>('numeroOrden');
  const [sortAsc, setSortAsc] = useState(true);

  const cargar = async () => {
    setLoading(true);
    try { setReservas(await apiFetch<Reserva[]>('/api/reservas')); }
    catch { } finally { setLoading(false); }
  };

  const cargarResidencias = async () => {
    try { setResidencias(await apiFetch<any[]>('/api/residencias')); }
    catch {}
  };

  useEffect(() => {
    void cargar();
    void cargarResidencias();
  }, []);

  const reservasFiltradas = reservas.filter(r => {
    // Only show guest reservations (where HuespedId is not null and EsBloqueo is false)
    if (r.esBloqueo || !r.huespedId) return false;

    const b = buscar.toLowerCase();
    const matchBuscar = !buscar || r.huespedNombreCompleto?.toLowerCase().includes(b) ||
      r.huespedDni?.toLowerCase().includes(b) || r.habitacionNumero.includes(b) ||
      r.residenciaNombre.toLowerCase().includes(b);
    const matchEstado = !filtroEstado || r.estado === filtroEstado;
    const matchResidencia = residenciaActiva === 'Todas' || r.residenciaNombre === residenciaActiva;

    return matchBuscar && matchEstado && matchResidencia;
  });

  const handleSort = (field: 'numeroOrden' | 'fechaEntrada' | 'fechaSalida') => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const reservasOrdenadas = [...reservasFiltradas].sort((a, b) => {
    let comparison = 0;
    if (sortField === 'numeroOrden') {
      comparison = a.numeroOrden - b.numeroOrden;
    } else if (sortField === 'fechaEntrada') {
      comparison = a.fechaEntrada.localeCompare(b.fechaEntrada);
    } else if (sortField === 'fechaSalida') {
      comparison = a.fechaSalida.localeCompare(b.fechaSalida);
    }
    return sortAsc ? comparison : -comparison;
  });

  return (
    <>
      {statusMsg && <div className="alert alert-success" onClick={() => setStatusMsg('')} style={{ cursor: 'pointer' }}>{statusMsg} ✕</div>}
      <div className="card">
        <div className="card-header">
          <h3>📋 Reservas activas</h3>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <div className="search-box">
              <span className="search-icon">🔍</span>
              <input placeholder="Buscar huésped, DNI, habitación..." value={buscar} onChange={e => setBuscar(e.target.value)} />
            </div>

            <select value={residenciaActiva} onChange={e => setResidenciaActiva(e.target.value)}
              style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              <option value="Todas">Todas las residencias</option>
              {residencias.map(r => (
                <option key={r.id} value={r.nombre}>{r.nombre}</option>
              ))}
            </select>

            <select value={filtroEstado} onChange={e => setFiltroEstado(e.target.value)}
              style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              <option value="">Todos los estados</option>
              <option value="Pendiente">Pendiente</option>
              <option value="Confirmada">Confirmada</option>
              <option value="CheckIn">Check-in</option>
              <option value="CheckOut">Check-out</option>
            </select>
            <button className="btn btn-primary btn-sm" onClick={() => setShowNueva(true)}>+ Nueva reserva</button>
          </div>
        </div>
        <div className="table-container">
          {loading ? (
            <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div>
          ) : reservasFiltradas.length === 0 ? (
            <div className="empty-state">
              <div className="empty-icon">📋</div>
              <p>No hay reservas {buscar ? 'que coincidan con la búsqueda' : 'activas'}</p>
              <button className="btn btn-primary mt-4" onClick={() => setShowNueva(true)}>+ Crear primera reserva</button>
            </div>
          ) : (
            <table>
              <thead>
                <tr>
                  <th style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('numeroOrden')}>
                    Nº {sortField === 'numeroOrden' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th>Huésped</th>
                  <th>Residencia · Hab.</th>
                  <th style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('fechaEntrada')}>
                    Entrada {sortField === 'fechaEntrada' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('fechaSalida')}>
                    Salida {sortField === 'fechaSalida' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th>Noches</th>
                  <th>Eval.</th>
                  <th>Importe</th>
                  <th>Estado</th>
                  <th>Pago</th>
                </tr>
              </thead>
              <tbody>
                {reservasOrdenadas.map(r => (
                  <tr key={r.id}>
                    <td style={{ cursor: 'pointer' }} onClick={() => {
                      sessionStorage.setItem('gotoReservaId', r.id);
                      sessionStorage.setItem('gotoReservaFecha', r.fechaEntrada);
                      setPage('calendario');
                    }}>
                      <strong style={{ color: 'var(--primary)', textDecoration: 'underline' }}>#{r.numeroOrden}</strong>
                    </td>
                    <td>
                      <strong>{r.huespedNombreCompleto ?? '—'}</strong>{r.huespedDni && <div className="text-muted" style={{ fontSize: 11 }}>{r.huespedDni}</div>}
                    </td>
                    <td>{r.residenciaNombre} · <strong>{r.habitacionNumero}</strong></td>
                    <td>{r.fechaEntrada}</td>
                    <td>{r.fechaSalida}</td>
                    <td style={{ textAlign: 'center' }}>{r.totalNoches}</td>
                    <td><span className="badge badge-primary" style={{ fontSize: 11 }}>{r.evaluacion || '—'}</span></td>
                    <td>{r.importeTotal.toFixed(2)} €</td>
                    <td><span className={`badge estado-${r.estado}`}>{r.estado}</span></td>
                    <td>{r.pagado
                      ? <span className="badge badge-success">✓ {r.formaPago}</span>
                      : <span className="badge badge-danger">Pendiente</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {showNueva && (
        <NuevaReservaModal
          onSaved={() => { setShowNueva(false); setStatusMsg('Reserva creada correctamente.'); void cargar(); }}
          onCancel={() => setShowNueva(false)}
        />
      )}
    </>
  );
}

// ─── PÁGINA BLOQUEOS ───────────────────────────────────────────────────────────
export function BloqueosPage() {
  const [bloqueos, setBloqueos] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [buscar, setBuscar] = useState('');
  const [statusMsg, setStatusMsg] = useState('');
  const [showModal, setShowModal] = useState(false);
  const [editando, setEditando] = useState<any | null>(null);

  // Residencias
  const [residencias, setResidencias] = useState<any[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>('Todas');

  const [sortField, setSortField] = useState<'numeroOrden' | 'fechaEntrada' | 'fechaSalida'>('numeroOrden');
  const [sortAsc, setSortAsc] = useState(true);

  const cargar = async () => {
    setLoading(true);
    try {
      const data = await apiFetch<any[]>('/api/reservas');
      const filtrados = data.filter(r => !r.huespedId);
      setBloqueos(filtrados);
    } catch {} finally { setLoading(false); }
  };

  const cargarResidencias = async () => {
    try {
      const data = await apiFetch<any[]>('/api/residencias');
      setResidencias(data);
    } catch {}
  };

  useEffect(() => {
    void cargar();
    void cargarResidencias();
  }, []);

  const eliminar = async (id: string) => {
    if (!confirm('¿Eliminar/Cancelar este bloqueo/reserva interna?')) return;
    try {
      await apiFetch(`/api/reservas/${id}`, { method: 'DELETE' });
      setStatusMsg('Registro eliminado correctamente.');
      void cargar();
    } catch (e: any) {
      alert(e.message);
    }
  };

  const handleSort = (field: 'numeroOrden' | 'fechaEntrada' | 'fechaSalida') => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const bloqueosFiltrados = bloqueos.filter(r => {
    const b = buscar.toLowerCase();
    const matchBuscar = !buscar || r.habitacionNumero.includes(b) ||
      (r.motivoBloqueo && r.motivoBloqueo.toLowerCase().includes(b)) ||
      r.residenciaNombre.toLowerCase().includes(b);
    const matchResidencia = residenciaActiva === 'Todas' || r.residenciaNombre === residenciaActiva;
    return matchBuscar && matchResidencia;
  });

  const bloqueosOrdenados = [...bloqueosFiltrados].sort((a, b) => {
    let comparison = 0;
    if (sortField === 'numeroOrden') {
      comparison = a.numeroOrden - b.numeroOrden;
    } else if (sortField === 'fechaEntrada') {
      comparison = a.fechaEntrada.localeCompare(b.fechaEntrada);
    } else if (sortField === 'fechaSalida') {
      comparison = a.fechaSalida.localeCompare(b.fechaSalida);
    }
    return sortAsc ? comparison : -comparison;
  });

  return (
    <>
      {statusMsg && <div className="alert alert-success" onClick={() => setStatusMsg('')} style={{ cursor: 'pointer' }}>{statusMsg} ✕</div>}
      <div className="card">
        <div className="card-header">
          <h3>🔒 Bloqueos y Reservas de Habitación</h3>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
            <div className="search-box">
              <span className="search-icon">🔍</span>
              <input placeholder="Buscar por habitación, motivo..." value={buscar} onChange={e => setBuscar(e.target.value)} />
            </div>
            
            <select value={residenciaActiva} onChange={e => setResidenciaActiva(e.target.value)}
              style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              <option value="Todas">Todas las residencias</option>
              {residencias.map(r => (
                <option key={r.id} value={r.nombre}>{r.nombre}</option>
              ))}
            </select>

            <button className="btn btn-primary btn-sm" onClick={() => { setEditando(null); setShowModal(true); }}>+ Nuevo Bloqueo / Reserva Interna</button>
          </div>
        </div>
        <div className="table-container">
          {loading ? (
            <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div>
          ) : bloqueosFiltrados.length === 0 ? (
            <div className="empty-state">
              <div className="empty-icon">🔒</div>
              <p>No hay bloqueos o reservas de habitación configuradas</p>
              <button className="btn btn-primary mt-4" onClick={() => { setEditando(null); setShowModal(true); }}>+ Crear primero</button>
            </div>
          ) : (
            <table>
              <thead>
                <tr>
                  <th>Tipo</th>
                  <th>Residencia</th>
                  <th>Habitación</th>
                  <th style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('fechaEntrada')}>
                    Entrada {sortField === 'fechaEntrada' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('fechaSalida')}>
                    Salida {sortField === 'fechaSalida' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th>Motivo / Causa</th>
                  <th>Acciones</th>
                </tr>
              </thead>
              <tbody>
                {bloqueosOrdenados.map(r => (
                  <tr key={r.id}>
                    <td>
                      {r.esBloqueo
                        ? <span className="badge badge-danger">🔒 Bloqueo</span>
                        : <span className="badge badge-warning">📋 Reserva Interna</span>}
                    </td>
                    <td>{r.residenciaNombre}</td>
                    <td><strong>{r.habitacionNumero}</strong></td>
                    <td>{r.fechaEntrada}</td>
                    <td>{r.fechaSalida}</td>
                    <td>{r.motivoBloqueo || '—'}</td>
                    <td>
                      <div className="actions-cell">
                        <button className="btn btn-ghost btn-sm" onClick={() => { setEditando(r); setShowModal(true); }}>✏️ Editar</button>
                        <button className="btn btn-danger btn-sm" onClick={() => eliminar(r.id)}>🗑️</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {showModal && (
        <BloqueoModal
          item={editando}
          onSaved={() => { setShowModal(false); setStatusMsg('Guardado correctamente.'); void cargar(); }}
          onCancel={() => setShowModal(false)}
        />
      )}
    </>
  );
}

// ─── MODAL BLOQUEO / RESERVA INTERNA ──────────────────────────────────────────
function BloqueoModal({ item, onSaved, onCancel }: { item?: any; onSaved: () => void; onCancel: () => void }) {
  const [loading, setLoading] = useState(false);
  
  const [residencias, setResidencias] = useState<any[]>([]);
  const [selectedResId, setSelectedResId] = useState('');
  const [habitaciones, setHabitaciones] = useState<any[]>([]);
  
  const [habitacionId, setHabitacionId] = useState(item?.habitacionId ?? '');
  const [fechaEntrada, setFechaEntrada] = useState(item?.fechaEntrada ?? '');
  const [fechaSalida, setFechaSalida] = useState(item?.fechaSalida ?? '');
  const [esBloqueo, setEsBloqueo] = useState(item?.esBloqueo ?? true);
  const [motivoBloqueo, setMotivoBloqueo] = useState(item?.motivoBloqueo ?? '');
  const [solapamiento, setSolapamiento] = useState<string | null>(null);

  useEffect(() => {
    const init = async () => {
      try {
        const resList = await apiFetch<any[]>('/api/residencias');
        setResidencias(resList);
        
        const habList = await apiFetch<any[]>('/api/habitaciones');
        setHabitaciones(habList);

        if (item) {
          const matchingHab = habList.find(h => h.id === item.habitacionId);
          if (matchingHab) {
            setSelectedResId(matchingHab.residenciaId);
          }
        } else if (resList.length > 0) {
          setSelectedResId(resList[0].id);
        }
      } catch {}
    };
    void init();
  }, [item]);

  const filteredHabitaciones = habitaciones.filter(h => h.residenciaId === selectedResId && h.activa);

  useEffect(() => {
    if (!item && filteredHabitaciones.length > 0) {
      const exists = filteredHabitaciones.some(h => h.id === habitacionId);
      if (!exists) {
        setHabitacionId(filteredHabitaciones[0].id);
      }
    }
  }, [selectedResId, habitaciones, item]);

  useEffect(() => {
    if (!habitacionId || !fechaEntrada || !fechaSalida) {
      setSolapamiento(null);
      return;
    }
    const t = setTimeout(async () => {
      try {
        const res = await apiFetch<any>(
          `/api/reservas/comprobar-solapamiento?habitacionId=${habitacionId}&fechaEntrada=${fechaEntrada}&fechaSalida=${fechaSalida}&esBloqueo=${esBloqueo}${item ? `&excluirReservaId=${item.id}` : ''}`
        );
        if (res.haySolapamiento) {
          setSolapamiento(res.mensaje);
          alert(res.mensaje);
        } else {
          setSolapamiento(null);
        }
      } catch {
        setSolapamiento(null);
      }
    }, 400);
    return () => clearTimeout(t);
  }, [habitacionId, fechaEntrada, fechaSalida, esBloqueo, item]);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (solapamiento) {
      alert(`No se puede guardar: ${solapamiento}`);
      return;
    }
    setLoading(true);
    try {
      const url = item ? `/api/reservas/${item.id}` : '/api/reservas';
      const method = item ? 'PUT' : 'POST';
      
      const body: any = {
        fechaEntrada,
        fechaSalida,
        numPersonas: 1,
        camasSupletorias: 0,
        esBloqueo,
        motivoBloqueo,
        observaciones: motivoBloqueo,
      };

      if (!item) {
        body.habitacionId = habitacionId;
        body.huespedId = null;
      } else {
        body.estado = item.estado;
        body.pagado = item.pagado;
        body.formaPago = item.formaPago;
      }

      await apiFetch(url, { method, body: JSON.stringify(body) });
      onSaved();
    } catch (err: any) {
      alert(err.message || 'Error al guardar el bloqueo/reserva');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onCancel()}>
      <div className="modal" style={{ maxWidth: 500 }}>
        <div className="modal-header">
          <h3>{item ? '✏️ Editar' : '🔒 Nuevo'} Bloqueo / Reserva Interna</h3>
          <button className="modal-close" onClick={onCancel}>×</button>
        </div>
        <form onSubmit={submit}>
          <div className="modal-body">
            {!item && (
              <>
                <div className="form-group">
                  <label>Residencia</label>
                  <select value={selectedResId} onChange={e => setSelectedResId(e.target.value)} required>
                    {residencias.map(r => (
                      <option key={r.id} value={r.id}>{r.nombre}</option>
                    ))}
                  </select>
                </div>
                <div className="form-group" style={{ marginTop: 12 }}>
                  <label>Habitación / Apartamento</label>
                  <select value={habitacionId} onChange={e => setHabitacionId(e.target.value)} required>
                    {filteredHabitaciones.map(h => (
                      <option key={h.id} value={h.id}>{h.numero} ({h.tipoNombre})</option>
                    ))}
                  </select>
                </div>
              </>
            )}

            <div className="form-group" style={{ marginTop: 12 }}>
              <label>Tipo de Registro</label>
              <select value={esBloqueo ? 'true' : 'false'} onChange={e => setEsBloqueo(e.target.value === 'true')} required>
                <option value="true">Bloqueo Completo (Lógica inclusiva)</option>
                <option value="false">Reserva de Habitación (Lógica exclusiva de checkin-checkout)</option>
              </select>
            </div>

            <div style={{ display: 'flex', gap: 12, marginTop: 12 }}>
              <div className="form-group" style={{ flex: 1 }}>
                <label>Fecha Entrada</label>
                <input type="date" value={fechaEntrada} onChange={e => setFechaEntrada(e.target.value)} required />
              </div>
              <div className="form-group" style={{ flex: 1 }}>
                <label>Fecha Salida</label>
                <input type="date" value={fechaSalida} min={fechaEntrada} onChange={e => setFechaSalida(e.target.value)} required />
              </div>
            </div>

            <div className="form-group" style={{ marginTop: 12 }}>
              <label>Motivo / Causa</label>
              <input value={motivoBloqueo} onChange={e => setMotivoBloqueo(e.target.value)} placeholder="Ej. Pintura, Fuga de agua, Limpieza..." required />
            </div>

            {solapamiento && (
              <div className="alert alert-danger" style={{ marginTop: 14 }}>
                ⚠️ <strong>Conflicto detectado:</strong> {solapamiento}
              </div>
            )}
          </div>
          <div className="modal-footer">
            <button type="button" className="btn btn-ghost" onClick={onCancel} disabled={loading}>Cancelar</button>
            <button type="submit" className="btn btn-primary" disabled={loading || !!solapamiento}>
              {loading ? 'Guardando...' : 'Guardar'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// ─── PÁGINA SOLICITUDES ───────────────────────────────────────────────────────
function EditarSolicitudModal({
  reserva, onSaved, onCancel
}: {
  reserva: any; onSaved: () => void; onCancel: () => void;
}) {
  const [loading, setLoading] = useState(false);
  const [fechaEntrada, setFechaEntrada] = useState(reserva.fechaEntrada);
  const [fechaSalida, setFechaSalida] = useState(reserva.fechaSalida);
  const [numPersonas, setNumPersonas] = useState(reserva.numPersonas);
  const [finalidad, setFinalidad] = useState(reserva.finalidad || 'Otros');
  const [resolucion, setResolucion] = useState(reserva.resolucion || 'SI');

  const formatDateTimeLocal = (isoStr: string) => {
    if (!isoStr) return '';
    const d = new Date(isoStr);
    if (isNaN(d.getTime())) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const h = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    return `${y}-${m}-${day}T${h}:${min}`;
  };

  const [fechaSolicitud, setFechaSolicitud] = useState(() => formatDateTimeLocal(reserva.fechaSolicitud));

  // Huésped
  const [nombre, setNombre] = useState(reserva.huespedNombre || '');
  const [apellidos, setApellidos] = useState(reserva.huespedApellidos || '');
  const [dni, setDni] = useState(reserva.huespedDni || '');
  const [telefono, setTelefono] = useState(reserva.huespedTelefono || '');

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    try {
      if (reserva.huespedId) {
        // Fetch existing guest first to preserve its fields
        const currentGuest = await apiFetch<any>(`/api/huespedes/${reserva.huespedId}`);
        await apiFetch(`/api/huespedes/${reserva.huespedId}`, {
          method: 'PUT',
          body: JSON.stringify({
            ...currentGuest,
            nombre: nombre.trim(),
            apellidos: apellidos.trim(),
            dni: dni.trim().toUpperCase(),
            telefono: telefono.trim(),
          })
        });
      }
      
      // Update reservation
      const currentRes = await apiFetch<any>(`/api/reservas/${reserva.id}`);
      await apiFetch(`/api/reservas/${reserva.id}`, {
        method: 'PUT',
        body: JSON.stringify({
          fechaEntrada,
          fechaSalida,
          numPersonas: +numPersonas,
          camasSupletorias: currentRes.camasSupletorias,
          estado: currentRes.estado,
          esBloqueo: currentRes.esBloqueo,
          motivoBloqueo: currentRes.motivoBloqueo,
          tarifaId: currentRes.tarifaId,
          observaciones: currentRes.observaciones,
          pagado: currentRes.pagado,
          formaPago: currentRes.formaPago,
          finalidad,
          resolucion,
          fechaSolicitud: fechaSolicitud || null,
        })
      });

      onSaved();
    } catch (err: any) {
      alert(err.message || 'Error al actualizar la solicitud');
    } finally {
      setLoading(false);
    }
  };

  const listaFinalidades = [
    'Otros', 'Comisión', 'Comisión NO indem.', 'Destino', 'Enfermedad', 'Urgencia', 'Sepelio', 'Máx. Estancia'
  ];

  const listaResoluciones = [
    'SI', 'CONCEDIDA', 'REEVALUADA', 'NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'
  ];

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onCancel()}>
      <div className="modal" style={{ maxWidth: 650 }}>
        <div className="modal-header">
          <h3>✏️ Editar Solicitud #{reserva.numeroOrden}</h3>
          <button className="modal-close" onClick={onCancel}>×</button>
        </div>
        <form onSubmit={submit}>
          <div className="modal-body" style={{ maxHeight: '70vh', overflowY: 'auto' }}>
            <div style={{ fontWeight: 600, marginBottom: 12, color: 'var(--primary)', borderBottom: '1px solid var(--border)', paddingBottom: 6 }}>📅 Datos de la solicitud</div>
            <div className="form-grid">
              <div className="form-group">
                <label>Fecha/Hora Solicitud</label>
                <input type="datetime-local" value={fechaSolicitud} onChange={e => setFechaSolicitud(e.target.value)} />
              </div>
              <div className="form-group">
                <label>Fecha Entrada</label>
                <input type="date" value={fechaEntrada} onChange={e => setFechaEntrada(e.target.value)} required />
              </div>
              <div className="form-group">
                <label>Fecha Salida</label>
                <input type="date" value={fechaSalida} min={fechaEntrada} onChange={e => setFechaSalida(e.target.value)} required />
              </div>
              <div className="form-group">
                <label>PAX (Personas)</label>
                <input type="number" min={1} max={10} value={numPersonas} onChange={e => setNumPersonas(+e.target.value)} required />
              </div>
              <div className="form-group">
                <label>Finalidad / Motivo</label>
                <select value={finalidad} onChange={e => setFinalidad(e.target.value)} required>
                  {listaFinalidades.map(fin => <option key={fin} value={fin}>{fin}</option>)}
                </select>
              </div>
              <div className="form-group" style={{ gridColumn: 'span 2' }}>
                <label>Resolución</label>
                <select value={resolucion} onChange={e => setResolucion(e.target.value)} required>
                  {listaResoluciones.map(res => <option key={res} value={res}>{res}</option>)}
                </select>
              </div>
            </div>

            {reserva.huespedId && (
              <>
                <div style={{ fontWeight: 600, marginTop: 20, marginBottom: 12, color: 'var(--primary)', borderBottom: '1px solid var(--border)', paddingBottom: 6 }}>👤 Datos del huésped</div>
                <div className="form-grid">
                  <div className="form-group">
                    <label>Nombre</label>
                    <input type="text" value={nombre} onChange={e => setNombre(e.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>Apellidos</label>
                    <input type="text" value={apellidos} onChange={e => setApellidos(e.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>DNI</label>
                    <input type="text" value={dni} onChange={e => setDni(e.target.value.toUpperCase())} required />
                  </div>
                  <div className="form-group">
                    <label>Teléfono</label>
                    <input type="text" value={telefono} onChange={e => setTelefono(e.target.value)} />
                  </div>
                </div>
              </>
            )}
          </div>
          <div className="modal-footer">
            <button type="button" className="btn btn-ghost" onClick={onCancel} disabled={loading}>Cancelar</button>
            <button type="submit" className="btn btn-primary" disabled={loading}>
              {loading ? 'Guardando...' : '💾 Guardar cambios'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

export function HuespedesPage() {
  const [solicitudes, setSolicitudes] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [buscar, setBuscar] = useState('');
  const [statusMsg, setStatusMsg] = useState('');
  const [showNueva, setShowNueva] = useState(false);
  const [editando, setEditando] = useState<any | null>(null);

  // Sorting
  const [sortField, setSortField] = useState<'fechaEntrada' | 'fechaSalida' | 'numPersonas' | 'finalidad' | 'resolucion' | 'numeroOrden'>('numeroOrden');
  const [sortAsc, setSortAsc] = useState(true);

  // Pagination
  const [pagina, setPagina] = useState(1);
  const filasPorPagina = 25;

  const cargar = async () => {
    setLoading(true);
    try {
      const data = await apiFetch<any[]>('/api/reservas?incluirCanceladas=true');
      const filtradas = data.filter(r => !r.esBloqueo && r.huespedId);
      setSolicitudes(filtradas);
    } catch { } finally { setLoading(false); }
  };

  useEffect(() => {
    void cargar();
  }, []);

  const eliminar = async (id: string) => {
    if (!confirm('¿Eliminar/Cancelar esta solicitud?')) return;
    try {
      await apiFetch(`/api/reservas/${id}`, { method: 'DELETE' });
      setStatusMsg('Solicitud eliminada/cancelada correctamente.');
      void cargar();
    } catch (e: any) {
      alert(e.message);
    }
  };

  // Filtered by search
  const solicitudesFiltradas = solicitudes.filter(r => {
    if (!buscar.trim()) return true;
    const query = buscar.toLowerCase();
    return (
      r.numeroOrden.toString().includes(query) ||
      (r.huespedNombreCompleto && r.huespedNombreCompleto.toLowerCase().includes(query)) ||
      (r.huespedNombre && r.huespedNombre.toLowerCase().includes(query)) ||
      (r.huespedApellidos && r.huespedApellidos.toLowerCase().includes(query)) ||
      (r.huespedDni && r.huespedDni.toLowerCase().includes(query)) ||
      (r.huespedEmail && r.huespedEmail.toLowerCase().includes(query)) ||
      (r.huespedTelefono && r.huespedTelefono.toLowerCase().includes(query))
    );
  });

  // Sort
  const handleSort = (field: typeof sortField) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
    setPagina(1);
  };

  const sortedSolicitudes = [...solicitudesFiltradas].sort((a, b) => {
    let comparison = 0;
    if (sortField === 'numeroOrden') {
      comparison = a.numeroOrden - b.numeroOrden;
    } else if (sortField === 'numPersonas') {
      comparison = a.numPersonas - b.numPersonas;
    } else if (sortField === 'fechaEntrada') {
      comparison = a.fechaEntrada.localeCompare(b.fechaEntrada);
    } else if (sortField === 'fechaSalida') {
      comparison = a.fechaSalida.localeCompare(b.fechaSalida);
    } else if (sortField === 'finalidad') {
      comparison = (a.finalidad || '').localeCompare(b.finalidad || '');
    } else if (sortField === 'resolucion') {
      comparison = (a.resolucion || '').localeCompare(b.resolucion || '');
    }
    return sortAsc ? comparison : -comparison;
  });

  // Paginated
  const totalPaginas = Math.ceil(sortedSolicitudes.length / filasPorPagina) || 1;
  const paginatedSolicitudes = sortedSolicitudes.slice((pagina - 1) * filasPorPagina, pagina * filasPorPagina);

  const getResolucionBadgeClass = (res: string) => {
    const r = res?.toUpperCase() || 'SI';
    if (r === 'SI' || r === 'CONCEDIDA') return 'badge-success';
    if (r === 'REEVALUADA') return 'badge-warning';
    return 'badge-danger';
  };

  const formatFecha = (isoStr: string) => {
    if (!isoStr) return '—';
    const d = new Date(isoStr);
    if (isNaN(d.getTime())) return '—';
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${minutes}`;
  };

  return (
    <>
      {statusMsg && <div className="alert alert-success" onClick={() => setStatusMsg('')} style={{ cursor: 'pointer' }}>{statusMsg} ✕</div>}
      <div className="card">
        <div className="card-header">
          <h3>👥 Solicitudes registradas</h3>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <div className="search-box">
              <span className="search-icon">🔍</span>
              <input placeholder="Buscar por Nº orden, nombre, DNI, teléfono..." value={buscar} onChange={e => { setBuscar(e.target.value); setPagina(1); }} />
            </div>
            <button className="btn btn-primary btn-sm" onClick={() => setShowNueva(true)}>+ Nueva solicitud</button>
          </div>
        </div>
        <div className="table-container" style={{ overflowX: 'auto' }}>
          {loading ? (
            <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div>
          ) : solicitudesFiltradas.length === 0 ? (
            <div className="empty-state">
              <div className="empty-icon">👥</div>
              <p>No hay solicitudes registradas</p>
              <button className="btn btn-primary mt-4" onClick={() => setShowNueva(true)}>+ Registrar primera</button>
            </div>
          ) : (
            <table style={{ fontSize: '11.5px', width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ background: 'var(--primary-dark)', color: 'white' }}>
                  <th style={{ cursor: 'pointer', userSelect: 'none', padding: '6px 8px' }} onClick={() => handleSort('numeroOrden')}>
                    Nº<br/>ORDEN {sortField === 'numeroOrden' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ padding: '6px 8px' }}>
                    FECHA/HORA<br/>SOLICITUD
                  </th>
                  <th style={{ padding: '6px 8px' }}>NOMBRE</th>
                  <th style={{ padding: '6px 8px' }}>APELLIDOS</th>
                  <th style={{ padding: '6px 8px' }}>DNI</th>
                  <th style={{ padding: '6px 8px' }}>TELÉFONO</th>
                  <th style={{ cursor: 'pointer', userSelect: 'none', padding: '6px 8px' }} onClick={() => handleSort('fechaEntrada')}>
                    ENTRADA {sortField === 'fechaEntrada' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ cursor: 'pointer', userSelect: 'none', padding: '6px 8px' }} onClick={() => handleSort('fechaSalida')}>
                    SALIDA {sortField === 'fechaSalida' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ cursor: 'pointer', userSelect: 'none', padding: '6px 8px', textAlign: 'center' }} onClick={() => handleSort('numPersonas')}>
                    Nº PAX {sortField === 'numPersonas' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ cursor: 'pointer', userSelect: 'none', padding: '6px 8px' }} onClick={() => handleSort('finalidad')}>
                    FINALIDAD {sortField === 'finalidad' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ cursor: 'pointer', userSelect: 'none', padding: '6px 8px' }} onClick={() => handleSort('resolucion')}>
                    RESOLUCIÓN {sortField === 'resolucion' ? (sortAsc ? '▲' : '▼') : ''}
                  </th>
                  <th style={{ padding: '6px 8px', width: 120 }}>ACCIONES</th>
                </tr>
              </thead>
              <tbody>
                {paginatedSolicitudes.map(r => (
                  <tr key={r.id} style={{ borderBottom: '1px solid var(--border-light)' }}>
                    <td style={{ padding: '5px 8px' }}><strong>#{r.numeroOrden}</strong></td>
                    <td style={{ padding: '5px 8px' }}>{formatFecha(r.fechaSolicitud || r.creadoEn)}</td>
                    <td style={{ padding: '5px 8px' }}>{r.huespedNombre}</td>
                    <td style={{ padding: '5px 8px' }}>{r.huespedApellidos}</td>
                    <td style={{ padding: '5px 8px' }}><code>{r.huespedDni}</code></td>
                    <td style={{ padding: '5px 8px' }}>{r.huespedTelefono || '—'}</td>
                    <td style={{ padding: '5px 8px' }}>{r.fechaEntrada}</td>
                    <td style={{ padding: '5px 8px' }}>{r.fechaSalida}</td>
                    <td style={{ padding: '5px 8px', textAlign: 'center' }}><strong>{r.numPersonas}</strong></td>
                    <td style={{ padding: '5px 8px' }}>{r.finalidad || 'Otros'}</td>
                    <td style={{ padding: '5px 8px' }}>
                      <span className={`badge ${getResolucionBadgeClass(r.resolucion)}`} style={{ fontSize: '10px', padding: '2px 6px' }}>
                        {r.resolucion || 'SI'}
                      </span>
                    </td>
                    <td style={{ padding: '5px 8px' }}>
                      <div className="actions-cell" style={{ display: 'flex', gap: 4 }}>
                        <button className="btn btn-ghost btn-sm" style={{ padding: '2px 6px', fontSize: 11 }} onClick={() => setEditando(r)}>✏️ Editar</button>
                        <button className="btn btn-danger btn-sm" style={{ padding: '2px 6px', fontSize: 11 }} onClick={() => eliminar(r.id)}>🗑️</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        {/* Paginación */}
        {totalPaginas > 1 && (
          <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: 12, padding: 14, borderTop: '1px solid var(--border-light)' }}>
            <button className="btn btn-ghost btn-sm" disabled={pagina === 1} onClick={() => setPagina(p => Math.max(p - 1, 1))}>◀</button>
            <span style={{ fontSize: 12, fontWeight: 600 }}>Página {pagina} de {totalPaginas} ({solicitudesFiltradas.length} resultados)</span>
            <button className="btn btn-ghost btn-sm" disabled={pagina === totalPaginas} onClick={() => setPagina(p => Math.min(p + 1, totalPaginas))}>▶</button>
          </div>
        )}
      </div>

      {showNueva && (
        <NuevaReservaModal
          isSolicitud={true}
          onSaved={() => { setShowNueva(false); setStatusMsg('Solicitud creada correctamente.'); void cargar(); }}
          onCancel={() => setShowNueva(false)}
        />
      )}

      {editando && (
        <EditarSolicitudModal
          reserva={editando}
          onSaved={() => { setEditando(null); setStatusMsg('Solicitud actualizada correctamente.'); void cargar(); }}
          onCancel={() => setEditando(null)}
        />
      )}
    </>
  );
}
