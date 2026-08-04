import { useState, useEffect, useRef } from 'react';
import { useCalendario } from '../../hooks/useCalendario';
import '../../calendar.css';
import { formatFechaCorta, addDays, toDateStr, parseLocal, diasEntre } from '../../utils/dateUtils';
import { MESES, DIAS_LONG, ESTADO_COLOR } from '../../utils/constants';
import { Reserva } from '../../types';
import { ReservaDetalleModal } from '../reservas/ReservaDetalleModal';

// ─── CALENDARIO GANTT ─────────────────────────────────────────────────────────
export function CalendarioGantt({ onNuevaReserva }: { onNuevaReserva: (habitacionId: string, fecha: string) => void }) {
  const hoy = new Date();
  const {
    inicio, setInicio,
    dias, setDias,
    residenciaFiltro, setResidenciaFiltro,
    calData,
    festivos,
    residencias,
    loading,
    tooltip, setTooltip,
    reservaSeleccionada, setReservaSeleccionada,
    highlightReservaId, setHighlightReservaId,
    draggedReserva, setDraggedReserva,
    dragOverTarget, setDragOverTarget,
    moverStatus, setMoverStatus,
    tooltipTimer,
    cargarCalData,
    handleDropReserva,
  } = useCalendario();

  const fin = toDateStr(addDays(parseLocal(inicio), dias));
  const fechas = diasEntre(inicio, fin);
  const hoyStr = toDateStr(hoy);

  const pendingSelectIdRef = useRef<string | null>(sessionStorage.getItem('gotoReservaId'));

  // Reordenación de Residencias en Calendario
  const [showOrdenModal, setShowOrdenModal] = useState(false);
  const [ordenResidencias, setOrdenResidencias] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('pref_orden_residencias_calendario');
      if (saved) return JSON.parse(saved);
    } catch {}
    return [];
  });

  const listaOrdenadaResidencias = [...residencias].sort((a, b) => {
    const idxA = ordenResidencias.indexOf(a.id);
    const idxB = ordenResidencias.indexOf(b.id);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.orden - b.orden;
  });

  const moverResidencia = (index: number, direction: 'up' | 'down') => {
    const list = [...listaOrdenadaResidencias];
    const targetIdx = direction === 'up' ? index - 1 : index + 1;
    if (targetIdx < 0 || targetIdx >= list.length) return;
    const temp = list[index];
    list[index] = list[targetIdx];
    list[targetIdx] = temp;

    const newOrderIds = list.map(r => r.id);
    setOrdenResidencias(newOrderIds);
    localStorage.setItem('pref_orden_residencias_calendario', JSON.stringify(newOrderIds));
  };

  const restablecerOrdenDefault = () => {
    setOrdenResidencias([]);
    localStorage.removeItem('pref_orden_residencias_calendario');
  };

  const getResidenciaRank = (resId?: string, resName?: string) => {
    if (!resId && !resName) return 999;
    const idxRes = listaOrdenadaResidencias.findIndex(r => r.id === resId || r.nombre === resName);
    return idxRes !== -1 ? idxRes : 999;
  };

  const habitacionesOrdenadas = [...(calData?.habitaciones ?? [])].sort((a, b) => {
    const rankA = getResidenciaRank(a.residenciaId, a.residenciaNombre);
    const rankB = getResidenciaRank(b.residenciaId, b.residenciaNombre);
    if (rankA !== rankB) return rankA - rankB;
    if (a.orden !== b.orden) return a.orden - b.orden;
    return (a.numero || '').localeCompare(b.numero || '', undefined, { numeric: true });
  });

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
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <select className="form-group select" value={residenciaFiltro} onChange={e => { const val = e.target.value; setResidenciaFiltro(val); localStorage.setItem('pref_residencia_calendario', val); }}
            style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
            <option value="">Todas las residencias</option>
            {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
          </select>
          {residencias.length > 1 && (
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              title="Configurar orden de residencias en el calendario"
              onClick={() => setShowOrdenModal(true)}
              style={{
                padding: '6px 10px',
                border: '1px solid var(--border)',
                borderRadius: 7,
                fontSize: 13,
                fontWeight: 600,
                background: 'var(--surface-1)',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: 5
              }}
            >
              ⚙️ <span style={{ fontSize: 12 }}>Orden</span>
            </button>
          )}
        </div>
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
          {habitacionesOrdenadas.map((hab, rowIdx) => {
            const row = rowIdx + 3;
            const prevHab = rowIdx > 0 ? habitacionesOrdenadas[rowIdx - 1] : null;
            const isFirstOfResidencia = !prevHab || prevHab.residenciaNombre !== hab.residenciaNombre;
            const isNewResidenciaBlock = isFirstOfResidencia && !residenciaFiltro && rowIdx > 0;
            const separatorStyle: React.CSSProperties = isNewResidenciaBlock ? { borderTop: '3px solid var(--primary-dark, #1e3a8a)' } : {};

            return (
              <div key={hab.habitacionId} className="cal-row">
                {/* Etiqueta habitación */}
                <div id={`room-label-${hab.habitacionId}`} className="cal-room-label" style={{
                  gridRow: row, gridColumn: 1, position: 'sticky', left: 0, zIndex: 5,
                  ...separatorStyle
                }}>
                  <div className="room-num">🔑 {hab.numero}</div>
                  <div className="room-res">{hab.residenciaNombre}</div>
                  <div className="room-tipo"><span className="badge badge-primary" style={{ fontSize: 9, padding: '1px 5px' }}>{hab.tipoCodigo}</span></div>
                </div>

                {/* Celdas vacías clicables y receptoras de Drag & Drop */}
                {fechas.map((d, ci) => {
                  const ds = toDateStr(d);
                  const dow = d.getDay();
                  const isWeekend = dow === 0 || dow === 6;
                  const isFestivo = festivos.has(ds);
                  const isToday = ds === hoyStr;
                  const isDragOver = dragOverTarget?.habitacionId === hab.habitacionId && dragOverTarget?.fecha === ds;
                  return (
                    <div key={ds} className={`cal-cell ${isWeekend ? 'weekend' : ''} ${isFestivo ? 'festivo' : ''} ${isToday ? 'today' : ''} ${isDragOver ? 'drag-over' : ''}`}
                      style={{ gridRow: row, gridColumn: ci + 2, ...separatorStyle }}
                      onClick={() => onNuevaReserva(hab.habitacionId, ds)}
                      onDragOver={e => {
                        if (draggedReserva) {
                          e.preventDefault();
                          e.dataTransfer.dropEffect = 'move';
                        }
                      }}
                      onDragEnter={() => {
                        if (draggedReserva) {
                          setDragOverTarget({ habitacionId: hab.habitacionId, fecha: ds });
                        }
                      }}
                      onDragLeave={() => {
                        if (dragOverTarget?.habitacionId === hab.habitacionId && dragOverTarget?.fecha === ds) {
                          setDragOverTarget(null);
                        }
                      }}
                      onDrop={e => {
                        e.preventDefault();
                        if (draggedReserva) {
                          void handleDropReserva(draggedReserva, hab.habitacionId, ds);
                        }
                      }}
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

                  if (resEnd <= calStart || resStart >= calEnd) return null;

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
                        draggable={!res.esBloqueo}
                        onDragStart={e => {
                          if (res.esBloqueo) return;
                          e.dataTransfer.setData('text/plain', res.id);
                          e.dataTransfer.effectAllowed = 'move';
                          setDraggedReserva(res);
                          setTooltip(null);
                        }}
                        onDragEnd={() => {
                          setDraggedReserva(null);
                          setDragOverTarget(null);
                        }}
                        className={`cal-reserva ${ESTADO_COLOR[res.estado] || 'confirmada'} ${draggedReserva?.id === res.id ? 'dragging' : ''}`}
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
                          border: res.id === highlightReservaId ? '2.5px solid var(--accent)' : undefined,
                          boxShadow: res.id === highlightReservaId ? '0 0 15px var(--accent-light)' : undefined,
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
          <div className="cal-legend-item"><div className="cal-legend-dot" style={{ background: 'var(--cal-cell-festivo)', border: '1px solid var(--cal-festivo-border)' }}></div>Festivo</div>
        </div>
      </div>

      {/* Tooltip flotante */}
      {tooltip && (
        <div className="reserva-tooltip"
          style={{ position: 'fixed', top: tooltip.y, left: Math.min(tooltip.x, window.innerWidth - 290) }}>
          <h4>{tooltip.reserva.esBloqueo ? '🔒 Bloqueo' : `👤 ${tooltip.reserva.huespedNombreCompleto}`}</h4>
          <p>🏠 {tooltip.reserva.habitacionNumero} — {tooltip.reserva.residenciaNombre}</p>
          <p>📅 {formatFechaCorta(tooltip.reserva.fechaEntrada)} → {formatFechaCorta(tooltip.reserva.fechaSalida)} ({tooltip.reserva.totalNoches}n)</p>
          {!tooltip.reserva.esBloqueo && <p>💰 {tooltip.reserva.importeTotal.toFixed(2)} € {tooltip.reserva.pagado ? '✓ Pagado' : '⏳ Pendiente'}</p>}
          <p style={{ marginTop: 4, opacity: 0.6, fontSize: 10 }}>Clic para ver detalle</p>
        </div>
      )}

      {/* Toast de notificación emergente elegante */}
      {moverStatus && (
        <div className="cal-toast-container">
          <div className={`cal-toast ${moverStatus.isError ? 'error' : 'success'}`}>
            <div className="cal-toast-icon">
              {moverStatus.isError ? '⚠️' : '✓'}
            </div>
            <div className="cal-toast-content">
              <div className="cal-toast-title">{moverStatus.title}</div>
              <div className="cal-toast-body">{moverStatus.details}</div>
            </div>
            <button className="cal-toast-close" onClick={() => setMoverStatus(null)} title="Cerrar">✕</button>
          </div>
        </div>
      )}

      {/* Modal detalle reserva */}
      {reservaSeleccionada && (
        <ReservaDetalleModal reserva={reservaSeleccionada} onClose={() => { setReservaSeleccionada(null); void cargarCalData(); }} />
      )}

      {/* Modal Ordenación Residencias */}
      {showOrdenModal && (
        <div className="modal-overlay" style={{
          position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
          background: 'rgba(0,0,0,0.6)', backdropFilter: 'blur(3px)',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          zIndex: 9999, padding: 16
        }}>
          <div className="modal-card" style={{
            background: 'var(--surface-0, #ffffff)', color: 'var(--text-main, #333)',
            borderRadius: 12, width: '100%', maxWidth: 480,
            boxShadow: '0 20px 25px -5px rgba(0,0,0,0.3), 0 10px 10px -5px rgba(0,0,0,0.2)',
            overflow: 'hidden'
          }}>
            <div style={{
              background: 'var(--primary-dark, #1e3a8a)', color: 'white',
              padding: '16px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center'
            }}>
              <h4 style={{ margin: 0, fontSize: 16, display: 'flex', alignItems: 'center', gap: 8 }}>
                ⚙️ Orden de Residencias en Calendario
              </h4>
              <button onClick={() => setShowOrdenModal(false)} style={{ background: 'none', border: 'none', color: 'white', fontSize: 18, cursor: 'pointer' }}>✕</button>
            </div>

            <div style={{ padding: 20 }}>
              <p style={{ fontSize: 12.5, color: 'var(--text-muted)', marginTop: 0, marginBottom: 16 }}>
                Utiliza las flechas para determinar la posición vertical de las residencias de arriba a abajo cuando se muestra la vista <strong>"Todas las residencias"</strong>.
              </p>

              <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                {listaOrdenadaResidencias.map((res, index) => (
                  <div key={res.id} style={{
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                    padding: '10px 14px', background: 'var(--surface-1, #f8fafc)',
                    border: '1px solid var(--border, #e2e8f0)', borderRadius: 8
                  }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                      <span className="badge badge-primary" style={{ fontSize: 11, padding: '3px 8px', minWidth: 28, textAlign: 'center', fontWeight: 700 }}>
                        #{index + 1}
                      </span>
                      <strong style={{ fontSize: 13.5 }}>🏛️ {res.nombre}</strong>
                    </div>
                    <div style={{ display: 'flex', gap: 6 }}>
                      <button
                        type="button"
                        className="btn btn-ghost btn-sm"
                        disabled={index === 0}
                        onClick={() => moverResidencia(index, 'up')}
                        style={{ padding: '4px 8px', fontSize: 12 }}
                      >
                        ▲ Subir
                      </button>
                      <button
                        type="button"
                        className="btn btn-ghost btn-sm"
                        disabled={index === listaOrdenadaResidencias.length - 1}
                        onClick={() => moverResidencia(index, 'down')}
                        style={{ padding: '4px 8px', fontSize: 12 }}
                      >
                        ▼ Bajar
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div style={{
              padding: '12px 20px', background: 'var(--surface-1, #f1f5f9)',
              borderTop: '1px solid var(--border, #e2e8f0)',
              display: 'flex', justifyContent: 'space-between', alignItems: 'center'
            }}>
              <button type="button" className="btn btn-ghost btn-sm" onClick={restablecerOrdenDefault}>
                ↺ Orden por defecto
              </button>
              <button type="button" className="btn btn-primary btn-sm" onClick={() => setShowOrdenModal(false)}>
                ✓ Aceptar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
