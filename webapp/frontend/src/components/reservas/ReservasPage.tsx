import { useState } from 'react';
import { formatFechaDisplay } from '../../api';
import { useColumnConfig, ColumnDef } from '../../useColumnConfig';
import { NuevaReservaModal } from './NuevaReservaModal';
import { useReservas } from '../../hooks/useReservas';

// --- DEFINICION DE COLUMNAS RESERVAS ACTIVAS ---
const RESERVAS_COLUMNS_DEF: ColumnDef[] = [
  { key: 'numeroOrden', label: 'Nº', defaultWidth: 44, minWidth: 35 },
  { key: 'huesped', label: 'Huésped', defaultWidth: 210, minWidth: 100 },
  { key: 'residenciaHab', label: 'Residencia · Hab.', defaultWidth: 160, minWidth: 100 },
  { key: 'fechaEntrada', label: 'Entrada', defaultWidth: 95, minWidth: 70 },
  { key: 'fechaSalida', label: 'Salida', defaultWidth: 95, minWidth: 70 },
  { key: 'totalNoches', label: 'Noches', defaultWidth: 55, minWidth: 40 },
  { key: 'evaluacion', label: 'Eval.', defaultWidth: 70, minWidth: 45 },
  { key: 'importeTotal', label: 'Importe', defaultWidth: 80, minWidth: 55 },
  { key: 'estado', label: 'Estado', defaultWidth: 90, minWidth: 60 },
  { key: 'pagado', label: 'Pago', defaultWidth: 90, minWidth: 60 },
  { key: 'acciones', label: 'Acciones', defaultWidth: 65, minWidth: 50, isFixed: true },
];

export function ReservasPage({ setPage }: { setPage: (p: any) => void }) {
  const {
    loading,
    buscar, setBuscar,
    filtroEstado, setFiltroEstado,
    statusMsg, setStatusMsg,
    residencias,
    residenciaActiva, setResidenciaActiva,
    sortField, sortAsc, handleSort,
    pagina, setPagina,
    totalPaginas,
    paginatedReservas,
    reservasFiltradas,
    cargar,
  } = useReservas();

  const [showNueva, setShowNueva] = useState(false);
  const [editandoReserva, setEditandoReserva] = useState<any | null>(null);

  // Column Manager
  const colConfig = useColumnConfig('pref_cols_reservas_activas', RESERVAS_COLUMNS_DEF);

  const formatNombreTrunc = (nombre?: string) => {
    if (!nombre) return '—';
    return nombre.length > 30 ? nombre.substring(0, 30) + '…' : nombre;
  };

  // Ancho dinámico del título según las dos primeras columnas para alineamiento perfecto
  const primerAncho = colConfig.columns.find(c => c.key === 'numeroOrden')?.visible ? (colConfig.columns.find(c => c.key === 'numeroOrden')?.width ?? 44) : 0;
  const segundoAncho = colConfig.columns.find(c => c.key === 'huesped')?.visible ? (colConfig.columns.find(c => c.key === 'huesped')?.width ?? 210) : 0;
  const tituloWidth = primerAncho + segundoAncho;

  return (
    <>
      {statusMsg && <div className="alert alert-success" onClick={() => setStatusMsg('')} style={{ cursor: 'pointer' }}>{statusMsg} ✕</div>}
      <div className="card">
        <div className="card-header" style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-start', gap: 0 }}>
          <h3 style={{ width: `${tituloWidth}px`, minWidth: `${tituloWidth}px`, margin: 0, paddingRight: 10, flexShrink: 0, transition: 'width 0.1s' }}>📋 Reservas activas</h3>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap', flex: 1 }}>
            {/* 1. Selector de Residencia a la izquierda */}
            <select value={residenciaActiva} onChange={e => { const val = e.target.value; setResidenciaActiva(val); localStorage.setItem('pref_residencia_reservas', val); }}
              style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              {residencias.map(r => (
                <option key={r.id} value={r.nombre}>{r.nombre}</option>
              ))}
            </select>

            {/* 2. Buscar huésped a la derecha del selector de Residencia */}
            <div className="search-box">
              <span className="search-icon">🔍</span>
              <input placeholder="Buscar huésped, DNI, habitación..." value={buscar} onChange={e => setBuscar(e.target.value)} />
            </div>

            {/* 3. Filtro de Estado */}
            <select value={filtroEstado} onChange={e => setFiltroEstado(e.target.value)}
              style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              <option value="">Todos los estados</option>
              <option value="Pendiente">Pendiente</option>
              <option value="Confirmada">Confirmada</option>
              <option value="CheckIn">Check-in</option>
              <option value="CheckOut">Check-out</option>
            </select>

            {/* 4. Selector de Visibilidad y Disposición de Columnas */}
            <div className="col-picker-container">
              <button
                className="btn btn-ghost btn-sm"
                style={{ padding: '7px 10px', fontSize: 12, border: '1px solid var(--border)', borderRadius: 7 }}
                onClick={() => colConfig.setShowPicker(!colConfig.showPicker)}
                title="Configurar visibilidad y orden de columnas"
              >
                ⚙️ Columnas
              </button>
              {colConfig.showPicker && (
                <div className="col-picker-dropdown">
                  <div className="col-picker-header">
                    <span>⚙️ Configuración de Columnas</span>
                    <button
                      style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 14, color: 'var(--text-muted)' }}
                      onClick={() => colConfig.setShowPicker(false)}
                    >
                      ✕
                    </button>
                  </div>
                  <div className="col-picker-list">
                    {colConfig.columns.map(c => (
                      <label key={c.key} className="col-picker-item">
                        <input
                          type="checkbox"
                          checked={c.visible}
                          disabled={c.isFixed}
                          onChange={() => colConfig.toggleVisibility(c.key)}
                        />
                        <span style={{ fontWeight: c.visible ? 600 : 400 }}>{c.label}</span>
                      </label>
                    ))}
                  </div>
                  <div className="col-picker-footer">
                    <button className="btn btn-ghost btn-sm" style={{ fontSize: 11 }} onClick={colConfig.resetDefaults}>
                      🔄 Restablecer diseño
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* 5. Botón Nueva Reserva */}
            <button className="btn btn-primary btn-sm" style={{ marginLeft: 'auto' }} onClick={() => setShowNueva(true)}>+ Nueva reserva</button>
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
            <table style={{ fontSize: '11px', width: '100%', borderCollapse: 'collapse', tableLayout: 'fixed' }}>
              <thead>
                <tr style={{ background: 'var(--primary-dark)', color: 'white' }}>
                  {colConfig.visibleColumns.map(col => {
                    const isDragging = colConfig.draggedKey === col.key;
                    const isDragOver = colConfig.dragOverKey === col.key;
                    const isSortable = col.key === 'numeroOrden' || col.key === 'fechaEntrada' || col.key === 'fechaSalida';

                    return (
                      <th
                        key={col.key}
                        draggable={!col.isFixed}
                        onDragStart={e => colConfig.handleDragStart(col.key, e)}
                        onDragOver={e => colConfig.handleDragOver(col.key, e)}
                        onDrop={e => colConfig.handleDrop(col.key, e)}
                        className={`col-draggable ${isDragging ? 'col-dragging' : ''} ${isDragOver ? 'col-drag-over' : ''}`}
                        style={{
                          width: `${col.width}px`,
                          minWidth: `${col.width}px`,
                          padding: '4px 6px',
                          boxSizing: 'border-box',
                          cursor: col.isFixed ? 'default' : 'grab',
                          textAlign: col.key === 'totalNoches' || col.key === 'acciones' ? 'center' : 'left',
                        }}
                        onClick={() => {
                          if (isSortable) handleSort(col.key as any);
                        }}
                        title={col.isFixed ? '' : 'Haz clic y arrastra para reordenar columna'}
                      >
                        <span>
                          {col.label} {isSortable && sortField === col.key ? (sortAsc ? '▲' : '▼') : ''}
                        </span>
                        <div
                          className="col-resizer"
                          onMouseDown={e => colConfig.handleResizeStart(col.key, e)}
                          onClick={e => e.stopPropagation()}
                          title="Arrastra para ajustar el ancho"
                        />
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody>
                {paginatedReservas.map(r => (
                  <tr key={r.id} style={{ borderBottom: '1px solid var(--border-light)', height: '26px' }}>
                    {colConfig.visibleColumns.map(col => {
                      switch (col.key) {
                        case 'numeroOrden':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px', cursor: 'pointer' }} onClick={() => {
                              sessionStorage.setItem('gotoReservaId', r.id);
                              sessionStorage.setItem('gotoReservaFecha', r.fechaEntrada);
                              setPage('calendario');
                            }}>
                              <strong style={{ color: 'var(--primary)', textDecoration: 'underline' }}>#{r.numeroOrden}</strong>
                            </td>
                          );
                        case 'huesped':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={r.huespedNombreCompleto ?? ''}>
                              <strong style={{ fontSize: '11px' }}>{formatNombreTrunc(r.huespedNombreCompleto)}</strong>
                              {r.huespedDni && <div className="text-muted" style={{ fontSize: '9.5px', lineHeight: 1 }}>{r.huespedDni}</div>}
                            </td>
                          );
                        case 'residenciaHab':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                              {r.residenciaNombre} · <strong>{r.habitacionNumero}</strong>
                            </td>
                          );
                        case 'fechaEntrada':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px', fontSize: '11px', whiteSpace: 'nowrap' }}>
                              <strong>{formatFechaDisplay(r.fechaEntrada)}</strong>
                            </td>
                          );
                        case 'fechaSalida':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px', fontSize: '11px', whiteSpace: 'nowrap' }}>
                              <strong>{formatFechaDisplay(r.fechaSalida)}</strong>
                            </td>
                          );
                        case 'totalNoches':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px', textAlign: 'center' }}>{r.totalNoches}</td>
                          );
                        case 'evaluacion':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px' }}>
                              <span className="badge badge-primary" style={{ fontSize: '9.5px', padding: '1px 4px' }}>{r.evaluacion || '—'}</span>
                            </td>
                          );
                        case 'importeTotal':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px' }}>{r.importeTotal.toFixed(2)} €</td>
                          );
                        case 'estado':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px' }}>
                              <span className={`badge estado-${r.estado}`} style={{ fontSize: '9.5px', padding: '1px 4px' }}>{r.estado}</span>
                            </td>
                          );
                        case 'pagado':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px' }}>
                              {r.pagado
                                ? <span className="badge badge-success" style={{ fontSize: '9.5px', padding: '1px 4px' }}>✓ {r.formaPago}</span>
                                : <span className="badge badge-danger" style={{ fontSize: '9.5px', padding: '1px 4px' }}>Pendiente</span>}
                            </td>
                          );
                        case 'acciones':
                          return (
                            <td key={col.key} style={{ padding: '2px 6px', textAlign: 'center' }}>
                              <button className="btn btn-ghost btn-sm" style={{ padding: '1px 5px', fontSize: '10px' }} title="Editar Solicitud / Reserva" onClick={() => setEditandoReserva(r)}>✏️</button>
                            </td>
                          );
                        default:
                          return <td key={col.key}>—</td>;
                      }
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        {totalPaginas > 1 && (
          <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: 12, padding: '8px 14px', borderTop: '1px solid var(--border-light)' }}>
            <button className="btn btn-ghost btn-sm" disabled={pagina === 1} onClick={() => setPagina(p => Math.max(p - 1, 1))}>◀</button>
            <span style={{ fontSize: 11, fontWeight: 600 }}>Página {pagina} de {totalPaginas} ({reservasFiltradas.length} reservas)</span>
            <button className="btn btn-ghost btn-sm" disabled={pagina === totalPaginas} onClick={() => setPagina(p => Math.min(p + 1, totalPaginas))}>▶</button>
          </div>
        )}
      </div>

      {showNueva && (
        <NuevaReservaModal
          onSaved={() => { setShowNueva(false); setStatusMsg('Reserva creada correctamente.'); void cargar(); }}
          onCancel={() => setShowNueva(false)}
        />
      )}

      {editandoReserva && (
        <NuevaReservaModal
          itemToEdit={editandoReserva}
          isSolicitud={true}
          onSaved={() => { setEditandoReserva(null); setStatusMsg('Solicitud/Reserva actualizada correctamente.'); void cargar(); }}
          onCancel={() => setEditandoReserva(null)}
        />
      )}

    </>
  );
}