import { useState, useEffect } from 'react';
import { apiFetch, formatFechaDisplay } from '../../api';
import { NuevaReservaModal } from '../reservas/NuevaReservaModal';

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

  const [residencias, setResidencias] = useState<any[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>(
    () => localStorage.getItem('pref_residencia_solicitudes') || 'Todas'
  );

  const cargar = async () => {
    setLoading(true);
    try {
      const res = await apiFetch<any>('/api/reservas?incluirCanceladas=true&pageSize=500');
      const data = Array.isArray(res) ? res : (res?.items ?? []);
      const filtradas = data.filter((r: any) => !r.esBloqueo);
      setSolicitudes(filtradas);
    } catch { } finally { setLoading(false); }
  };

  const cargarResidencias = async () => {
    try {
      const list = await apiFetch<any[]>('/api/residencias');
      setResidencias(list);
      if (list.length > 0) {
        const saved = localStorage.getItem('pref_residencia_solicitudes');
        if (!saved || saved === 'Todas' || !list.some((r: any) => r.nombre === saved)) {
          setResidenciaActiva(list[0].nombre);
          localStorage.setItem('pref_residencia_solicitudes', list[0].nombre);
        }
      }
    } catch { }
  };

  useEffect(() => {
    void cargar();
    void cargarResidencias();
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

  // Filtered by residence and search
  const solicitudesFiltradas = solicitudes.filter(r => {
    const matchResidencia = r.residenciaNombre === residenciaActiva;
    if (!matchResidencia) return false;

    if (!buscar.trim()) return true;
    const query = buscar.toLowerCase();
    return (
      r.numeroOrden?.toString().includes(query) ||
      (r.huespedNombreCompleto && r.huespedNombreCompleto.toLowerCase().includes(query)) ||
      (r.huespedNombre && r.huespedNombre.toLowerCase().includes(query)) ||
      (r.huespedApellidos && r.huespedApellidos.toLowerCase().includes(query)) ||
      (r.huespedDni && r.huespedDni.toLowerCase().includes(query)) ||
      (r.huespedEmail && r.huespedEmail.toLowerCase().includes(query)) ||
      (r.huespedTelefono && r.huespedTelefono.toLowerCase().includes(query)) ||
      (r.residenciaNombre && r.residenciaNombre.toLowerCase().includes(query))
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
        <div className="card-header" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <h3 style={{ margin: 0 }}>👥 Solicitudes registradas</h3>
            <select
              value={residenciaActiva}
              onChange={e => {
                const val = e.target.value;
                setResidenciaActiva(val);
                localStorage.setItem('pref_residencia_solicitudes', val);
                setPagina(1);
              }}
              style={{ padding: '6px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13, fontWeight: 600, background: 'var(--surface-1)', cursor: 'pointer' }}
            >
              {residencias.map(r => (
                <option key={r.id} value={r.nombre}>{r.nombre}</option>
              ))}
            </select>
          </div>
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
                    <td style={{ padding: '5px 8px' }}><strong>{formatFechaDisplay(r.fechaEntrada)}</strong></td>
                    <td style={{ padding: '5px 8px' }}><strong>{formatFechaDisplay(r.fechaSalida)}</strong></td>
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
        <NuevaReservaModal
          itemToEdit={editando}
          isSolicitud={true}
          onSaved={() => { setEditando(null); setStatusMsg('Solicitud actualizada correctamente.'); void cargar(); }}
          onCancel={() => setEditando(null)}
        />
      )}
    </>
  );
}