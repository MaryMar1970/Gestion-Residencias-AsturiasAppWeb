import { useState } from 'react';
import { formatFechaDisplay } from '../../api';
import { BloqueoModal } from './BloqueoModal';
import { useBloqueosData } from '../../hooks/useBloqueosData';

export function BloqueosPage() {
  const {
    loading,
    buscar, setBuscar,
    statusMsg, setStatusMsg,
    residencias,
    residenciaActiva, setResidenciaActiva,
    sortField, sortAsc, handleSort,
    bloqueosOrdenados,
    eliminar,
    cargar,
  } = useBloqueosData();

  const [showModal, setShowModal] = useState(false);
  const [editando, setEditando] = useState<any | null>(null);

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
            
            <select value={residenciaActiva} onChange={e => { const val = e.target.value; setResidenciaActiva(val); localStorage.setItem('pref_residencia_bloqueos', val); }}
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
          ) : bloqueosOrdenados.length === 0 ? (
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
                    <td><strong>{formatFechaDisplay(r.fechaEntrada)}</strong></td>
                    <td><strong>{formatFechaDisplay(r.fechaSalida)}</strong></td>
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