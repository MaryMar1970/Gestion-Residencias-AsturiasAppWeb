import { useState } from 'react';
import { apiFetch, formatFechaDisplay } from '../../api';
import { Reserva } from '../../types';

// ─── MODAL DETALLE RESERVA ─────────────────────────────────────────────────────
export function ReservaDetalleModal({ reserva, onClose }: { reserva: Reserva; onClose: () => void }) {
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
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Entrada</td><td><strong>{formatFechaDisplay(reserva.fechaEntrada)}</strong></td></tr>
              <tr><td style={{ padding: '5px 0', color: 'var(--text-muted)' }}>Salida</td><td><strong>{formatFechaDisplay(reserva.fechaSalida)}</strong></td></tr>
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
