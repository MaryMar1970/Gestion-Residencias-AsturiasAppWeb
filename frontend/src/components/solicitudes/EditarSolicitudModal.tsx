import React, { useState } from 'react';
import { apiFetch } from '../../api';

export function EditarSolicitudModal({
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
            telefono: telefono.trim() || null,
            email: currentGuest?.email?.trim() || null,
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
          <button className="modal-close" onClick={onCancel}>✕</button>
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
