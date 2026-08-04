import { useState, useEffect } from 'react';
import { apiFetch } from '../../api';

export function BloqueoModal({ item, onSaved, onCancel }: { item?: any; onSaved: () => void; onCancel: () => void }) {
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
          <h3>{item ? '✏️ Editar' : '➕ Nuevo'} Bloqueo / Reserva Interna</h3>
          <button className="modal-close" onClick={onCancel}>✕</button>
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