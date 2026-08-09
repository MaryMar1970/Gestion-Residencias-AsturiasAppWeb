import { useState, useEffect, FormEvent } from 'react';
import { apiFetch } from '../../api';
import { TipoHab, Tarifa } from '../../types';

function TarifaForm({ item, defaultResidenciaId, onSaved, onCancel }: { item?: Tarifa; defaultResidenciaId: string; onSaved: () => void; onCancel: () => void }) {
  const [tipos, setTipos] = useState<TipoHab[]>([]);
  const [residencias, setResidencias] = useState<any[]>([]);
  const [f, setF] = useState({
    residenciaId: item?.residenciaId ?? defaultResidenciaId,
    tipoHabitacionId: item?.tipoHabitacionId ?? '',
    nombreTarifa: item?.nombreTarifa ?? '',
    precioNoche: item?.precioNoche ?? 0,
    precioMes: item?.precioMes ?? 0,
    porcentajeIva: item?.porcentajeIva ?? 10,
    descripcion: item?.descripcion ?? '',
    activa: item?.activa ?? true,
    vigenteDesde: null as string | null,
    vigenteHasta: null as string | null
  });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));

  useEffect(() => {
    Promise.all([
      apiFetch<TipoHab[]>('/api/tipos-habitacion'),
      apiFetch<any[]>('/api/residencias')
    ]).then(([t, r]) => {
      setTipos(t);
      setResidencias(r);
    });
  }, []);

  const submit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      const url = item ? `/api/tarifas/${item.id}` : '/api/tarifas';
      await apiFetch(url, { method: item ? 'PUT' : 'POST', body: JSON.stringify(f) });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  return (
    <form onSubmit={submit}>
      <div className="modal-body">
        <div className="form-grid">
          <div className="form-group">
            <label>Residencia *</label>
            <select value={f.residenciaId} onChange={e => set('residenciaId', e.target.value)} required>
              <option value="">Seleccione residencia...</option>
              {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
            </select>
          </div>
          <div className="form-group">
            <label>Tipo habitación *</label>
            <select value={f.tipoHabitacionId} onChange={e => set('tipoHabitacionId', e.target.value)} required>
              <option value="">Seleccione tipo...</option>
              {tipos.map(t => <option key={t.id} value={t.id}>{t.nombre}</option>)}
            </select>
          </div>
          <div className="form-group" style={{ gridColumn: 'span 2' }}>
            <label>Nombre tarifa *</label>
            <input value={f.nombreTarifa} onChange={e => set('nombreTarifa', e.target.value)} placeholder="Estudiante, Investigador, Externo..." required />
          </div>
          <div className="form-group">
            <label>Precio / noche (€)</label>
            <input type="number" step="0.01" min={0} value={f.precioNoche} onChange={e => set('precioNoche', +e.target.value)} />
          </div>
          <div className="form-group">
            <label>Precio / mes (€)</label>
            <input type="number" step="0.01" min={0} value={f.precioMes} onChange={e => set('precioMes', +e.target.value)} />
          </div>
          <div className="form-group">
            <label>IVA (%)</label>
            <input type="number" step="0.01" min={0} max={100} value={f.porcentajeIva} onChange={e => set('porcentajeIva', +e.target.value)} />
          </div>
          <div className="form-group">
            <label>Descripción</label>
            <input value={f.descripcion} onChange={e => set('descripcion', e.target.value)} />
          </div>
        </div>
        <div style={{ marginTop: 14 }}>
          <label className="toggle-wrap">
            <span className="toggle">
              <input type="checkbox" checked={f.activa} onChange={e => set('activa', e.target.checked)} />
              <span className="toggle-slider"></span>
            </span>
            Tarifa activa
          </label>
        </div>
      </div>
      <div className="modal-footer">
        <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Guardando...' : '💾 Guardar'}</button>
      </div>
    </form>
  );
}

export function TarifasPage() {
  const [tarifas, setTarifas] = useState<Tarifa[]>([]);
  const [residencias, setResidencias] = useState<any[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>(() => localStorage.getItem('pref_residencia_tarifas') || '');
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState<Tarifa | undefined>(undefined);

  const cargarResidencias = async () => {
    try {
      const res = await apiFetch<any[]>('/api/residencias');
      setResidencias(res);
      const saved = localStorage.getItem('pref_residencia_tarifas');
      if (saved && res.some(r => r.id === saved)) {
        setResidenciaActiva(saved);
      } else if (res.length > 0 && !residenciaActiva) {
        setResidenciaActiva(res[0].id);
      }
    } catch {}
  };

  const cargarTarifas = async () => {
    if (!residenciaActiva) return;
    setLoading(true);
    try {
      const data = await apiFetch<Tarifa[]>(`/api/tarifas?residenciaId=${residenciaActiva}`);
      setTarifas(data);
    } catch {}
    finally { setLoading(false); }
  };

  useEffect(() => {
    cargarResidencias();
  }, []);

  useEffect(() => {
    cargarTarifas();
  }, [residenciaActiva]);

  const handleEdit = (item: Tarifa) => {
    setEditItem(item);
    setShowModal(true);
  };

  const handleDelete = async (id: string) => {
    if (!confirm('¿Está seguro de eliminar esta tarifa?')) return;
    try {
      await apiFetch(`/api/tarifas/${id}`, { method: 'DELETE' });
      cargarTarifas();
    } catch (e: any) {
      alert(e.message);
    }
  };

  return (
    <div className="card">
      <div className="card-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ fontSize: 24 }}>💰</span>
          <div>
            <h3>Tarifas y precios por Residencia</h3>
            <p className="text-muted" style={{ fontSize: 12, margin: 0 }}>Configure las tarifas y precios de cada una de las residencias</p>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
          <select value={residenciaActiva} onChange={e => { const val = e.target.value; setResidenciaActiva(val); localStorage.setItem('pref_residencia_tarifas', val); }}
            style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
            {residencias.map(r => (
              <option key={r.id} value={r.id}>{r.nombre}</option>
            ))}
          </select>
          <button className="btn btn-primary btn-sm" onClick={() => { setEditItem(undefined); setShowModal(true); }}>+ Nueva tarifa</button>
        </div>
      </div>

      <div className="table-container">
        {loading ? (
          <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div>
        ) : tarifas.length === 0 ? (
          <div className="empty-state">
            <div className="empty-icon">💰</div>
            <p>No hay tarifas configuradas para esta residencia</p>
            <button className="btn btn-primary mt-4" onClick={() => { setEditItem(undefined); setShowModal(true); }}>+ Configurar primera tarifa</button>
          </div>
        ) : (
          <table>
            <thead>
              <tr>
                <th>Tipo habitación</th>
                <th>Nombre tarifa</th>
                <th>Precio / noche</th>
                <th>Precio / mes</th>
                <th>IVA</th>
                <th style={{ width: 120 }}>Acciones</th>
              </tr>
            </thead>
            <tbody>
              {tarifas.map(t => (
                <tr key={t.id}>
                  <td>{t.tipoHabitacionNombre}</td>
                  <td><strong>{t.nombreTarifa}</strong></td>
                  <td>{t.precioNoche > 0 ? `${t.precioNoche.toFixed(2)} €` : '—'}</td>
                  <td>{t.precioMes > 0 ? `${t.precioMes.toFixed(2)} €` : '—'}</td>
                  <td><span className="badge badge-accent">{t.porcentajeIva}%</span></td>
                  <td>
                    <div className="actions-cell">
                      <button className="btn btn-ghost btn-sm" onClick={() => handleEdit(t)}>✏️ Editar</button>
                      <button className="btn btn-danger btn-sm" onClick={() => handleDelete(t.id)}>🗑️</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {showModal && (
        <div className="modal-overlay" onClick={e => e.target === e.currentTarget && setShowModal(false)}>
          <div className="modal" style={{ maxWidth: 550 }}>
            <div className="modal-header">
              <h3>{editItem ? '✏️ Editar Tarifa' : '➕ Nueva Tarifa'}</h3>
              <button className="modal-close" onClick={() => setShowModal(false)}>✕</button>
            </div>
            <TarifaForm
              item={editItem}
              defaultResidenciaId={residenciaActiva}
              onSaved={() => { setShowModal(false); cargarTarifas(); }}
              onCancel={() => setShowModal(false)}
            />
          </div>
        </div>
      )}
    </div>
  );
}