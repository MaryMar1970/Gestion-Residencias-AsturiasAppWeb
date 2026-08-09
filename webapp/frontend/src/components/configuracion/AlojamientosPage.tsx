import { useState, useEffect, FormEvent } from 'react';
import { apiFetch } from '../../api';
import { Residencia, TipoHab, Habitacion } from '../../types';
import { CrudPage } from '../shared/CrudPage';

function TipoHabForm({ item, onSaved, onCancel }: { item?: TipoHab; onSaved: () => void; onCancel: () => void }) {
  const [f, setF] = useState({ nombre: item?.nombre ?? '', codigo: item?.codigo ?? '', capacidadMaxima: item?.capacidadMaxima ?? 1, admiteSupletorias: item?.admiteSupletorias ?? false, descripcion: item?.descripcion ?? '', activo: item?.activo ?? true, orden: item?.orden ?? 0 });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));
  const submit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      const url = item ? `/api/tipos-habitacion/${item.id}` : '/api/tipos-habitacion';
      await apiFetch(url, { method: item ? 'PUT' : 'POST', body: JSON.stringify(f) });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };
  return (
    <form onSubmit={submit}>
      <div className="modal-body">
        <div className="form-grid">
          <div className="form-group"><label>Nombre *</label><input value={f.nombre} onChange={e => set('nombre', e.target.value)} placeholder="Ej: Individual, Doble, Apartamento" required /></div>
          <div className="form-group"><label>Código *</label><input value={f.codigo} onChange={e => set('codigo', e.target.value.toUpperCase())} placeholder="IND, DOB, APT..." maxLength={10} required /></div>
          <div className="form-group"><label>Capacidad máxima (personas)</label><input type="number" min={1} max={20} value={f.capacidadMaxima} onChange={e => set('capacidadMaxima', +e.target.value)} /></div>
          <div className="form-group"><label>Orden en listados</label><input type="number" value={f.orden} onChange={e => set('orden', +e.target.value)} /></div>
          <div className="form-group" style={{ gridColumn: '1 / -1' }}><label>Descripción</label><textarea value={f.descripcion} onChange={e => set('descripcion', e.target.value)} /></div>
        </div>
        <div style={{ marginTop: 14, display: 'flex', gap: 20 }}>
          <label className="toggle-wrap"><span className="toggle"><input type="checkbox" checked={f.admiteSupletorias} onChange={e => set('admiteSupletorias', e.target.checked)} /><span className="toggle-slider"></span></span> Admite camas supletorias</label>
          <label className="toggle-wrap"><span className="toggle"><input type="checkbox" checked={f.activo} onChange={e => set('activo', e.target.checked)} /><span className="toggle-slider"></span></span> Tipo activo</label>
        </div>
      </div>
      <div className="modal-footer">
        <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Guardando...' : '💾 Guardar'}</button>
      </div>
    </form>
  );
}

function TiposPage() {
  return <CrudPage<TipoHab> title="Tipos de habitación" icon="🛏️" fetchUrl="/api/tipos-habitacion"
    columns={['Nombre', 'Código', 'Capacidad', 'Supletorias']}
    emptyLabel="No hay tipos de habitación configurados"
    FormComponent={TipoHabForm}
    renderRow={(t: TipoHab, onEdit: () => void, onDelete: () => void) => (
      <tr key={t.id}>
        <td><strong>{t.nombre}</strong>{t.descripcion && <div className="text-muted" style={{ fontSize: 11 }}>{t.descripcion}</div>}</td>
        <td><span className="badge badge-primary">{t.codigo}</span></td>
        <td>{t.capacidadMaxima} persona{t.capacidadMaxima !== 1 ? 's' : ''}</td>
        <td>{t.admiteSupletorias ? <span className="badge badge-success">✓ Sí</span> : <span className="badge badge-danger">✕ No</span>}</td>
        <td><div className="actions-cell"><button className="btn btn-ghost btn-sm" onClick={onEdit}>✏️ Editar</button><button className="btn btn-danger btn-sm" onClick={onDelete}>🗑️</button></div></td>
      </tr>
    )} />;
}

// ─── Habitaciones Page ────────────────────────────────────────────────────────
function HabitacionForm({ item, onSaved, onCancel }: { item?: Habitacion; onSaved: () => void; onCancel: () => void }) {
  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [tipos, setTipos] = useState<TipoHab[]>([]);
  const [f, setF] = useState({
    residenciaId: item?.residenciaId ?? '',
    tipoHabitacionId: item?.tipoHabitacionId ?? '',
    numero: item?.numero ?? '',
    nombre: item?.nombre ?? '',
    tipoCamaPrincipal: item?.tipoCamaPrincipal ?? 'IND',
    capacidadPersonas: item?.capacidadPersonas ?? 1,
    admiteSupletorias: item?.admiteSupletorias ?? false,
    plazasSupletorias: item?.plazasSupletorias ?? 0,
    activa: item?.activa ?? true,
    notas: item?.notas ?? '',
    orden: item?.orden ?? 0
  });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));

  useEffect(() => {
    Promise.all([apiFetch<Residencia[]>('/api/residencias'), apiFetch<TipoHab[]>('/api/tipos-habitacion')])
      .then(([r, t]) => {
        setResidencias(r); setTipos(t);
        if (!item && r.length > 0) set('residenciaId', r[0].id);
        if (!item && t.length > 0) set('tipoHabitacionId', t[0].id);
      });
  }, []);

  const handleTipoHabChange = (tipoId: string) => {
    const t = tipos.find(x => x.id === tipoId);
    set('tipoHabitacionId', tipoId);
    if (t) {
      set('capacidadPersonas', t.capacidadMaxima);
      set('admiteSupletorias', t.admiteSupletorias);
      if (t.capacidadMaxima === 1) {
        set('tipoCamaPrincipal', 'IND');
        set('plazasSupletorias', 0);
      } else if (t.capacidadMaxima === 2) {
        set('tipoCamaPrincipal', 'TWIN');
        set('plazasSupletorias', 0);
      } else if (t.capacidadMaxima > 2) {
        set('tipoCamaPrincipal', 'TWIN');
        set('admiteSupletorias', true);
        set('plazasSupletorias', t.capacidadMaxima - 2);
      }
    }
  };

  const handleCapacidadChange = (cap: number) => {
    const val = Math.min(6, Math.max(1, cap));
    set('capacidadPersonas', val);
    if (val === 1) {
      set('tipoCamaPrincipal', 'IND');
      set('plazasSupletorias', 0);
    } else if (val === 2) {
      if (f.tipoCamaPrincipal === 'IND') set('tipoCamaPrincipal', 'TWIN');
      set('plazasSupletorias', 0);
    } else {
      if (f.tipoCamaPrincipal === 'IND') set('tipoCamaPrincipal', 'TWIN');
      set('admiteSupletorias', true);
      set('plazasSupletorias', val - 2);
    }
  };

  const submit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      const url = item ? `/api/habitaciones/${item.id}` : '/api/habitaciones';
      await apiFetch(url, { method: item ? 'PUT' : 'POST', body: JSON.stringify(f) });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  return (
    <form onSubmit={submit}>
      <div className="modal-body">
        <div className="form-grid">
          <div className="form-group"><label>Residencia *</label>
            <select value={f.residenciaId} onChange={e => set('residenciaId', e.target.value)} required>
              <option value="">Seleccione residencia...</option>
              {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
            </select>
          </div>
          <div className="form-group"><label>Categoría / Tipo de Habitación *</label>
            <select value={f.tipoHabitacionId} onChange={e => handleTipoHabChange(e.target.value)} required>
              <option value="">Seleccione tipo...</option>
              {tipos.map(t => <option key={t.id} value={t.id}>{t.nombre} ({t.codigo})</option>)}
            </select>
          </div>
          <div className="form-group"><label>Número / Identificador *</label><input value={f.numero} onChange={e => set('numero', e.target.value)} placeholder="101, Apt.1, EST.2..." required /></div>
          <div className="form-group"><label>Nombre descriptivo</label><input value={f.nombre} onChange={e => set('nombre', e.target.value)} placeholder="Ej: Habitación Vistas Mar / Estudio 1" /></div>
          
          <div className="form-group">
            <label>Tipo de Cama Principal *</label>
            <select value={f.tipoCamaPrincipal} onChange={e => set('tipoCamaPrincipal', e.target.value)}>
              <option value="IND">Individual - IND (1 pax)</option>
              <option value="TWIN">2 Camas Separadas - TWIN (2 pax)</option>
              <option value="MAT">Doble Matrimonio - MAT (2 pax)</option>
            </select>
          </div>

          <div className="form-group">
            <label>Capacidad Máxima Ocupantes (1 - 6 pax)</label>
            <input type="number" min={1} max={6} value={f.capacidadPersonas} onChange={e => handleCapacidadChange(+e.target.value)} required />
          </div>

          <div className="form-group">
            <label>Plazas Supletorias (SUP - 1 pax c/u)</label>
            <input type="number" min={0} max={4} value={f.plazasSupletorias} onChange={e => set('plazasSupletorias', +e.target.value)} disabled={!f.admiteSupletorias} />
          </div>

          <div className="form-group"><label>Orden en listado / calendario</label><input type="number" value={f.orden} onChange={e => set('orden', +e.target.value)} /></div>
          <div className="form-group" style={{ gridColumn: '1 / -1' }}><label>Notas u Observaciones</label><input value={f.notas} onChange={e => set('notas', e.target.value)} placeholder="Detalles de equipamiento o estado" /></div>
        </div>

        {/* Banner Explicativo de Distribución de Plazas */}
        <div style={{ marginTop: 16, padding: '12px 14px', background: 'var(--bg-muted)', borderLeft: '4px solid var(--primary)', borderRadius: 6, fontSize: 13 }}>
          💡 <strong>Desglose de Alojamiento:</strong> Cama principal ({f.tipoCamaPrincipal === 'IND' ? '1 pax Individual' : '2 pax Doble'})
          {f.admiteSupletorias && f.plazasSupletorias > 0 ? ` + ${f.plazasSupletorias} Camas Supletorias (SUP)` : ''} 
          &nbsp;= <strong>Capacidad Total {f.capacidadPersonas} personas</strong>.
        </div>

        <div style={{ marginTop: 14, display: 'flex', gap: 20 }}>
          <label className="toggle-wrap"><span className="toggle"><input type="checkbox" checked={f.admiteSupletorias} onChange={e => set('admiteSupletorias', e.target.checked)} /><span className="toggle-slider"></span></span> Admite camas supletorias (SUP)</label>
          <label className="toggle-wrap"><span className="toggle"><input type="checkbox" checked={f.activa} onChange={e => set('activa', e.target.checked)} /><span className="toggle-slider"></span></span> Habitación activa</label>
        </div>
      </div>
      <div className="modal-footer">
        <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Guardando...' : '💾 Guardar Alojamiento'}</button>
      </div>
    </form>
  );
}

// Modal de Creación Rápida en Lote
function CrearLoteModal({ onSaved, onCancel }: { onSaved: () => void; onCancel: () => void }) {
  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [tipos, setTipos] = useState<TipoHab[]>([]);
  const [f, setF] = useState({
    residenciaId: '',
    tipoHabitacionId: '',
    prefijo: 'Hab. ',
    numeroInicio: 101,
    cantidad: 10,
    tipoCamaPrincipal: 'DOB',
    capacidadPersonas: 2,
    admiteSupletorias: false,
    plazasSupletorias: 0
  });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));

  useEffect(() => {
    Promise.all([apiFetch<Residencia[]>('/api/residencias'), apiFetch<TipoHab[]>('/api/tipos-habitacion')])
      .then(([r, t]) => {
        setResidencias(r); setTipos(t);
        if (r.length > 0) set('residenciaId', r[0].id);
        if (t.length > 0) set('tipoHabitacionId', t[0].id);
      });
  }, []);

  const handleTipoHabChange = (tipoId: string) => {
    const t = tipos.find(x => x.id === tipoId);
    set('tipoHabitacionId', tipoId);
    if (t) {
      set('capacidadPersonas', t.capacidadMaxima);
      set('admiteSupletorias', t.admiteSupletorias);
      if (t.capacidadMaxima > 2) {
        set('tipoCamaPrincipal', 'DOB');
        set('plazasSupletorias', t.capacidadMaxima - 2);
      } else {
        set('tipoCamaPrincipal', t.capacidadMaxima === 1 ? 'IND' : 'DOB');
        set('plazasSupletorias', 0);
      }
    }
  };

  const submit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      await apiFetch('/api/habitaciones/lote', { method: 'POST', body: JSON.stringify(f) });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onCancel()}>
      <div className="modal" style={{ maxWidth: 620 }}>
        <div className="modal-header">
          <h3>⚡ Generador Rápido de Habitaciones por Lote</h3>
          <button className="btn-close" onClick={onCancel}>✕</button>
        </div>
        <form onSubmit={submit}>
          <div className="modal-body">
            <p style={{ fontSize: 13, color: 'var(--text-muted)', marginBottom: 16 }}>
              Cree de forma masiva una secuencia numerada de habitaciones asignadas a una Residencia.
            </p>
            <div className="form-grid">
              <div className="form-group"><label>Residencia *</label>
                <select value={f.residenciaId} onChange={e => set('residenciaId', e.target.value)} required>
                  {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
                </select>
              </div>
              <div className="form-group"><label>Tipo de Habitación *</label>
                <select value={f.tipoHabitacionId} onChange={e => handleTipoHabChange(e.target.value)} required>
                  {tipos.map(t => <option key={t.id} value={t.id}>{t.nombre} ({t.codigo})</option>)}
                </select>
              </div>
              <div className="form-group"><label>Prefijo numeración</label><input value={f.prefijo} onChange={e => set('prefijo', e.target.value)} placeholder="Ej: Hab. / Apt. / " /></div>
              <div className="form-group"><label>Número Inicio *</label><input type="number" value={f.numeroInicio} onChange={e => set('numeroInicio', +e.target.value)} required /></div>
              <div className="form-group"><label>Cantidad a Generar *</label><input type="number" min={1} max={50} value={f.cantidad} onChange={e => set('cantidad', +e.target.value)} required /></div>
              <div className="form-group">
                <label>Tipo Cama Principal</label>
                <select value={f.tipoCamaPrincipal} onChange={e => set('tipoCamaPrincipal', e.target.value)}>
                  <option value="IND">Individual - IND (1 pax)</option>
                  <option value="DOB">Doble 2 camas - DOB (2 pax)</option>
                  <option value="MAT">Doble Matrimonio - MAT (2 pax)</option>
                </select>
              </div>
              <div className="form-group"><label>Capacidad total (pax)</label><input type="number" min={1} max={6} value={f.capacidadPersonas} onChange={e => set('capacidadPersonas', +e.target.value)} /></div>
              <div className="form-group"><label>Plazas supletorias (SUP)</label><input type="number" min={0} max={4} value={f.plazasSupletorias} onChange={e => set('plazasSupletorias', +e.target.value)} /></div>
            </div>
          </div>
          <div className="modal-footer">
            <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
            <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Generando...' : `⚡ Generar ${f.cantidad} habitaciones`}</button>
          </div>
        </form>
      </div>
    </div>
  );
}

function HabitacionesPage() {
  const [items, setItems] = useState<Habitacion[]>([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [showBatchModal, setShowBatchModal] = useState(false);
  const [editing, setEditing] = useState<Habitacion | undefined>(undefined);
  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>(() => localStorage.getItem('pref_residencia_alojamientos') || 'Todas');

  const load = async () => {
    setLoading(true);
    try {
      const [habList, resList] = await Promise.all([
        apiFetch<Habitacion[]>('/api/habitaciones'),
        apiFetch<Residencia[]>('/api/residencias')
      ]);
      setItems(habList);
      setResidencias(resList);
    } catch {} finally { setLoading(false); }
  };

  useEffect(() => { void load(); }, []);

  const handleDelete = async (item: Habitacion) => {
    if (!confirm(`¿Eliminar la habitación ${item.numero} (${item.residenciaNombre})?`)) return;
    try {
      await apiFetch(`/api/habitaciones/${item.id}`, { method: 'DELETE' });
      void load();
    } catch (e: any) { alert(e.message); }
  };

  const handleEdit = (item: Habitacion) => { setEditing(item); setShowModal(true); };
  const handleNew  = () => { setEditing(undefined); setShowModal(true); };
  const handleSaved = () => { setShowModal(false); setShowBatchModal(false); setEditing(undefined); void load(); };

  const filtradas = items.filter(h => residenciaActiva === 'Todas' || h.residenciaNombre === residenciaActiva);

  // Cálculos Estadísticos
  const totalHabitaciones = filtradas.length;
  const totalPlazasBase = filtradas.reduce((acc, h) => acc + (h.tipoCamaPrincipal === 'IND' ? 1 : 2), 0);
  const totalSupletorias = filtradas.reduce((acc, h) => acc + (h.admiteSupletorias ? (h.plazasSupletorias ?? 0) : 0), 0);
  const capacidadTotal = totalPlazasBase + totalSupletorias;

  return (
    <>
      {/* Tarjetas Informativas de Capacidad */}
      <div className="stats-row" style={{ marginBottom: 18 }}>
        <div className="stat-card">
          <div className="stat-icon">🔑</div>
          <div className="stat-value">{totalHabitaciones}</div>
          <div className="stat-label">Alojamientos / Habitaciones</div>
        </div>
        <div className="stat-card">
          <div className="stat-icon">🛏️</div>
          <div className="stat-value">{totalPlazasBase}</div>
          <div className="stat-label">Plazas Camas Principales (IND/DOB/MAT)</div>
        </div>
        <div className="stat-card">
          <div className="stat-icon">➕</div>
          <div className="stat-value">{totalSupletorias}</div>
          <div className="stat-label">Camas Supletorias (SUP)</div>
        </div>
        <div className="stat-card">
          <div className="stat-icon">👥</div>
          <div className="stat-value" style={{ color: 'var(--primary)' }}>{capacidadTotal}</div>
          <div className="stat-label">Capacidad Máxima Ocupantes</div>
        </div>
      </div>

      <div className="card">
        <div className="card-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 10 }}>
          <h3>🔑 Configuración de Alojamientos y Habitaciones</h3>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <label style={{ fontSize: 13, fontWeight: 600, color: 'var(--text-muted)' }}>Residencia:</label>
            <select value={residenciaActiva} onChange={e => { const val = e.target.value; setResidenciaActiva(val); localStorage.setItem('pref_residencia_alojamientos', val); }}
              style={{ padding: '6px 12px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13, fontWeight: 500 }}>
              <option value="Todas">Todas las residencias</option>
              {residencias.map(r => (
                <option key={r.id} value={r.nombre}>{r.nombre}</option>
              ))}
            </select>
            <button className="btn btn-ghost btn-sm" onClick={() => setShowBatchModal(true)}>⚡ Crear por Lote</button>
            <button className="btn btn-primary btn-sm" onClick={handleNew}>+ Nueva Habitación</button>
          </div>
        </div>
        <div className="table-container">
          {loading ? (
            <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div>
          ) : filtradas.length === 0 ? (
            <div className="empty-state">
              <div className="empty-icon">🔑</div>
              <p>No hay habitaciones configuradas para la residencia seleccionada</p>
              <div style={{ display: 'flex', gap: 10, marginTop: 12 }}>
                <button className="btn btn-ghost" onClick={() => setShowBatchModal(true)}>⚡ Generar por Lote</button>
                <button className="btn btn-primary" onClick={handleNew}>+ Nueva habitación</button>
              </div>
            </div>
          ) : (
            <table>
              <thead>
                <tr>
                  <th>Residencia</th>
                  <th>Nº / ID</th>
                  <th>Tipo Habitación & Cama Principal</th>
                  <th>Capacidad Ocupantes</th>
                  <th>Camas Supletorias</th>
                  <th>Estado</th>
                  <th style={{ width: 120 }}>Acciones</th>
                </tr>
              </thead>
              <tbody>
                {filtradas.map(h => {
                  const badgeBed = h.tipoCamaPrincipal === 'MAT' ? 'badge-primary' : (h.tipoCamaPrincipal === 'TWIN' || h.tipoCamaPrincipal === 'DOB') ? 'badge-info' : 'badge-secondary';
                  const labelBed = h.tipoCamaPrincipal === 'MAT' ? '🛏️ Matrimonio (MAT)' : (h.tipoCamaPrincipal === 'TWIN' || h.tipoCamaPrincipal === 'DOB') ? '🛏️ 2 Camas (TWIN)' : '🛏️ Individual (IND)';
                  return (
                    <tr key={h.id}>
                      <td><strong>{h.residenciaNombre}</strong></td>
                      <td>
                        <span style={{ fontSize: 14, fontWeight: 700 }}>{h.numero}</span>
                        {h.nombre && <div className="text-muted" style={{ fontSize: 11 }}>{h.nombre}</div>}
                      </td>
                      <td>
                        <span className="badge badge-primary">{h.tipoCodigo}</span> <strong>{h.tipoNombre}</strong>
                        <div style={{ marginTop: 4 }}>
                          <span className={`badge ${badgeBed}`} style={{ fontSize: 10.5, padding: '2px 7px' }}>{labelBed}</span>
                        </div>
                      </td>
                      <td>
                        <strong style={{ fontSize: 14 }}>{h.capacidadPersonas} pax</strong>
                      </td>
                      <td>
                        {h.admiteSupletorias && (h.plazasSupletorias ?? 0) > 0 ? (
                          <span className="badge badge-success">✓ +{h.plazasSupletorias} SUP</span>
                        ) : (
                          <span className="badge badge-secondary" style={{ opacity: 0.6 }}>Sin supletoria</span>
                        )}
                      </td>
                      <td><span className={`badge ${h.activa ? 'badge-success' : 'badge-danger'}`}>{h.activa ? '✓ Activa' : '✕ Bloqueada'}</span></td>
                      <td>
                        <div className="actions-cell">
                          <button className="btn btn-ghost btn-sm" onClick={() => handleEdit(h)}>✏️ Editar</button>
                          <button className="btn btn-danger btn-sm" onClick={() => handleDelete(h)}>🗑️</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {showModal && (
        <div className="modal-overlay" onClick={e => e.target === e.currentTarget && setShowModal(false)}>
          <div className="modal">
            <div className="modal-header">
              <h3>{editing ? '✏️ Editar Alojamiento' : '➕ Configurar Nuevo Alojamiento'}</h3>
              <button className="btn-close" onClick={() => setShowModal(false)}>✕</button>
            </div>
            <HabitacionForm item={editing} onSaved={handleSaved} onCancel={() => setShowModal(false)} />
          </div>
        </div>
      )}

      {showBatchModal && (
        <CrearLoteModal onSaved={handleSaved} onCancel={() => setShowBatchModal(false)} />
      )}
    </>
  );
}

export function AlojamientosPage() {
  const [tab, setTab] = useState<'tipos' | 'habitaciones'>('tipos');
  return (
    <>
      <div style={{ display: 'flex', gap: 10, marginBottom: 15, padding: '0 5px' }}>
        <button className={`btn ${tab === 'tipos' ? 'btn-primary' : 'btn-ghost'}`} style={{ borderRadius: 8, padding: '8px 16px', fontWeight: 600 }} onClick={() => setTab('tipos')}>
          🛏️ 1. Categorías / Tipos de alojamiento
        </button>
        <button className={`btn ${tab === 'habitaciones' ? 'btn-primary' : 'btn-ghost'}`} style={{ borderRadius: 8, padding: '8px 16px', fontWeight: 600 }} onClick={() => setTab('habitaciones')}>
          🔑 2. Habitaciones / Apartamentos por Residencia
        </button>
      </div>
      <div>
        {tab === 'tipos' ? <TiposPage /> : <HabitacionesPage />}
      </div>
    </>
  );
}