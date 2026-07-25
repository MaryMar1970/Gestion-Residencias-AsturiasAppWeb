import { useState, useEffect, FormEvent } from 'react';
import './index.css';
import { apiFetch } from './api';
import { CalendarioGantt, NuevaReservaModal, ReservasPage, HuespedesPage, BloqueosPage } from './Reservas';
import { FacturasPage } from './Facturas';

// ─── Types ────────────────────────────────────────────────────────────────────
type User = { id: string; email: string; fullName: string; roles: string[] };
type Residencia = { id: string; nombre: string; razonSocial?: string; cif?: string; direccion?: string; codigoPostal?: string; municipio?: string; provincia?: string; telefono?: string; email?: string; activa: boolean; orden: number; totalHabitaciones: number };
type TipoHab   = { id: string; nombre: string; codigo: string; capacidadMaxima: number; admiteSupletorias: boolean; descripcion?: string; activo: boolean; orden: number };
type Habitacion = { id: string; residenciaId: string; residenciaNombre: string; tipoHabitacionId: string; tipoNombre: string; tipoCodigo: string; numero: string; nombre?: string; capacidadPersonas: number; admiteSupletorias: boolean; plazasSupletorias: number; activa: boolean; notas?: string; orden: number };
type Tarifa    = { id: string; residenciaId: string; residenciaNombre: string; tipoHabitacionId: string; tipoHabitacionNombre: string; nombreTarifa: string; precioNoche: number; precioMes: number; porcentajeIva: number; descripcion?: string; activa: boolean };
type Festivo   = { id: string; fecha: string; descripcion: string; ambito: string; activo: boolean };

type Page = 'dashboard' | 'residencias' | 'alojamientos' | 'tarifas' | 'festivos' | 'calendario' | 'reservas' | 'huespedes' | 'facturas' | 'bloqueos';

// ─── Login ─────────────────────────────────────────────────────────────────────
function LoginPage({ onLogin }: { onLogin: (token: string, user: User) => void }) {
  const [email, setEmail] = useState('admin@residencia.local');
  const [password, setPassword] = useState('Admin123!');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true); setError('');
    try {
      // 1. Obtener token (backend devuelve { token, email, fullName })
      const loginData = await apiFetch<{ token: string; email: string; fullName: string }>('/auth/login', {
        method: 'POST', body: JSON.stringify({ email, password })
      });
      localStorage.setItem('token', loginData.token);
      // 2. Obtener datos completos del usuario (con roles)
      const me = await apiFetch<User>('/users/me');
      onLogin(loginData.token, me);
    } catch { setError('Credenciales incorrectas. Inténtelo de nuevo.'); }
    finally { setLoading(false); }
  };

  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-logo">
          <div className="logo-icon">🏨</div>
          <h1>ResidenciaApp</h1>
          <p>Sistema de gestión de residencias</p>
        </div>
        {error && <div className="alert alert-danger">{error}</div>}
        <form onSubmit={handleSubmit}>
          <div className="form-group" style={{ marginBottom: 14 }}>
            <label>Correo electrónico</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} required />
          </div>
          <div className="form-group" style={{ marginBottom: 20 }}>
            <label>Contraseña</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} required />
          </div>
          <button type="submit" className="btn btn-primary" style={{ width: '100%', padding: 11, justifyContent: 'center' }} disabled={loading}>
            {loading ? <><span className="spinner"></span> Entrando...</> : '→ Iniciar sesión'}
          </button>
        </form>
      </div>
    </div>
  );
}

// ─── Sidebar ────────────────────────────────────────────────────────────────────
function Sidebar({ page, setPage, user, onLogout }: { page: Page; setPage: (p: Page) => void; user: User; onLogout: () => void }) {
  const isAdmin = user.roles.includes('Admin');
  const navItem = (p: Page, icon: string, label: string) => (
    <button className={`nav-item ${page === p ? 'active' : ''}`} onClick={() => setPage(p)}>
      <span className="icon">{icon}</span>{label}
    </button>
  );
  return (
    <div className="sidebar">
      <div className="sidebar-logo">
        <h2>🏨 ResidenciaApp</h2>
        <p>Panel de gestión</p>
      </div>
      <nav className="sidebar-nav">
        <div className="nav-section">Principal</div>
        {navItem('dashboard', '📊', 'Dashboard')}
        {navItem('huespedes', '👥', 'Solicitudes')}
        {navItem('calendario', '📅', 'Calendario')}
        {navItem('bloqueos', '🔒', 'Bloqueos')}
        {navItem('reservas', '📋', 'Reservas activas')}
        {navItem('facturas', '🧾', 'Facturación')}
        <div className="nav-section">Configuración</div>
        {navItem('residencias', '🏠', 'Residencias')}
        {navItem('alojamientos', '🔑', 'Alojamientos')}
        {navItem('tarifas', '💰', 'Tarifas y precios')}
        {navItem('festivos', '🎉', 'Festivos')}
      </nav>
      <div className="sidebar-footer">
        <div className="user-info">
          <div className="user-avatar">{user.fullName?.[0] ?? '?'}</div>
          <div>
            <div className="user-name">{user.fullName}</div>
            <div className="user-role">{isAdmin ? '👑 Administrador' : '👤 Recepcionista'}</div>
          </div>
        </div>
        <button className="btn-logout" onClick={onLogout}>⏻ Cerrar sesión</button>
      </div>
    </div>
  );
}

// ─── Dashboard ──────────────────────────────────────────────────────────────────
function Dashboard({ setPage }: { setPage: (p: Page) => void }) {
  const [stats, setStats] = useState({ residencias: 0, habitaciones: 0, tipos: 0, festivos: 0 });
  useEffect(() => {
    Promise.all([
      apiFetch<Residencia[]>('/api/residencias'),
      apiFetch<Habitacion[]>('/api/habitaciones'),
      apiFetch<TipoHab[]>('/api/tipos-habitacion'),
      apiFetch<Festivo[]>('/api/festivos'),
    ]).then(([r, h, t, f]) => setStats({ residencias: r.length, habitaciones: h.length, tipos: t.length, festivos: f.length })).catch(() => {});
  }, []);

  return (
    <>
      <div className="stats-row">
        <div className="stat-card" style={{ cursor: 'pointer' }} onClick={() => setPage('residencias')}>
          <div className="stat-icon">🏠</div>
          <div className="stat-value">{stats.residencias}</div>
          <div className="stat-label">Residencias</div>
        </div>
        <div className="stat-card" style={{ cursor: 'pointer' }} onClick={() => setPage('alojamientos')}>
          <div className="stat-icon">🔑</div>
          <div className="stat-value">{stats.habitaciones}</div>
          <div className="stat-label">Habitaciones activas</div>
        </div>
        <div className="stat-card" style={{ cursor: 'pointer' }} onClick={() => setPage('alojamientos')}>
          <div className="stat-icon">🛏️</div>
          <div className="stat-value">{stats.tipos}</div>
          <div className="stat-label">Tipos de habitación</div>
        </div>
        <div className="stat-card" style={{ cursor: 'pointer' }} onClick={() => setPage('festivos')}>
          <div className="stat-icon">🎉</div>
          <div className="stat-value">{stats.festivos}</div>
          <div className="stat-label">Festivos configurados</div>
        </div>
      </div>
      <div className="card">
        <div className="card-header"><h3>Accesos rápidos</h3></div>
        <div className="card-body">
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
            <button className="btn btn-primary" onClick={() => setPage('residencias')}>🏠 Gestionar residencias</button>
            <button className="btn btn-primary" onClick={() => setPage('alojamientos')}>🔑 Gestionar alojamientos</button>
            <button className="btn btn-accent" onClick={() => setPage('tarifas')}>💰 Configurar tarifas</button>
            <button className="btn btn-ghost" onClick={() => setPage('festivos')}>🎉 Añadir festivos</button>
          </div>
        </div>
      </div>
    </>
  );
}

// ─── Generic CRUD Table + Modal ─────────────────────────────────────────────────
function CrudPage<T extends { id: string }>({
  title, icon, fetchUrl,
  columns, renderRow,
  emptyLabel, FormComponent,
}: {
  title: string; icon: string; fetchUrl: string;
  columns: string[];
  renderRow: (item: T, onEdit: () => void, onDelete: () => void) => React.ReactNode;
  emptyLabel: string;
  FormComponent: React.FC<{ item?: T; onSaved: () => void; onCancel: () => void }>;
}) {
  const [items, setItems] = useState<T[]>([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [editing, setEditing] = useState<T | undefined>();
  const [alert, setAlert] = useState<{ type: 'success' | 'danger'; msg: string } | null>(null);

  const load = async () => { setLoading(true); try { setItems(await apiFetch<T[]>(fetchUrl)); } catch { } finally { setLoading(false); } };
  useEffect(() => { void load(); }, [fetchUrl]);

  const handleEdit = (item: T) => { setEditing(item); setShowModal(true); };
  const handleNew  = () => { setEditing(undefined); setShowModal(true); };
  const handleDelete = async (item: T) => {
    if (!confirm(`¿Eliminar este registro?`)) return;
    try { await apiFetch(`${fetchUrl}/${item.id}`, { method: 'DELETE' }); setAlert({ type: 'success', msg: 'Eliminado correctamente.' }); void load(); }
    catch (e: any) { setAlert({ type: 'danger', msg: e.message }); }
  };
  const handleSaved = () => { setShowModal(false); setEditing(undefined); setAlert({ type: 'success', msg: 'Guardado correctamente.' }); void load(); };

  return (
    <>
      {alert && <div className={`alert alert-${alert.type}`} onClick={() => setAlert(null)} style={{ cursor: 'pointer' }}>{alert.msg} ✕</div>}
      <div className="card">
        <div className="card-header">
          <h3>{icon} {title}</h3>
          <button className="btn btn-primary btn-sm" onClick={handleNew}>+ Nuevo</button>
        </div>
        <div className="table-container">
          {loading ? <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div> : items.length === 0 ? (
            <div className="empty-state"><div className="empty-icon">{icon}</div><p>{emptyLabel}</p><button className="btn btn-primary mt-4" onClick={handleNew}>+ Crear primero</button></div>
          ) : (
            <table>
              <thead><tr>{columns.map(c => <th key={c}>{c}</th>)}<th>Acciones</th></tr></thead>
              <tbody>{items.map(item => renderRow(item, () => handleEdit(item), () => handleDelete(item)))}</tbody>
            </table>
          )}
        </div>
      </div>
      {showModal && (
        <div className="modal-overlay" onClick={e => e.target === e.currentTarget && setShowModal(false)}>
          <div className="modal">
            <div className="modal-header">
              <h3>{editing ? 'Editar' : 'Nuevo'} — {title}</h3>
              <button className="modal-close" onClick={() => setShowModal(false)}>×</button>
            </div>
            <FormComponent item={editing} onSaved={handleSaved} onCancel={() => setShowModal(false)} />
          </div>
        </div>
      )}
    </>
  );
}

// ─── Residencias Page ───────────────────────────────────────────────────────────
function ResidenciaForm({ item, onSaved, onCancel }: { item?: Residencia; onSaved: () => void; onCancel: () => void }) {
  const [f, setF] = useState({ nombre: item?.nombre ?? '', razonSocial: item?.razonSocial ?? '', cif: item?.cif ?? '', direccion: item?.direccion ?? '', codigoPostal: item?.codigoPostal ?? '', municipio: item?.municipio ?? '', provincia: item?.provincia ?? '', telefono: item?.telefono ?? '', email: item?.email ?? '', activa: item?.activa ?? true, orden: item?.orden ?? 0 });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));
  const submit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      const url = item ? `/api/residencias/${item.id}` : '/api/residencias';
      await apiFetch(url, { method: item ? 'PUT' : 'POST', body: JSON.stringify(f) });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };
  return (
    <form onSubmit={submit}>
      <div className="modal-body">
        <div className="form-grid">
          <div className="form-group"><label>Nombre *</label><input value={f.nombre} onChange={e => set('nombre', e.target.value)} required /></div>
          <div className="form-group"><label>Razón social</label><input value={f.razonSocial} onChange={e => set('razonSocial', e.target.value)} /></div>
          <div className="form-group"><label>CIF/NIF</label><input value={f.cif} onChange={e => set('cif', e.target.value)} /></div>
          <div className="form-group"><label>Teléfono</label><input value={f.telefono} onChange={e => set('telefono', e.target.value)} /></div>
          <div className="form-group"><label>Email</label><input type="email" value={f.email} onChange={e => set('email', e.target.value)} /></div>
          <div className="form-group"><label>Código Postal</label><input value={f.codigoPostal} onChange={e => set('codigoPostal', e.target.value)} /></div>
          <div className="form-group"><label>Municipio</label><input value={f.municipio} onChange={e => set('municipio', e.target.value)} /></div>
          <div className="form-group"><label>Provincia</label><input value={f.provincia} onChange={e => set('provincia', e.target.value)} /></div>
          <div className="form-group"><label>Dirección completa</label><input value={f.direccion} onChange={e => set('direccion', e.target.value)} /></div>
          <div className="form-group"><label>Orden</label><input type="number" value={f.orden} onChange={e => set('orden', +e.target.value)} /></div>
        </div>
        <div style={{ marginTop: 14 }}>
          <label className="toggle-wrap"><span className="toggle"><input type="checkbox" checked={f.activa} onChange={e => set('activa', e.target.checked)} /><span className="toggle-slider"></span></span> Residencia activa</label>
        </div>
      </div>
      <div className="modal-footer">
        <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Guardando...' : '💾 Guardar'}</button>
      </div>
    </form>
  );
}

function ResidenciasPage() {
  return <CrudPage<Residencia> title="Residencias" icon="🏠" fetchUrl="/api/residencias"
    columns={['Nombre', 'Municipio', 'Teléfono', 'Habitaciones', 'Estado']}
    emptyLabel="No hay residencias configuradas"
    FormComponent={ResidenciaForm}
    renderRow={(r, onEdit, onDelete) => (
      <tr key={r.id}>
        <td><strong>{r.nombre}</strong>{r.razonSocial && <div className="text-muted" style={{ fontSize: 11 }}>{r.razonSocial}</div>}</td>
        <td>{r.municipio ?? '—'}{r.provincia && <span className="text-muted"> ({r.provincia})</span>}</td>
        <td>{r.telefono ?? '—'}</td>
        <td><span className="badge badge-primary">{r.totalHabitaciones} hab.</span></td>
        <td><span className={`badge ${r.activa ? 'badge-success' : 'badge-danger'}`}>{r.activa ? '✓ Activa' : '✗ Inactiva'}</span></td>
        <td><div className="actions-cell"><button className="btn btn-ghost btn-sm" onClick={onEdit}>✏️ Editar</button><button className="btn btn-danger btn-sm" onClick={onDelete}>🗑️</button></div></td>
      </tr>
    )} />;
}

// ─── Tipos Habitación Page ──────────────────────────────────────────────────────
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
    renderRow={(t, onEdit, onDelete) => (
      <tr key={t.id}>
        <td><strong>{t.nombre}</strong>{t.descripcion && <div className="text-muted" style={{ fontSize: 11 }}>{t.descripcion}</div>}</td>
        <td><span className="badge badge-primary">{t.codigo}</span></td>
        <td>{t.capacidadMaxima} persona{t.capacidadMaxima !== 1 ? 's' : ''}</td>
        <td>{t.admiteSupletorias ? <span className="badge badge-success">✓ Sí</span> : <span className="badge badge-danger">✗ No</span>}</td>
        <td><div className="actions-cell"><button className="btn btn-ghost btn-sm" onClick={onEdit}>✏️ Editar</button><button className="btn btn-danger btn-sm" onClick={onDelete}>🗑️</button></div></td>
      </tr>
    )} />;
}

// ─── Habitaciones Page ──────────────────────────────────────────────────────────
function HabitacionForm({ item, onSaved, onCancel }: { item?: Habitacion; onSaved: () => void; onCancel: () => void }) {
  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [tipos, setTipos] = useState<TipoHab[]>([]);
  const [f, setF] = useState({ residenciaId: item?.residenciaId ?? '', tipoHabitacionId: item?.tipoHabitacionId ?? '', numero: item?.numero ?? '', nombre: item?.nombre ?? '', capacidadPersonas: item?.capacidadPersonas ?? 1, admiteSupletorias: item?.admiteSupletorias ?? false, plazasSupletorias: item?.plazasSupletorias ?? 0, activa: item?.activa ?? true, notas: item?.notas ?? '', orden: item?.orden ?? 0 });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));

  useEffect(() => {
    Promise.all([apiFetch<Residencia[]>('/api/residencias'), apiFetch<TipoHab[]>('/api/tipos-habitacion')])
      .then(([r, t]) => { setResidencias(r); setTipos(t); });
  }, []);

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
              <option value="">— Seleccione —</option>
              {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
            </select>
          </div>
          <div className="form-group"><label>Tipo habitación *</label>
            <select value={f.tipoHabitacionId} onChange={e => { const t = tipos.find(x => x.id === e.target.value); set('tipoHabitacionId', e.target.value); if (t) set('capacidadPersonas', t.capacidadMaxima); }} required>
              <option value="">— Seleccione —</option>
              {tipos.map(t => <option key={t.id} value={t.id}>{t.nombre} ({t.codigo})</option>)}
            </select>
          </div>
          <div className="form-group"><label>Número / Identificador *</label><input value={f.numero} onChange={e => set('numero', e.target.value)} placeholder="1, OF.1, Ap.2, EST.1..." required /></div>
          <div className="form-group"><label>Nombre descriptivo</label><input value={f.nombre} onChange={e => set('nombre', e.target.value)} placeholder="Opcional" /></div>
          <div className="form-group"><label>Capacidad personas</label><input type="number" min={1} max={20} value={f.capacidadPersonas} onChange={e => set('capacidadPersonas', +e.target.value)} /></div>
          <div className="form-group"><label>Plazas supletorias</label><input type="number" min={0} max={5} value={f.plazasSupletorias} onChange={e => set('plazasSupletorias', +e.target.value)} /></div>
          <div className="form-group"><label>Orden en calendario</label><input type="number" value={f.orden} onChange={e => set('orden', +e.target.value)} /></div>
          <div className="form-group"><label>Notas</label><input value={f.notas} onChange={e => set('notas', e.target.value)} /></div>
        </div>
        <div style={{ marginTop: 14, display: 'flex', gap: 20 }}>
          <label className="toggle-wrap"><span className="toggle"><input type="checkbox" checked={f.admiteSupletorias} onChange={e => set('admiteSupletorias', e.target.checked)} /><span className="toggle-slider"></span></span> Admite supletorias</label>
          <label className="toggle-wrap"><span className="toggle"><input type="checkbox" checked={f.activa} onChange={e => set('activa', e.target.checked)} /><span className="toggle-slider"></span></span> Habitación activa</label>
        </div>
      </div>
      <div className="modal-footer">
        <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Guardando...' : '💾 Guardar'}</button>
      </div>
    </form>
  );
}

function HabitacionesPage() {
  const [items, setItems] = useState<Habitacion[]>([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [editing, setEditing] = useState<Habitacion | undefined>(undefined);
  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>('Todas');

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
    if (!confirm(`¿Eliminar esta habitación?`)) return;
    try {
      await apiFetch(`/api/habitaciones/${item.id}`, { method: 'DELETE' });
      void load();
    } catch (e: any) { alert(e.message); }
  };

  const handleEdit = (item: Habitacion) => { setEditing(item); setShowModal(true); };
  const handleNew  = () => { setEditing(undefined); setShowModal(true); };
  const handleSaved = () => { setShowModal(false); setEditing(undefined); void load(); };

  const filtradas = items.filter(h => residenciaActiva === 'Todas' || h.residenciaNombre === residenciaActiva);

  return (
    <>
      <div className="card">
        <div className="card-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 10 }}>
          <h3>🔑 Habitaciones por Residencia</h3>
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <label style={{ fontSize: 13, fontWeight: 600, color: 'var(--text-muted)' }}>Filtrar residencia:</label>
            <select value={residenciaActiva} onChange={e => setResidenciaActiva(e.target.value)}
              style={{ padding: '6px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              <option value="Todas">Todas las residencias</option>
              {residencias.map(r => (
                <option key={r.id} value={r.nombre}>{r.nombre}</option>
              ))}
            </select>
            <button className="btn btn-primary btn-sm" onClick={handleNew}>+ Nueva habitación</button>
          </div>
        </div>
        <div className="table-container">
          {loading ? (
            <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div>
          ) : filtradas.length === 0 ? (
            <div className="empty-state">
              <div className="empty-icon">🔑</div>
              <p>No hay habitaciones configuradas para esta residencia</p>
              <button className="btn btn-primary mt-4" onClick={handleNew}>+ Nueva habitación</button>
            </div>
          ) : (
            <table>
              <thead>
                <tr>
                  <th>Residencia</th>
                  <th>Nº / ID</th>
                  <th>Tipo</th>
                  <th>Capacidad</th>
                  <th>Estado</th>
                  <th style={{ width: 120 }}>Acciones</th>
                </tr>
              </thead>
              <tbody>
                {filtradas.map(h => (
                  <tr key={h.id}>
                    <td>{h.residenciaNombre}</td>
                    <td><strong>{h.numero}</strong>{h.nombre && <span className="text-muted"> — {h.nombre}</span>}</td>
                    <td><span className="badge badge-primary">{h.tipoCodigo}</span> {h.tipoNombre}</td>
                    <td>{h.capacidadPersonas}p{h.admiteSupletorias ? ` +${h.plazasSupletorias}sup` : ''}</td>
                    <td><span className={`badge ${h.activa ? 'badge-success' : 'badge-danger'}`}>{h.activa ? '✓ Activa' : '✗ Bloqueada'}</span></td>
                    <td>
                      <div className="actions-cell">
                        <button className="btn btn-ghost btn-sm" onClick={() => handleEdit(h)}>✏️ Editar</button>
                        <button className="btn btn-danger btn-sm" onClick={() => handleDelete(h)}>🗑️</button>
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
        <div className="modal-overlay" onClick={e => e.target === e.currentTarget && setShowModal(false)}>
          <div className="modal">
            <div className="modal-header">
              <h3>{editing ? '✏️ Editar Habitación' : '➕ Nueva Habitación'}</h3>
              <button className="modal-close" onClick={() => setShowModal(false)}>×</button>
            </div>
            <HabitacionForm item={editing} onSaved={handleSaved} onCancel={() => setShowModal(false)} />
          </div>
        </div>
      )}
    </>
  );
}

function AlojamientosPage() {
  const [tab, setTab] = useState<'tipos' | 'habitaciones'>('tipos');
  return (
    <>
      <div style={{ display: 'flex', gap: 10, marginBottom: 15, padding: '0 5px' }}>
        <button className={`btn ${tab === 'tipos' ? 'btn-primary' : 'btn-ghost'}`} style={{ borderRadius: 8, padding: '8px 16px', fontWeight: 600 }} onClick={() => setTab('tipos')}>
          🛏️ 1. Categorías / Tipos de Habitación
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

// ─── Tarifas Page ───────────────────────────────────────────────────────────────
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
              <option value="">— Seleccione —</option>
              {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
            </select>
          </div>
          <div className="form-group">
            <label>Tipo habitación *</label>
            <select value={f.tipoHabitacionId} onChange={e => set('tipoHabitacionId', e.target.value)} required>
              <option value="">— Seleccione —</option>
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

function TarifasPage() {
  const [tarifas, setTarifas] = useState<Tarifa[]>([]);
  const [residencias, setResidencias] = useState<any[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>('');
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState<Tarifa | undefined>(undefined);

  const cargarResidencias = async () => {
    try {
      const res = await apiFetch<any[]>('/api/residencias');
      setResidencias(res);
      if (res.length > 0 && !residenciaActiva) {
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
          <select value={residenciaActiva} onChange={e => setResidenciaActiva(e.target.value)}
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
              <button className="modal-close" onClick={() => setShowModal(false)}>×</button>
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

// ─── Festivos Page ──────────────────────────────────────────────────────────────
function FestivoForm({ item, onSaved, onCancel }: { item?: Festivo; onSaved: () => void; onCancel: () => void }) {
  const [f, setF] = useState({ fecha: item?.fecha ?? '', descripcion: item?.descripcion ?? '', ambito: item?.ambito ?? 'NACIONAL', activo: item?.activo ?? true });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));
  const submit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      const url = item ? `/api/festivos/${item.id}` : '/api/festivos';
      await apiFetch(url, { method: item ? 'PUT' : 'POST', body: JSON.stringify(f) });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };
  return (
    <form onSubmit={submit}>
      <div className="modal-body">
        <div className="form-grid">
          <div className="form-group"><label>Fecha *</label><input type="date" value={f.fecha} onChange={e => set('fecha', e.target.value)} required /></div>
          <div className="form-group"><label>Ámbito</label>
            <select value={f.ambito} onChange={e => set('ambito', e.target.value)}>
              <option value="NACIONAL">Nacional</option>
              <option value="REGIONAL">Regional (Asturias)</option>
              <option value="LOCAL">Local</option>
            </select>
          </div>
          <div className="form-group" style={{ gridColumn: '1 / -1' }}><label>Descripción *</label><input value={f.descripcion} onChange={e => set('descripcion', e.target.value)} placeholder="Ej: Navidad, Año Nuevo, Día de Asturias..." required /></div>
        </div>
      </div>
      <div className="modal-footer">
        <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
        <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Guardando...' : '💾 Guardar'}</button>
      </div>
    </form>
  );
}

function FestivosPage() {
  return <CrudPage<Festivo> title="Festivos" icon="🎉" fetchUrl="/api/festivos"
    columns={['Fecha', 'Descripción', 'Ámbito']}
    emptyLabel="No hay festivos configurados"
    FormComponent={FestivoForm}
    renderRow={(f, onEdit, onDelete) => (
      <tr key={f.id}>
        <td><strong>{new Date(f.fecha + 'T00:00:00').toLocaleDateString('es-ES', { weekday: 'short', day: 'numeric', month: 'long', year: 'numeric' })}</strong></td>
        <td>{f.descripcion}</td>
        <td><span className={`badge ${f.ambito === 'NACIONAL' ? 'badge-primary' : f.ambito === 'REGIONAL' ? 'badge-accent' : 'badge-success'}`}>{f.ambito}</span></td>
        <td><div className="actions-cell"><button className="btn btn-ghost btn-sm" onClick={onEdit}>✏️ Editar</button><button className="btn btn-danger btn-sm" onClick={onDelete}>🗑️</button></div></td>
      </tr>
    )} />;
}

// ─── Calendario Page (wrapper con modal nueva reserva) ─────────────────────────
function CalendarioPage() {
  const [showModal, setShowModal] = useState(false);
  const [habInicial, setHabInicial] = useState('');
  const [fechaInicial, setFechaInicial] = useState('');
  const [refreshKey, setRefreshKey] = useState(0);

  const handleNuevaReserva = (habitacionId: string, fecha: string) => {
    setHabInicial(habitacionId);
    setFechaInicial(fecha);
    setShowModal(true);
  };

  return (
    <>
      <div style={{ marginBottom: 12, display: 'flex', justifyContent: 'flex-end' }}>
        <button className="btn btn-primary" onClick={() => { setHabInicial(''); setFechaInicial(''); setShowModal(true); }}>
          + Nueva reserva
        </button>
      </div>
      <div className="card" style={{ padding: 0 }}>
        <CalendarioGantt key={refreshKey} onNuevaReserva={handleNuevaReserva} />
      </div>
      {showModal && (
        <NuevaReservaModal
          habitacionIdInicial={habInicial}
          fechaEntradaInicial={fechaInicial}
          onSaved={() => { setShowModal(false); setRefreshKey(k => k + 1); }}
          onCancel={() => setShowModal(false)}
        />
      )}
    </>
  );
}

// ─── Page Titles ────────────────────────────────────────────────────────────────
const PAGE_TITLES: Record<Page, { title: string; sub: string }> = {
  dashboard:    { title: 'Dashboard',             sub: 'Vista general del sistema' },
  calendario:   { title: 'Calendario',             sub: 'Vista Gantt de ocupación por habitación' },
  reservas:     { title: 'Reservas activas',       sub: 'Gestión de reservas de huéspedes activas' },
  bloqueos:     { title: 'Bloqueos',               sub: 'Gestión de bloqueos e inhabilitaciones de habitaciones/apartamentos' },
  huespedes:    { title: 'Solicitudes',            sub: 'Solicitudes de alojamiento registradas' },
  facturas:     { title: 'Facturación',             sub: 'Emisión y gestión de facturas' },
  residencias:  { title: 'Residencias',            sub: 'Gestión de centros y sedes' },
  alojamientos: { title: 'Alojamientos',           sub: 'Gestión de tipos de habitación y habitaciones por Residencia' },
  tarifas:      { title: 'Tarifas y precios',       sub: 'Configuración de precios e IVA' },
  festivos:     { title: 'Festivos',                sub: 'Calendario de días no laborables' },
};

// ─── App Root ───────────────────────────────────────────────────────────────────
export default function App() {
  const [token, setToken] = useState<string | null>(localStorage.getItem('token'));
  const [user, setUser]   = useState<User | null>(null);
  const [page, setPage]   = useState<Page>('dashboard');

  useEffect(() => {
    if (token && !user) {
      apiFetch<User>('/users/me').then(setUser).catch(() => { localStorage.removeItem('token'); setToken(null); });
    }
  }, [token]);

  const handleLogin = (t: string, u: User) => { setToken(t); setUser(u); };
  const handleLogout = () => { localStorage.removeItem('token'); setToken(null); setUser(null); };

  if (!token || !user) return <LoginPage onLogin={handleLogin} />;

  const pt = PAGE_TITLES[page];
  return (
    <div className="app-layout">
      <Sidebar page={page} setPage={setPage} user={user} onLogout={handleLogout} />
      <div className="main-content">
        <div className="topbar">
          <div><div className="topbar-title">{pt.title}</div><div className="topbar-subtitle">{pt.sub}</div></div>
        </div>
        <div className="page-content">
          {page === 'dashboard'    && <Dashboard setPage={setPage} />}
          {page === 'calendario'   && <CalendarioPage />}
          {page === 'reservas'     && <ReservasPage setPage={setPage} />}
          {page === 'huespedes'    && <HuespedesPage />}
          {page === 'facturas'     && <FacturasPage />}
          {page === 'residencias'  && <ResidenciasPage />}
          {page === 'alojamientos' && <AlojamientosPage />}
          {page === 'tarifas'      && <TarifasPage />}
          {page === 'festivos'     && <FestivosPage />}
          {page === 'bloqueos'     && <BloqueosPage />}
        </div>
      </div>
    </div>
  );
}
