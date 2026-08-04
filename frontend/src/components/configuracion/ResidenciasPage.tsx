import { useState, FormEvent } from 'react';
import { apiFetch } from '../../api';
import { Residencia } from '../../types';
import { CrudPage } from '../shared/CrudPage';

function ResidenciaForm({ item, onSaved, onCancel }: { item?: Residencia; onSaved: () => void; onCancel: () => void }) {
  const [f, setF] = useState({ nombre: item?.nombre ?? '', razonSocial: item?.razonSocial ?? '', cif: item?.cif ?? '', direccion: item?.direccion ?? '', codigoPostal: item?.codigoPostal ?? '', municipio: item?.municipio ?? '', provincia: item?.provincia ?? '', telefono: item?.telefono ?? '', email: item?.email ?? '', activa: item?.activa ?? true, orden: item?.orden ?? 0 });
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));
  const submit = async (e: FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      const payload = {
        ...f,
        nombre: f.nombre.trim(),
        razonSocial: f.razonSocial?.trim() || null,
        cif: f.cif?.trim() || null,
        direccion: f.direccion?.trim() || null,
        codigoPostal: f.codigoPostal?.trim() || null,
        municipio: f.municipio?.trim() || null,
        provincia: f.provincia?.trim() || null,
        telefono: f.telefono?.trim() || null,
        email: f.email?.trim() || null,
      };
      const url = item ? `/api/residencias/${item.id}` : '/api/residencias';
      await apiFetch(url, { method: item ? 'PUT' : 'POST', body: JSON.stringify(payload) });
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

export function ResidenciasPage() {
  return <CrudPage<Residencia> title="Residencias" icon="🏠" fetchUrl="/api/residencias"
    columns={['Nombre', 'Municipio', 'Teléfono', 'Habitaciones', 'Estado']}
    emptyLabel="No hay residencias configuradas"
    FormComponent={ResidenciaForm}
    renderRow={(r: Residencia, onEdit: () => void, onDelete: () => void) => (
      <tr key={r.id}>
        <td><strong>{r.nombre}</strong>{r.razonSocial && <div className="text-muted" style={{ fontSize: 11 }}>{r.razonSocial}</div>}</td>
        <td>{r.municipio ?? '—'}{r.provincia && <span className="text-muted"> ({r.provincia})</span>}</td>
        <td>{r.telefono ?? '—'}</td>
        <td><span className="badge badge-primary">{r.totalHabitaciones} hab.</span></td>
        <td><span className={`badge ${r.activa ? 'badge-success' : 'badge-danger'}`}>{r.activa ? '✓ Activa' : '✕ Inactiva'}</span></td>
        <td><div className="actions-cell"><button className="btn btn-ghost btn-sm" onClick={onEdit}>✏️ Editar</button><button className="btn btn-danger btn-sm" onClick={onDelete}>🗑️</button></div></td>
      </tr>
    )} />;
}