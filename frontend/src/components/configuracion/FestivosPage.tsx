import { useState, FormEvent } from 'react';
import { apiFetch } from '../../api';
import { Festivo } from '../../types';
import { CrudPage } from '../shared/CrudPage';

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

export function FestivosPage() {
  return <CrudPage<Festivo> title="Festivos" icon="🎉" fetchUrl="/api/festivos"
    columns={['Fecha', 'Descripción', 'Ámbito']}
    emptyLabel="No hay festivos configurados"
    FormComponent={FestivoForm}
    renderRow={(f: Festivo, onEdit: () => void, onDelete: () => void) => (
      <tr key={f.id}>
        <td><strong>{new Date(f.fecha + 'T00:00:00').toLocaleDateString('es-ES', { weekday: 'short', day: 'numeric', month: 'long', year: 'numeric' })}</strong></td>
        <td>{f.descripcion}</td>
        <td><span className={`badge ${f.ambito === 'NACIONAL' ? 'badge-primary' : f.ambito === 'REGIONAL' ? 'badge-accent' : 'badge-success'}`}>{f.ambito}</span></td>
        <td><div className="actions-cell"><button className="btn btn-ghost btn-sm" onClick={onEdit}>✏️ Editar</button><button className="btn btn-danger btn-sm" onClick={onDelete}>🗑️</button></div></td>
      </tr>
    )} />;
}