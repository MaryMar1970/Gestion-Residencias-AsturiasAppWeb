import { useState, useEffect } from 'react';
import React from 'react';
import { apiFetch } from '../../api';

export function CrudPage<T extends { id: string }>({
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