import { useState, useEffect, useMemo } from 'react';
import { apiFetch } from '../../api';
import { setMatrizEvaluacionCache, matrizPorDefecto } from '../../utils/evaluacionUtils';

interface ReglaItem {
  id: string; // key original
  finalidad: string;
  empleo: string;
  situacion: string;
  evaluacion: string;
}

type SortField = 'finalidad' | 'empleo' | 'situacion' | 'evaluacion';
type SortOrder = 'asc' | 'desc';

function compareEvaluaciones(a: string, b: string): number {
  const partsA = a.split(/[,.\s]+/).map(x => parseInt(x, 10) || 0);
  const partsB = b.split(/[,.\s]+/).map(x => parseInt(x, 10) || 0);
  const len = Math.max(partsA.length, partsB.length);
  for (let i = 0; i < len; i++) {
    const numA = partsA[i] ?? 0;
    const numB = partsB[i] ?? 0;
    if (numA !== numB) return numA - numB;
  }
  return a.localeCompare(b);
}

export function EvaluacionSolicitudesPage() {
  const [matriz, setMatriz] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [filterText, setFilterText] = useState('');
  const [successMsg, setSuccessMsg] = useState<string | null>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  // Estado de Ordenación (Por defecto: por Clasificación/Evaluación de menor a mayor)
  const [sortField, setSortField] = useState<SortField>('evaluacion');
  const [sortOrder, setSortOrder] = useState<SortOrder>('asc');

  // Modal para Añadir / Editar
  const [modalOpen, setModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState<ReglaItem | null>(null);
  const [formData, setFormData] = useState({
    finalidad: '',
    empleo: '',
    situacion: '',
    evaluacion: '',
  });

  const cargarDatos = async () => {
    setLoading(true);
    setErrorMsg(null);
    try {
      const data = await apiFetch<Record<string, string>>('/api/configuracion/matriz-evaluacion');
      if (data && typeof data === 'object') {
        setMatriz(data);
        setMatrizEvaluacionCache(data);
      } else {
        setMatriz(matrizPorDefecto);
      }
    } catch (err: any) {
      console.error('Error al cargar la matriz de evaluación:', err);
      setErrorMsg('No se pudo cargar la matriz desde el servidor. Mostrando matriz por defecto.');
      setMatriz(matrizPorDefecto);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void cargarDatos();
  }, []);

  const itemsList: ReglaItem[] = useMemo(() => {
    return Object.entries(matriz).map(([key, val]) => {
      const parts = key.split('|');
      return {
        id: key,
        finalidad: parts[0] ? parts[0].toUpperCase() : 'OTROS',
        empleo: parts[1] ? parts[1].toUpperCase() : 'GC',
        situacion: parts[2] ? parts[2].toUpperCase() : 'ACTIVO',
        evaluacion: val,
      };
    });
  }, [matriz]);

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortOrder(prev => (prev === 'asc' ? 'desc' : 'asc'));
    } else {
      setSortField(field);
      setSortOrder('asc');
    }
  };

  const itemsProcesados = useMemo(() => {
    let list = itemsList;
    if (filterText.trim()) {
      const q = filterText.toLowerCase();
      list = list.filter(
        item =>
          item.finalidad.toLowerCase().includes(q) ||
          item.empleo.toLowerCase().includes(q) ||
          item.situacion.toLowerCase().includes(q) ||
          item.evaluacion.toLowerCase().includes(q)
      );
    }

    return [...list].sort((a, b) => {
      let res = 0;
      if (sortField === 'evaluacion') {
        res = compareEvaluaciones(a.evaluacion, b.evaluacion);
      } else {
        res = a[sortField].localeCompare(b[sortField], 'es', { sensitivity: 'base' });
      }
      return sortOrder === 'asc' ? res : -res;
    });
  }, [itemsList, filterText, sortField, sortOrder]);

  const handleOpenAdd = () => {
    setEditingItem(null);
    setFormData({ finalidad: 'Comisión', empleo: 'GC', situacion: 'Activo', evaluacion: '1, 1, 5' });
    setModalOpen(true);
  };

  const handleOpenEdit = (item: ReglaItem) => {
    setEditingItem(item);
    setFormData({
      finalidad: item.finalidad,
      empleo: item.empleo,
      situacion: item.situacion,
      evaluacion: item.evaluacion,
    });
    setModalOpen(true);
  };

  const handleDelete = (item: ReglaItem) => {
    if (!confirm(`¿Eliminar la regla para ${item.finalidad} | ${item.empleo} | ${item.situacion}?`)) return;
    const newMatriz = { ...matriz };
    delete newMatriz[item.id];
    setMatriz(newMatriz);
  };

  const handleSaveModal = (e: React.FormEvent) => {
    e.preventDefault();
    const newKey = `${formData.finalidad.trim().toLowerCase()}|${formData.empleo.trim().toLowerCase()}|${formData.situacion.trim().toLowerCase()}`;
    if (!newKey) return;

    const newMatriz = { ...matriz };
    if (editingItem && editingItem.id !== newKey) {
      delete newMatriz[editingItem.id];
    }
    newMatriz[newKey] = formData.evaluacion.trim();

    setMatriz(newMatriz);
    setModalOpen(false);
  };

  const handleGuardarCambiosServidor = async () => {
    setSaving(true);
    setSuccessMsg(null);
    setErrorMsg(null);
    try {
      const updated = await apiFetch<Record<string, string>>('/api/configuracion/matriz-evaluacion', {
        method: 'PUT',
        body: JSON.stringify(matriz),
      });
      setMatriz(updated);
      setMatrizEvaluacionCache(updated);
      setSuccessMsg('Matriz de evaluación guardada con éxito en el servidor.');
      setTimeout(() => setSuccessMsg(null), 4000);
    } catch (err: any) {
      console.error('Error al guardar la matriz:', err);
      setErrorMsg(err.message || 'Error al guardar la matriz en el servidor.');
    } finally {
      setSaving(false);
    }
  };

  const handleRestablecerPorDefecto = () => {
    if (!confirm('¿Seguro que deseas restablecer la matriz a los valores por defecto del sistema?')) return;
    setMatriz(matrizPorDefecto);
  };

  const renderTh = (field: SortField, label: string) => {
    const isActive = sortField === field;
    return (
      <th
        onClick={() => handleSort(field)}
        style={{
          padding: '6px 10px',
          fontSize: 12,
          cursor: 'pointer',
          userSelect: 'none',
          whiteSpace: 'nowrap',
        }}
        title={`Pulsar para ordenar por ${label}`}
      >
        <div style={{ display: 'inline-flex', alignItems: 'center', gap: 5 }}>
          <span>{label}</span>
          <span style={{ fontSize: 10, color: isActive ? 'var(--primary)' : 'var(--text-muted)', opacity: isActive ? 1 : 0.4 }}>
            {isActive ? (sortOrder === 'asc' ? '▲' : '▼') : '↕'}
          </span>
        </div>
      </th>
    );
  };

  return (
    <div>
      {/* BARRA SUPERIOR DE ACCIONES Y BÚSQUEDA */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <input
            type="text"
            className="search-input"
            placeholder="🔍 Buscar regla..."
            value={filterText}
            onChange={e => setFilterText(e.target.value)}
            style={{ minWidth: 260, padding: '5px 10px', fontSize: 13 }}
          />
          <span style={{ fontSize: 12, color: 'var(--text-muted)' }}>
            Total: <strong>{itemsProcesados.length}</strong> de <strong>{itemsList.length}</strong> reglas
          </span>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button className="btn btn-ghost btn-sm" onClick={handleRestablecerPorDefecto} title="Restaurar matriz por defecto">
            🔄 Valores por defecto
          </button>
          <button className="btn btn-secondary btn-sm" onClick={handleOpenAdd}>
            + Nueva regla
          </button>
          <button className="btn btn-primary btn-sm" onClick={handleGuardarCambiosServidor} disabled={saving}>
            {saving ? '💾 Guardando...' : '💾 Guardar cambios'}
          </button>
        </div>
      </div>

      {successMsg && <div className="alert alert-success" style={{ marginBottom: 10, padding: '6px 12px' }}>✅ {successMsg}</div>}
      {errorMsg && <div className="alert alert-danger" style={{ marginBottom: 10, padding: '6px 12px' }}>⚠️ {errorMsg}</div>}

      {/* TABLA COMPACTA CON COLUMNAS ORDENABLES */}
      <div className="card" style={{ padding: 0 }}>
        {loading ? (
          <div style={{ padding: 20, textAlign: 'center', color: 'var(--text-muted)' }}>Cargando tabla de evaluación...</div>
        ) : itemsProcesados.length === 0 ? (
          <div style={{ padding: 20, textAlign: 'center', color: 'var(--text-muted)' }}>No se encontraron reglas de evaluación.</div>
        ) : (
          <div className="table-responsive" style={{ maxHeight: 'calc(100vh - 210px)', overflowY: 'auto' }}>
            <table className="table" style={{ width: '100%', fontSize: 12.5, borderCollapse: 'collapse' }}>
              <thead style={{ position: 'sticky', top: 0, background: 'var(--surface)', zIndex: 2 }}>
                <tr>
                  {renderTh('finalidad', 'FINALIDAD')}
                  {renderTh('empleo', 'EMPLEO')}
                  {renderTh('situacion', 'SITUACIÓN ADMINISTRATIVA')}
                  {renderTh('evaluacion', 'CLASIFICACIÓN / EVALUACIÓN')}
                  <th style={{ padding: '6px 10px', fontSize: 12, textAlign: 'right' }}>ACCIONES</th>
                </tr>
              </thead>
              <tbody>
                {itemsProcesados.map((item) => (
                  <tr key={item.id} style={{ height: 32 }}>
                    <td style={{ padding: '4px 10px' }}>
                      <strong style={{ color: 'var(--text-main)' }}>{item.finalidad}</strong>
                    </td>
                    <td style={{ padding: '4px 10px' }}>{item.empleo}</td>
                    <td style={{ padding: '4px 10px' }}>{item.situacion}</td>
                    <td style={{ padding: '4px 10px' }}>
                      <span className="badge badge-primary" style={{ fontSize: 11, padding: '2px 6px', fontWeight: 600 }}>
                        {item.evaluacion}
                      </span>
                    </td>
                    <td style={{ padding: '4px 10px', textAlign: 'right' }}>
                      <div className="actions-cell" style={{ justifyContent: 'flex-end', gap: 4 }}>
                        <button
                          className="btn btn-ghost btn-sm"
                          onClick={() => handleOpenEdit(item)}
                          style={{ padding: '2px 6px', fontSize: 11.5 }}
                        >
                          ✏️ Editar
                        </button>
                        <button
                          className="btn btn-danger btn-sm"
                          onClick={() => handleDelete(item)}
                          style={{ padding: '2px 6px', fontSize: 11.5 }}
                        >
                          🗑️
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* MODAL EDITAR / NUEVA REGLA */}
      {modalOpen && (
        <div className="modal-overlay" onClick={e => e.target === e.currentTarget && setModalOpen(false)}>
          <div className="modal" style={{ maxWidth: 460 }}>
            <div className="modal-header" style={{ padding: '12px 18px' }}>
              <h3 style={{ fontSize: 15, margin: 0 }}>
                {editingItem ? '✏️ Editar Regla de Evaluación' : '➕ Nueva Regla de Evaluación'}
              </h3>
              <button className="modal-close" onClick={() => setModalOpen(false)}>✕</button>
            </div>
            <form onSubmit={handleSaveModal}>
              <div className="modal-body" style={{ padding: '16px 18px', display: 'flex', flexDirection: 'column', gap: 12 }}>
                <div className="form-group">
                  <label style={{ fontSize: 12, marginBottom: 3 }}>Finalidad *</label>
                  <input
                    type="text"
                    value={formData.finalidad}
                    onChange={e => setFormData({ ...formData, finalidad: e.target.value })}
                    placeholder="Ej: Comisión, Destino, Enfermedad, Otros..."
                    style={{ padding: '6px 10px', fontSize: 13 }}
                    required
                  />
                </div>
                <div className="form-group">
                  <label style={{ fontSize: 12, marginBottom: 3 }}>Empleo *</label>
                  <input
                    type="text"
                    value={formData.empleo}
                    onChange={e => setFormData({ ...formData, empleo: e.target.value })}
                    placeholder="Ej: GC, Alumno GC, Militar en GC, Funcionario en GC..."
                    style={{ padding: '6px 10px', fontSize: 13 }}
                    required
                  />
                </div>
                <div className="form-group">
                  <label style={{ fontSize: 12, marginBottom: 3 }}>Situación Administrativa *</label>
                  <input
                    type="text"
                    value={formData.situacion}
                    onChange={e => setFormData({ ...formData, situacion: e.target.value })}
                    placeholder="Ej: Activo, Reserva, Retirado, Viogen, Excedencia..."
                    style={{ padding: '6px 10px', fontSize: 13 }}
                    required
                  />
                </div>
                <div className="form-group">
                  <label style={{ fontSize: 12, marginBottom: 3 }}>Clasificación / Evaluación (Puntuación) *</label>
                  <input
                    type="text"
                    value={formData.evaluacion}
                    onChange={e => setFormData({ ...formData, evaluacion: e.target.value })}
                    placeholder="Ej: 1, 1, 2"
                    style={{ padding: '6px 10px', fontSize: 13 }}
                    required
                  />
                </div>
              </div>
              <div className="modal-footer" style={{ padding: '10px 18px' }}>
                <button type="button" className="btn btn-ghost btn-sm" onClick={() => setModalOpen(false)}>
                  Cancelar
                </button>
                <button type="submit" className="btn btn-primary btn-sm">
                  {editingItem ? 'Guardar regla' : 'Añadir regla'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
