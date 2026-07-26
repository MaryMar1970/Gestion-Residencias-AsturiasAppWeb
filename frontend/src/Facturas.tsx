import { useState, useEffect, useCallback } from 'react';
import { apiFetch } from './api';

// ─── Types ─────────────────────────────────────────────────────────────────────
type Residencia = { id: string; nombre: string };
type LineaDto   = { id: string; orden: number; concepto: string; cantidad: number; unidad: string; precioUnidad: number; descuento: number; baseLinea: number; porcentajeIva: number; cuotaIvaLinea: number; totalLinea: number };
type Factura = {
  id: string; numeroFactura: string; serie: string; ejercicio: number; numeroOrden: number;
  residenciaId: string; residenciaNombre: string;
  reservaId?: string; huespedId?: string;
  fechaEmision: string; fechaVencimiento?: string; fechaPago?: string;
  destinatarioNombre: string; destinatarioDni: string;
  destinatarioDireccion?: string; destinatarioCp?: string; destinatarioMunicipio?: string;
  emisorNombre: string; emisorCif?: string; emisorDireccion?: string; emisorTelefono?: string;
  baseImponible: number; porcentajeIva: number; cuotaIva: number; total: number;
  formaPago?: string; estado: string; estadoInt: number;
  observaciones?: string; lineas: LineaDto[];
};
type Reserva = {
  id: string; habitacionNumero: string; residenciaNombre: string;
  huespedNombreCompleto?: string; huespedDni?: string;
  fechaEntrada: string; fechaSalida: string; totalNoches: number;
  importeTotal: number; estado: string; facturado: boolean; pagado: boolean;
  esBloqueo: boolean;
};

const API_BASE = 'http://localhost:5260';

const ESTADO_CLASE: Record<string, string> = {
  Borrador: 'badge-warning', Emitida: 'badge-info',
  Pagada: 'badge-success',   Anulada: 'badge-danger'
};

// ─── PÁGINA PRINCIPAL DE FACTURAS ─────────────────────────────────────────────
export function FacturasPage() {
  const [facturas, setFacturas]           = useState<Factura[]>([]);
  const [residencias, setResidencias]     = useState<Residencia[]>([]);
  const [filtroRes, setFiltroRes]         = useState('');
  const [filtroAnio, setFiltroAnio]       = useState(String(new Date().getFullYear()));
  const [filtroEstado, setFiltroEstado]   = useState('');
  const [loading, setLoading]             = useState(true);
  const [alert, setAlert]                 = useState('');
  const [alertType, setAlertType]         = useState<'success'|'danger'>('success');
  const [vistaFactura, setVistaFactura]   = useState<Factura | null>(null);
  const [showDesdeReserva, setShowDesdeReserva] = useState(false);
  const [showManual, setShowManual]       = useState(false);

  const cargar = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (filtroRes)    params.set('residenciaId', filtroRes);
      if (filtroAnio)   params.set('ejercicio', filtroAnio);
      if (filtroEstado) params.set('estado', filtroEstado);
      setFacturas(await apiFetch<Factura[]>(`/api/facturas?${params}`));
    } catch { } finally { setLoading(false); }
  }, [filtroRes, filtroAnio, filtroEstado]);

  useEffect(() => { apiFetch<Residencia[]>('/api/residencias').then(setResidencias).catch(() => {}); }, []);
  useEffect(() => { void cargar(); }, [cargar]);

  const showMsg = (msg: string, t: 'success'|'danger' = 'success') => {
    setAlert(msg); setAlertType(t);
    setTimeout(() => setAlert(''), 4000);
  };

  const abrirHtml = async (id: string) => {
    try {
      const token = localStorage.getItem('token') ?? '';
      const res = await fetch(`${API_BASE}/api/facturas/${id}/html`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      const html = await res.text();
      const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
      const url  = URL.createObjectURL(blob);
      const win  = window.open(url, '_blank');
      // Liberar el objeto URL después de que se abra la pestaña
      if (win) setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch { showMsg('Error al generar la vista de impresión.', 'danger'); }
  };

  const cambiarEstado = async (f: Factura, accion: 'emitir'|'pagar'|'anular') => {
    const msgs: Record<string, string> = {
      emitir: '¿Emitir esta factura? Pasará a estado EMITIDA.',
      pagar:  '¿Marcar como PAGADA? Indica la forma de pago:',
      anular: '¿ANULAR esta factura? Esta acción es irreversible.'
    };
    if (!confirm(msgs[accion])) return;
    try {
      if (accion === 'emitir') await apiFetch(`/api/facturas/${f.id}/emitir`, { method: 'POST', body: JSON.stringify({ formaPago: null }) });
      if (accion === 'pagar')  {
        const forma = prompt('Forma de pago:', 'Transferencia');
        await apiFetch(`/api/facturas/${f.id}/pagar?formaPago=${encodeURIComponent(forma ?? '')}`, { method: 'POST' });
      }
      if (accion === 'anular') await apiFetch(`/api/facturas/${f.id}/anular`, { method: 'POST', body: JSON.stringify({ motivo: '' }) });
      showMsg(`Factura ${f.numeroFactura} actualizada.`);
      await cargar();
    } catch (e: any) { showMsg(e.message, 'danger'); }
  };

  const anios = Array.from({ length: 4 }, (_, i) => String(new Date().getFullYear() - i));

  return (
    <>
      {alert && <div className={`alert alert-${alertType}`} style={{ marginBottom: 12 }}>{alert}</div>}

      {/* Estadísticas rápidas */}
      <div className="stats-row" style={{ marginBottom: 18 }}>
        {[
          { label: 'Emitidas', val: facturas.filter(f => f.estado === 'Emitida').length, color: 'var(--primary-light)', icon: '📄' },
          { label: 'Pagadas',  val: facturas.filter(f => f.estado === 'Pagada').length,  color: 'var(--success)', icon: '✅' },
          { label: 'Borradores', val: facturas.filter(f => f.estado === 'Borrador').length, color: 'var(--accent)', icon: '📝' },
          { label: 'Total facturado', val: `${facturas.filter(f => f.estado !== 'Anulada').reduce((s,f)=>s+f.total,0).toFixed(2)} €`, color: 'var(--primary)', icon: '💶' },
        ].map(s => (
          <div key={s.label} className="stat-card" style={{ borderTop: `3px solid ${s.color}` }}>
            <div className="stat-icon">{s.icon}</div>
            <div className="stat-value" style={{ color: s.color }}>{s.val}</div>
            <div className="stat-label">{s.label}</div>
          </div>
        ))}
      </div>

      <div className="card">
        <div className="card-header">
          <h3>🧾 Facturas</h3>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
            <select value={filtroRes} onChange={e => setFiltroRes(e.target.value)} style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              <option value="">Todas las residencias</option>
              {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
            </select>
            <select value={filtroAnio} onChange={e => setFiltroAnio(e.target.value)} style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              {anios.map(a => <option key={a} value={a}>{a}</option>)}
            </select>
            <select value={filtroEstado} onChange={e => setFiltroEstado(e.target.value)} style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
              <option value="">Todos los estados</option>
              <option value="Borrador">Borrador</option>
              <option value="Emitida">Emitida</option>
              <option value="Pagada">Pagada</option>
              <option value="Anulada">Anulada</option>
            </select>
            <button className="btn btn-ghost btn-sm" onClick={() => setShowDesdeReserva(true)}>📋 Desde reserva</button>
            <button className="btn btn-primary btn-sm" onClick={() => setShowManual(true)}>+ Nueva manual</button>
          </div>
        </div>

        <div className="table-container">
          {loading ? (
            <div style={{ padding: 30, textAlign: 'center' }}><span className="spinner"></span></div>
          ) : facturas.length === 0 ? (
            <div className="empty-state">
              <div className="empty-icon">🧾</div>
              <p>No hay facturas {filtroEstado ? `en estado "${filtroEstado}"` : 'para este período'}</p>
              <div style={{ display: 'flex', gap: 10, justifyContent: 'center', marginTop: 12 }}>
                <button className="btn btn-ghost" onClick={() => setShowDesdeReserva(true)}>Facturar una reserva</button>
                <button className="btn btn-primary" onClick={() => setShowManual(true)}>Crear manual</button>
              </div>
            </div>
          ) : (
            <table>
              <thead>
                <tr>
                  <th>Nº Factura</th>
                  <th>Fecha</th>
                  <th>Destinatario</th>
                  <th>Residencia</th>
                  <th>Base</th>
                  <th>IVA</th>
                  <th style={{ fontWeight: 700 }}>Total</th>
                  <th>Estado</th>
                  <th>Acciones</th>
                </tr>
              </thead>
              <tbody>
                {facturas.map(f => (
                  <tr key={f.id} style={{ opacity: f.estado === 'Anulada' ? 0.5 : 1 }}>
                    <td><strong style={{ fontFamily: 'monospace', color: 'var(--primary)' }}>{f.numeroFactura}</strong></td>
                    <td>{f.fechaEmision}</td>
                    <td>
                      <strong>{f.destinatarioNombre}</strong>
                      {f.destinatarioDni && <div style={{ fontSize: 11, color: 'var(--text-muted)' }}>{f.destinatarioDni}</div>}
                    </td>
                    <td>{f.residenciaNombre}</td>
                    <td style={{ textAlign: 'right' }}>{f.baseImponible.toFixed(2)} €</td>
                    <td style={{ textAlign: 'right' }}>{f.cuotaIva.toFixed(2)} €</td>
                    <td style={{ textAlign: 'right', fontWeight: 700 }}>{f.total.toFixed(2)} €</td>
                    <td>
                      <span className={`badge ${ESTADO_CLASE[f.estado] ?? 'badge-primary'}`}>{f.estado}</span>
                      {f.formaPago && <div style={{ fontSize: 10, color: 'var(--text-muted)', marginTop: 2 }}>{f.formaPago}</div>}
                    </td>
                    <td>
                      <div className="actions-cell">
                        <button className="btn btn-ghost btn-sm" onClick={() => setVistaFactura(f)} title="Ver detalle">👁️</button>
                        <button className="btn btn-ghost btn-sm" onClick={() => abrirHtml(f.id)} title="Imprimir / PDF">🖨️</button>
                        {f.estado === 'Borrador' && <button className="btn btn-ghost btn-sm" onClick={() => cambiarEstado(f, 'emitir')} title="Emitir">📤</button>}
                        {(f.estado === 'Emitida' || f.estado === 'Borrador') && <button className="btn btn-ghost btn-sm" onClick={() => cambiarEstado(f, 'pagar')} title="Marcar pagada">💰</button>}
                        {f.estado !== 'Anulada' && <button className="btn btn-danger btn-sm" onClick={() => cambiarEstado(f, 'anular')} title="Anular">✗</button>}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

      {/* Modal vista previa */}
      {vistaFactura && <FacturaDetalleModal factura={vistaFactura} onClose={() => { setVistaFactura(null); void cargar(); }} onImprimir={() => abrirHtml(vistaFactura.id)} />}

      {/* Modal desde reserva */}
      {showDesdeReserva && <FacturarReservaModal onSaved={() => { setShowDesdeReserva(false); showMsg('Factura generada correctamente.'); void cargar(); }} onCancel={() => setShowDesdeReserva(false)} />}

      {/* Modal manual */}
      {showManual && <FacturaManualModal residencias={residencias} onSaved={() => { setShowManual(false); showMsg('Factura creada correctamente.'); void cargar(); }} onCancel={() => setShowManual(false)} />}
    </>
  );
}

// ─── MODAL DETALLE FACTURA ─────────────────────────────────────────────────────
function FacturaDetalleModal({ factura: f, onClose, onImprimir }: { factura: Factura; onClose: () => void; onImprimir: () => void }) {
  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 720 }}>
        <div className="modal-header">
          <h3>🧾 {f.numeroFactura} — <span className={`badge ${ESTADO_CLASE[f.estado]}`}>{f.estado}</span></h3>
          <button className="modal-close" onClick={onClose}>×</button>
        </div>
        <div className="modal-body">
          {/* Cabecera emisor / destinatario */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 16, padding: '12px 14px', background: 'var(--surface-2)', borderRadius: 8 }}>
            <div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', color: 'var(--text-muted)', marginBottom: 4 }}>Emisor</div>
              <strong>{f.emisorNombre}</strong><br />
              <span style={{ fontSize: 12, color: 'var(--text-muted)' }}>{f.emisorCif} · {f.emisorDireccion}</span>
            </div>
            <div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', color: 'var(--text-muted)', marginBottom: 4 }}>Destinatario</div>
              <strong>{f.destinatarioNombre}</strong><br />
              <span style={{ fontSize: 12, color: 'var(--text-muted)' }}>{f.destinatarioDni} · {f.destinatarioMunicipio}</span>
            </div>
          </div>

          {/* Info fechas y pago */}
          <div style={{ display: 'flex', gap: 20, flexWrap: 'wrap', marginBottom: 16, fontSize: 13 }}>
            <div><span style={{ color: 'var(--text-muted)' }}>Emisión:</span> <strong>{f.fechaEmision}</strong></div>
            {f.fechaVencimiento && <div><span style={{ color: 'var(--text-muted)' }}>Vencimiento:</span> <strong>{f.fechaVencimiento}</strong></div>}
            {f.formaPago && <div><span style={{ color: 'var(--text-muted)' }}>Pago:</span> <strong>{f.formaPago}</strong></div>}
            {f.fechaPago && <div><span style={{ color: 'var(--text-muted)' }}>Fecha pago:</span> <strong>{f.fechaPago}</strong></div>}
          </div>

          {/* Líneas */}
          <table style={{ fontSize: 12, marginBottom: 12 }}>
            <thead><tr style={{ background: 'var(--primary)', color: 'white' }}>
              <th style={{ padding: '6px 8px', textAlign: 'left' }}>Concepto</th>
              <th style={{ padding: '6px 8px', textAlign: 'center' }}>Cant.</th>
              <th style={{ padding: '6px 8px', textAlign: 'right' }}>Precio</th>
              <th style={{ padding: '6px 8px', textAlign: 'right' }}>Base</th>
              <th style={{ padding: '6px 8px', textAlign: 'right' }}>IVA</th>
              <th style={{ padding: '6px 8px', textAlign: 'right', fontWeight: 700 }}>Total</th>
            </tr></thead>
            <tbody>
              {f.lineas.map(l => (
                <tr key={l.id} style={{ borderBottom: '1px solid var(--border)' }}>
                  <td style={{ padding: '6px 8px' }}>{l.concepto}</td>
                  <td style={{ padding: '6px 8px', textAlign: 'center' }}>{l.cantidad} {l.unidad}</td>
                  <td style={{ padding: '6px 8px', textAlign: 'right' }}>{l.precioUnidad.toFixed(2)} €</td>
                  <td style={{ padding: '6px 8px', textAlign: 'right' }}>{l.baseLinea.toFixed(2)} €</td>
                  <td style={{ padding: '6px 8px', textAlign: 'right' }}>{l.porcentajeIva.toFixed(0)}%</td>
                  <td style={{ padding: '6px 8px', textAlign: 'right', fontWeight: 600 }}>{l.totalLinea.toFixed(2)} €</td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* Totales */}
          <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
            <div style={{ width: 240, fontSize: 13 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', padding: '4px 0', borderBottom: '1px solid var(--border)' }}>
                <span style={{ color: 'var(--text-muted)' }}>Base imponible</span><span>{f.baseImponible.toFixed(2)} €</span>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', padding: '4px 0', borderBottom: '1px solid var(--border)' }}>
                <span style={{ color: 'var(--text-muted)' }}>IVA ({f.porcentajeIva.toFixed(0)}%)</span><span>{f.cuotaIva.toFixed(2)} €</span>
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 0', fontWeight: 800, fontSize: 15, color: 'var(--primary)' }}>
                <span>TOTAL</span><span>{f.total.toFixed(2)} €</span>
              </div>
            </div>
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-accent" onClick={onImprimir}>🖨️ Imprimir / PDF</button>
          <button className="btn btn-ghost" onClick={onClose}>Cerrar</button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL FACTURAR DESDE RESERVA ─────────────────────────────────────────────
function FacturarReservaModal({ onSaved, onCancel }: { onSaved: () => void; onCancel: () => void }) {
  const [reservas, setReservas]     = useState<Reserva[]>([]);
  const [loading, setLoading]       = useState(true);
  const [generando, setGenerando]   = useState(false);
  const [buscar, setBuscar]         = useState('');
  const [formaPago, setFormaPago]   = useState('Transferencia');

  useEffect(() => {
    apiFetch<Reserva[]>('/api/reservas').then(r => setReservas(r.filter(x => !x.facturado && !x.esBloqueo))).catch(() => {}).finally(() => setLoading(false));
  }, []);

  const filtradas = reservas.filter(r => {
    const b = buscar.toLowerCase();
    return !buscar || r.huespedNombreCompleto?.toLowerCase().includes(b) ||
      r.huespedDni?.toLowerCase().includes(b) || r.habitacionNumero.includes(b);
  });

  const facturar = async (reservaId: string) => {
    setGenerando(true);
    try {
      await apiFetch(`/api/facturas/desde-reserva/${reservaId}`, { method: 'POST', body: JSON.stringify({ formaPago }) });
      onSaved();
    } catch (e: any) { alert(e.message); setGenerando(false); }
  };

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onCancel()}>
      <div className="modal" style={{ maxWidth: 700 }}>
        <div className="modal-header">
          <h3>📋 Facturar desde reserva</h3>
          <button className="modal-close" onClick={onCancel}>×</button>
        </div>
        <div className="modal-body">
          <div style={{ display: 'flex', gap: 10, marginBottom: 14, alignItems: 'center' }}>
            <input placeholder="Buscar huésped, DNI, habitación..." value={buscar} onChange={e => setBuscar(e.target.value)}
              style={{ flex: 1, padding: '8px 12px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }} />
            <label style={{ fontSize: 13, display: 'flex', alignItems: 'center', gap: 6, whiteSpace: 'nowrap' }}>
              Pago:
              <select value={formaPago} onChange={e => setFormaPago(e.target.value)}
                style={{ padding: '7px 10px', border: '1px solid var(--border)', borderRadius: 7, fontSize: 13 }}>
                <option>Transferencia</option><option>Efectivo</option><option>Tarjeta</option><option>Bizum</option>
              </select>
            </label>
          </div>
          {loading ? <div style={{ padding: 20, textAlign: 'center' }}><span className="spinner"></span></div>
            : filtradas.length === 0 ? (
              <div className="empty-state" style={{ padding: 20 }}>
                <p style={{ color: 'var(--text-muted)' }}>No hay reservas pendientes de facturar{buscar ? ' con ese criterio' : ''}.</p>
              </div>
            ) : (
              <div style={{ maxHeight: 380, overflowY: 'auto' }}>
                {filtradas.map(r => (
                  <div key={r.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 12px', borderBottom: '1px solid var(--border)', gap: 10 }}>
                    <div style={{ flex: 1 }}>
                      <strong>{r.huespedNombreCompleto ?? '—'}</strong>
                      <span style={{ marginLeft: 8, fontSize: 11, color: 'var(--text-muted)' }}>{r.huespedDni}</span>
                      <div style={{ fontSize: 12, color: 'var(--text-muted)', marginTop: 2 }}>
                        {r.residenciaNombre} · Hab. {r.habitacionNumero} · {r.fechaEntrada} → {r.fechaSalida} ({r.totalNoches}n)
                      </div>
                    </div>
                    <div style={{ textAlign: 'right' }}>
                      <div style={{ fontWeight: 700, color: 'var(--primary)' }}>{r.importeTotal.toFixed(2)} €</div>
                      <button className="btn btn-primary btn-sm" style={{ marginTop: 4 }}
                        onClick={() => facturar(r.id)} disabled={generando}>
                        🧾 Facturar
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
        </div>
        <div className="modal-footer">
          <button className="btn btn-ghost" onClick={onCancel}>Cerrar</button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL FACTURA MANUAL ──────────────────────────────────────────────────────
type LineaForm = { orden: number; concepto: string; cantidad: string; unidad: string; precioUnidad: string; descuento: string; porcentajeIva: string };

function FacturaManualModal({ residencias, onSaved, onCancel }: { residencias: Residencia[]; onSaved: () => void; onCancel: () => void }) {
  const hoy = new Date().toISOString().slice(0, 10);
  const [f, setF] = useState({
    residenciaId: residencias[0]?.id ?? '',
    huespedId: '',
    fechaEmision: hoy,
    fechaVencimiento: '',
    destinatarioNombre: '', destinatarioDni: '',
    destinatarioDireccion: '', destinatarioCp: '', destinatarioMunicipio: '',
    formaPago: 'Transferencia', observaciones: '', porcentajeIva: '10'
  });
  const [lineas, setLineas] = useState<LineaForm[]>([
    { orden: 1, concepto: 'Alojamiento', cantidad: '1', unidad: 'noche', precioUnidad: '0', descuento: '0', porcentajeIva: '10' }
  ]);
  const [loading, setLoading] = useState(false);
  const set = (k: string, v: string) => setF(p => ({ ...p, [k]: v }));

  const setLinea = (i: number, k: keyof LineaForm, v: string) =>
    setLineas(prev => prev.map((l, idx) => idx === i ? { ...l, [k]: v } : l));

  const addLinea = () => setLineas(p => [...p, { orden: p.length + 1, concepto: '', cantidad: '1', unidad: 'ud', precioUnidad: '0', descuento: '0', porcentajeIva: f.porcentajeIva }]);
  const removeLinea = (i: number) => setLineas(p => p.filter((_, idx) => idx !== i));

  // Preview totales
  const preview = lineas.reduce((acc, l) => {
    const bruto = +l.cantidad * +l.precioUnidad;
    const base  = Math.round(bruto * (1 - +l.descuento / 100) * 100) / 100;
    const iva   = Math.round(base * (+l.porcentajeIva / 100) * 100) / 100;
    acc.base += base; acc.iva += iva;
    return acc;
  }, { base: 0, iva: 0 });

  const submit = async (e: React.FormEvent) => {
    e.preventDefault(); setLoading(true);
    try {
      await apiFetch('/api/facturas', {
        method: 'POST',
        body: JSON.stringify({
          ...f,
          huespedId: f.huespedId || null,
          fechaVencimiento: f.fechaVencimiento || null,
          porcentajeIva: +f.porcentajeIva,
          lineas: lineas.map(l => ({
            orden: l.orden, concepto: l.concepto,
            cantidad: +l.cantidad, unidad: l.unidad,
            precioUnidad: +l.precioUnidad, descuento: +l.descuento, porcentajeIva: +l.porcentajeIva
          }))
        })
      });
      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onCancel()}>
      <div className="modal" style={{ maxWidth: 760 }}>
        <div className="modal-header">
          <h3>🧾 Nueva factura manual</h3>
          <button className="modal-close" onClick={onCancel}>×</button>
        </div>
        <form onSubmit={submit}>
          <div className="modal-body" style={{ maxHeight: '70vh', overflowY: 'auto' }}>
            <div className="form-grid">
              <div className="form-group">
                <label>Residencia (Emisor) *</label>
                <select value={f.residenciaId} onChange={e => set('residenciaId', e.target.value)} required>
                  {residencias.map(r => <option key={r.id} value={r.id}>{r.nombre}</option>)}
                </select>
              </div>
              <div className="form-group"><label>Fecha emisión *</label><input type="date" value={f.fechaEmision} onChange={e => set('fechaEmision', e.target.value)} required /></div>
              <div className="form-group"><label>Fecha vencimiento</label><input type="date" value={f.fechaVencimiento} onChange={e => set('fechaVencimiento', e.target.value)} /></div>
              <div className="form-group">
                <label>Forma de pago</label>
                <select value={f.formaPago} onChange={e => set('formaPago', e.target.value)}>
                  <option>Transferencia</option><option>Efectivo</option><option>Tarjeta</option><option>Bizum</option>
                </select>
              </div>
            </div>

            <div style={{ fontWeight: 600, fontSize: 13, color: 'var(--primary)', margin: '12px 0 8px' }}>👤 Destinatario</div>
            <div className="form-grid">
              <div className="form-group"><label>Nombre *</label><input value={f.destinatarioNombre} onChange={e => set('destinatarioNombre', e.target.value)} required /></div>
              <div className="form-group"><label>DNI/NIF *</label><input value={f.destinatarioDni} onChange={e => set('destinatarioDni', e.target.value.toUpperCase())} required /></div>
              <div className="form-group"><label>Dirección</label><input value={f.destinatarioDireccion} onChange={e => set('destinatarioDireccion', e.target.value)} /></div>
              <div className="form-group"><label>C.P.</label><input value={f.destinatarioCp} onChange={e => set('destinatarioCp', e.target.value)} /></div>
              <div className="form-group"><label>Municipio</label><input value={f.destinatarioMunicipio} onChange={e => set('destinatarioMunicipio', e.target.value)} /></div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', margin: '14px 0 8px' }}>
              <div style={{ fontWeight: 600, fontSize: 13, color: 'var(--primary)' }}>📋 Líneas de factura</div>
              <button type="button" className="btn btn-ghost btn-sm" onClick={addLinea}>+ Añadir línea</button>
            </div>

            <div style={{ background: 'var(--surface-2)', borderRadius: 8, padding: 12 }}>
              {lineas.map((l, i) => (
                <div key={i} style={{ display: 'grid', gridTemplateColumns: '3fr 1fr 1fr 1fr 1fr 1fr auto', gap: 6, marginBottom: 8, alignItems: 'center' }}>
                  <input placeholder="Concepto" value={l.concepto} onChange={e => setLinea(i,'concepto',e.target.value)} style={{ padding: '6px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12 }} required />
                  <input placeholder="Cant." type="number" value={l.cantidad} onChange={e => setLinea(i,'cantidad',e.target.value)} style={{ padding: '6px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12 }} />
                  <select value={l.unidad} onChange={e => setLinea(i,'unidad',e.target.value)} style={{ padding: '6px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12 }}>
                    <option value="noche">noche</option><option value="mes">mes</option><option value="ud">ud.</option><option value="servicio">servicio</option>
                  </select>
                  <input placeholder="€/ud." type="number" step="0.01" value={l.precioUnidad} onChange={e => setLinea(i,'precioUnidad',e.target.value)} style={{ padding: '6px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12 }} />
                  <input placeholder="Dto.%" type="number" value={l.descuento} onChange={e => setLinea(i,'descuento',e.target.value)} style={{ padding: '6px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12 }} />
                  <input placeholder="IVA%" type="number" value={l.porcentajeIva} onChange={e => setLinea(i,'porcentajeIva',e.target.value)} style={{ padding: '6px 8px', border: '1px solid var(--border)', borderRadius: 6, fontSize: 12 }} />
                  {lineas.length > 1 && <button type="button" style={{ background: 'none', border: 'none', color: '#dc3545', cursor: 'pointer', fontSize: 16 }} onClick={() => removeLinea(i)}>✕</button>}
                </div>
              ))}

              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 20, marginTop: 10, fontSize: 13, fontWeight: 600 }}>
                <span>Base: {preview.base.toFixed(2)} €</span>
                <span>IVA: {preview.iva.toFixed(2)} €</span>
                <span style={{ color: 'var(--primary)', fontSize: 15 }}>Total: {(preview.base + preview.iva).toFixed(2)} €</span>
              </div>
            </div>

            <div className="form-group" style={{ marginTop: 12 }}>
              <label>Observaciones / pie de factura</label>
              <textarea value={f.observaciones} onChange={e => set('observaciones', e.target.value)} rows={2} />
            </div>
          </div>
          <div className="modal-footer">
            <button type="button" className="btn btn-ghost" onClick={onCancel}>Cancelar</button>
            <button type="submit" className="btn btn-primary" disabled={loading}>{loading ? 'Creando...' : '💾 Crear factura'}</button>
          </div>
        </form>
      </div>
    </div>
  );
}
