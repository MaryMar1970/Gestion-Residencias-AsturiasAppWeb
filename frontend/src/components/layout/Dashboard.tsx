import { useState, useEffect } from 'react';
import { apiFetch } from '../../api';
import { Residencia, Habitacion, TipoHab, Festivo, Page } from '../../types';

export function Dashboard({ setPage }: { setPage: (p: Page) => void }) {
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