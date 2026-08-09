import { useState, useEffect } from 'react';
import './index.css';
import { apiFetch } from './api';
import { User, Page } from './types';
import { cargarMatrizEvaluacion } from './utils/evaluacionUtils';

// ─── Layout Components ─────────────────────────────────────────────────────────
import { LoginPage } from './components/layout/LoginPage';
import { Sidebar } from './components/layout/Sidebar';
import { Dashboard } from './components/layout/Dashboard';

// ─── Feature Components (from Reservas.tsx refactor) ────────────────────────────
import { CalendarioGantt } from './components/calendario/CalendarioGantt';
import { NuevaReservaModal } from './components/reservas/NuevaReservaModal';
import { ReservasPage } from './components/reservas/ReservasPage';
import { HuespedesPage } from './components/solicitudes/HuespedesPage';
import { BloqueosPage } from './components/bloqueos/BloqueosPage';
import { FacturasPage } from './Facturas';

// ─── Configuration Components ──────────────────────────────────────────────────
import { ResidenciasPage } from './components/configuracion/ResidenciasPage';
import { AlojamientosPage } from './components/configuracion/AlojamientosPage';
import { TarifasPage } from './components/configuracion/TarifasPage';
import { FestivosPage } from './components/configuracion/FestivosPage';
import { EvaluacionSolicitudesPage } from './components/configuracion/EvaluacionSolicitudesPage';

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
  evaluacion:   { title: 'Evaluación solicitudes',  sub: 'Matriz de prioridades y ponderación de solicitudes' },
};

// ─── App Root ───────────────────────────────────────────────────────────────────
export default function App() {
  const [token, setToken] = useState<string | null>(localStorage.getItem('token'));
  const [user, setUser]   = useState<User | null>(null);
  const [page, setPage]   = useState<Page>('dashboard');

  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    const stored = localStorage.getItem('theme');
    if (stored === 'light' || stored === 'dark') return stored;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  });

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
  }, [theme]);

  const toggleTheme = () => setTheme(prev => prev === 'light' ? 'dark' : 'light');

  useEffect(() => {
    if (token) {
      if (!user) {
        apiFetch<User>('/users/me').then(setUser).catch(() => { localStorage.removeItem('token'); setToken(null); });
      }
      void cargarMatrizEvaluacion();
    }
  }, [token, user]);

  const handleLogin = (t: string, u: User) => { setToken(t); setUser(u); };
  const handleLogout = () => { localStorage.removeItem('token'); setToken(null); setUser(null); };

  if (!token || !user) return <LoginPage onLogin={handleLogin} theme={theme} toggleTheme={toggleTheme} />;

  const pt = PAGE_TITLES[page];
  return (
    <div className="app-layout">
      <Sidebar page={page} setPage={setPage} user={user} onLogout={handleLogout} />
      <div className="main-content">
        <div className="topbar">
          <div><div className="topbar-title">{pt.title}</div><div className="topbar-subtitle">{pt.sub}</div></div>
          <button className="theme-toggle" onClick={toggleTheme} aria-label="Cambiar tema">
            {theme === 'light' ? (
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#c9a84c' }}>
                <circle cx="12" cy="12" r="5"></circle>
                <line x1="12" y1="1" x2="12" y2="3"></line>
                <line x1="12" y1="21" x2="12" y2="23"></line>
                <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line>
                <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line>
                <line x1="1" y1="12" x2="3" y2="12"></line>
                <line x1="21" y1="12" x2="23" y2="12"></line>
                <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line>
                <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line>
              </svg>
            ) : (
              <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#a5b4fc' }}>
                <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path>
              </svg>
            )}
          </button>
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
          {page === 'evaluacion'   && <EvaluacionSolicitudesPage />}
          {page === 'bloqueos'     && <BloqueosPage />}
        </div>
      </div>
    </div>
  );
}
