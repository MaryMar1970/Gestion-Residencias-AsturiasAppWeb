import { User, Page } from '../../types';

export function Sidebar({ page, setPage, user, onLogout }: { page: Page; setPage: (p: Page) => void; user: User; onLogout: () => void }) {
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
        {navItem('reservas', '📋', 'Reservas activas')}
        {navItem('facturas', '🧾', 'Facturación')}
        <div className="nav-section">Configuración</div>
        {navItem('residencias', '🏠', 'Residencias')}
        {navItem('alojamientos', '🔑', 'Alojamientos')}
        {navItem('bloqueos', '🔒', 'Bloqueos')}
        {navItem('tarifas', '💰', 'Tarifas y precios')}
        {navItem('festivos', '🎉', 'Festivos')}
        {navItem('evaluacion', '⚖️', 'Evaluación solicitudes')}
      </nav>
      <div className="sidebar-footer">
        <div className="user-info">
          <div className="user-avatar">{user.fullName?.[0] ?? '?'}</div>
          <div>
            <div className="user-name">{user.fullName}</div>
            <div className="user-role">{isAdmin ? '👑 Administrador' : '👤 Recepcionista'}</div>
          </div>
        </div>
        <button className="btn-logout" onClick={onLogout}>🚪 Cerrar sesión</button>
      </div>
    </div>
  );
}