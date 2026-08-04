import { useState, FormEvent } from 'react';
import { apiFetch } from '../../api';
import { User } from '../../types';

export function LoginPage({ onLogin, theme, toggleTheme }: { onLogin: (token: string, user: User) => void; theme: 'light' | 'dark'; toggleTheme: () => void }) {
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
    <div className="login-page" style={{ position: 'relative' }}>
      <div style={{ position: 'absolute', top: 20, right: 20 }}>
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