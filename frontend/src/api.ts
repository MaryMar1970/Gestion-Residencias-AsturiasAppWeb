let activeApiBase = 'http://localhost:5260';
export const API_BASE = activeApiBase;

export async function apiFetch<T>(
  path: string,
  options?: RequestInit
): Promise<T> {
  const token = localStorage.getItem('token');
  let res: Response;
  try {
    res = await fetch(`${activeApiBase}${path}`, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...options?.headers,
      },
    });
  } catch (err) {
    const fallbackPort = activeApiBase.includes('5260') ? 'http://localhost:5000' : 'http://localhost:5260';
    try {
      res = await fetch(`${fallbackPort}${path}`, {
        ...options,
        headers: {
          'Content-Type': 'application/json',
          ...(token ? { Authorization: `Bearer ${token}` } : {}),
          ...options?.headers,
        },
      });
      activeApiBase = fallbackPort;
    } catch {
      throw err;
    }
  }
  if (res.status === 401) {
    localStorage.removeItem('token');
    window.location.reload();
    throw new Error('Sesión expirada. Inicie sesión de nuevo.');
  }
  if (!res.ok) {
    const text = await res.text();
    throw new Error(text || `Error ${res.status}`);
  }
  if (res.status === 204) return undefined as T;
  return res.json();
}

export type DateFormatStyle = 'dd-mm-yyyy' | 'dd/mm/yyyy' | 'short_month' | 'short_day_month';

export function formatFechaDisplay(
  isoStr?: string | null,
  style: DateFormatStyle = 'dd-mm-yyyy'
): string {
  if (!isoStr) return '—';
  const clean = isoStr.split('T')[0];
  const parts = clean.split('-');
  if (parts.length !== 3) return isoStr;
  const [y, m, d] = parts;
  if (!y || !m || !d) return isoStr;

  const day = d.padStart(2, '0');
  const month = m.padStart(2, '0');

  switch (style) {
    case 'dd/mm/yyyy':
      return `${day}/${month}/${y}`;
    case 'short_month': {
      const mesesAbr = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];
      const idx = parseInt(month, 10) - 1;
      const mesText = (idx >= 0 && idx < 12) ? mesesAbr[idx] : month;
      return `${day}/${mesText}`;
    }
    case 'short_day_month':
      return `${day}/${month}`;
    case 'dd-mm-yyyy':
    default:
      return `${day}-${month}-${y}`;
  }
}

