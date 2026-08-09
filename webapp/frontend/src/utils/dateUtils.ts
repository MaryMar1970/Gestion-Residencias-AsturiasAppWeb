// ─── Date Utilities ─────────────────────────────────────────────────────────

export function addDays(date: Date, n: number): Date {
  const d = new Date(date); d.setDate(d.getDate() + n); return d;
}

export function toDateStr(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

export function formatFechaCorta(isoStr: string): string {
  if (!isoStr) return '';
  const clean = isoStr.split('T')[0];
  const parts = clean.split('-');
  if (parts.length !== 3) return isoStr;
  const [y, m, d] = parts;
  return `${d.padStart(2, '0')}/${m.padStart(2, '0')}/${y}`;
}

export function parseLocal(s: string): Date {
  const [y,m,d] = s.split('-').map(Number);
  return new Date(y, m-1, d);
}

export function diasEntre(desde: string, hasta: string): Date[] {
  const out: Date[] = [];
  const cur = parseLocal(desde);
  const end = parseLocal(hasta);
  while (cur < end) { out.push(new Date(cur)); cur.setDate(cur.getDate() + 1); }
  return out;
}
