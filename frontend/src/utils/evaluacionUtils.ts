// ─── Algoritmo de evaluación (matriz de prioridades dinámica) ──────────────────
import { apiFetch } from '../api';

let cachedMatriz: Record<string, string> | null = null;

export function setMatrizEvaluacionCache(matriz: Record<string, string>) {
  cachedMatriz = matriz;
}

export async function cargarMatrizEvaluacion(): Promise<Record<string, string>> {
  try {
    const matriz = await apiFetch<Record<string, string>>('/api/configuracion/matriz-evaluacion');
    if (matriz && typeof matriz === 'object') {
      cachedMatriz = matriz;
      return matriz;
    }
  } catch (err) {
    console.warn('No se pudo cargar la matriz de evaluación desde la API, usando respaldo:', err);
  }
  return matrizPorDefecto;
}

export function calcularEvaluacion(
  finalidad: string,
  empleoCategoria: string,
  situacion: string,
  customMatriz?: Record<string, string>
): string {
  const f = (finalidad || 'Otros').trim().toLowerCase();
  let finalidadMapeada = 'Otros';
  if (f.includes('comision no') || f.includes('comisión no') || f.includes('indem')) {
    finalidadMapeada = 'Comisión NO indem.';
  } else if (f.includes('comision') || f.includes('comisión')) {
    finalidadMapeada = 'Comisión';
  } else if (f.includes('destino')) {
    finalidadMapeada = 'Destino';
  } else if (f.includes('enfermedad')) {
    finalidadMapeada = 'Enfermedad';
  } else if (f.includes('urgencia')) {
    finalidadMapeada = 'Urgencia';
  } else if (f.includes('sepelio')) {
    finalidadMapeada = 'Sepelio';
  } else if (f.includes('estancia')) {
    finalidadMapeada = 'Máx. Estancia';
  }

  const e = (empleoCategoria || 'GC').trim().toLowerCase();
  let empleoCat = 'GC';
  if (e.includes('alumno')) {
    empleoCat = 'Alumno';
  } else if (e.includes('funcionario')) {
    empleoCat = 'Funcionario en GC';
  } else if (e.includes('militar en')) {
    empleoCat = 'Militar en GC';
  } else if (e.includes('militar no') || e.includes('militar')) {
    empleoCat = 'Militar no  GC';
  }

  const s = (situacion || 'Activo').trim().toLowerCase();
  let situacionMapeada = 'Activo';
  if (s.includes('viogen')) {
    situacionMapeada = 'Viogen';
  } else if (s.includes('asoc')) {
    situacionMapeada = 'Asociación';
  } else if (s.includes('reserva activo') || s.includes('reserva activa')) {
    situacionMapeada = 'Reserva activo';
  } else if (s.includes('reserva')) {
    situacionMapeada = 'Reserva';
  } else if (s.includes('excedencia')) {
    situacionMapeada = 'Excedencia';
  } else if (s.includes('especial')) {
    situacionMapeada = 'Especiales';
  } else if (s.includes('retirado') || s.includes('jubilado')) {
    situacionMapeada = 'Retirado';
  } else if (s.includes('viuda')) {
    situacionMapeada = 'Viuda';
  } else if (s.includes('huerfano') || s.includes('huérfano')) {
    situacionMapeada = 'Huerfano';
  }

  const clave = `${finalidadMapeada.toLowerCase()}|${empleoCat.toLowerCase()}|${situacionMapeada.toLowerCase()}`;

  const matrizActiva = customMatriz || cachedMatriz || matrizPorDefecto;

  return matrizActiva[clave] || '(NO VÁLIDO)';
}

export const matrizPorDefecto: Record<string, string> = {
  'comisión no indem.|gc|viogen': '1, 1, 1',
  'comisión no indem.|gc|activo': '1, 1, 2',
  'comisión no indem.|gc|reserva activo': '1, 1, 2',
  'destino|gc|activo': '1, 1, 3',
  'comisión|gc|activo': '1, 1, 4',
  'enfermedad|gc|retirado': '2, 8, 1',
  'comisión|alumno|activo': '1, 2, 2',
  'comisión|militar en gc|activo': '1, 3, 2',
  'destino|militar en gc|activo': '1, 3, 2',
  'comisión|funcionario en gc|activo': '1, 4, 2',
  'destino|funcionario en gc|activo': '1, 4, 2',
  'enfermedad|gc|activo': '2, 1, 1',
  'otros|gc|viogen': '2, 1, 2',
  'urgencia|gc|activo': '2, 1, 3',
  'sepelio|gc|activo': '2, 1, 4',
  'máx. estancia|gc|activo': '2, 1, 5',
  'otros|gc|asociación': '2, 1, 6',
  'otros|gc|activo': '2, 1, 7',
  'otros|gc|reserva activo': '2, 1, 8',
  'otros|alumno|activo': '2, 2, 7',
  'otros|gc|reserva': '2, 3, 7',
  'otros|militar en gc|activo': '2, 4, 7',
  'otros|funcionario en gc|activo': '2, 5, 7',
  'otros|gc|excedencia': '2, 6, 7',
  'otros|gc|especiales': '2, 7, 7',
  'otros|gc|retirado': '2, 8, 7',
  'otros|gc|viuda': '2, 9, 7',
  'otros|gc|huerfano': '2, 9, 7',
  'otros|militar no  gc|activo': '2, 10, 7',
  'otros|militar no  gc|reserva': '2, 10, 7',
  'otros|militar no  gc|retirado': '2, 10, 7'
};
