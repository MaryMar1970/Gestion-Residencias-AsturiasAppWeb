import { useState, useEffect, useCallback } from 'react';
import { apiFetch } from '../api';
import { Reserva, Residencia } from '../types';

export function useReservas() {
  const [reservas, setReservas] = useState<Reserva[]>([]);
  const [loading, setLoading] = useState(true);
  const [buscar, setBuscar] = useState('');
  const [filtroEstado, setFiltroEstado] = useState('');
  const [statusMsg, setStatusMsg] = useState('');

  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>(
    () => localStorage.getItem('pref_residencia_reservas') || 'Todas'
  );

  const [sortField, setSortField] = useState<'numeroOrden' | 'fechaEntrada' | 'fechaSalida'>('numeroOrden');
  const [sortAsc, setSortAsc] = useState(true);

  const [pagina, setPagina] = useState(1);
  const filasPorPagina = 18;

  const cargar = useCallback(async () => {
    setLoading(true);
    try {
      const res = await apiFetch<any>('/api/reservas?pageSize=500');
      setReservas(Array.isArray(res) ? res : (res?.items ?? []));
    } catch {} finally {
      setLoading(false);
    }
  }, []);

  const cargarResidencias = useCallback(async () => {
    try {
      const list = await apiFetch<Residencia[]>('/api/residencias');
      setResidencias(list);
      if (list.length > 0) {
        const saved = localStorage.getItem('pref_residencia_reservas');
        if (!saved || saved === 'Todas' || !list.some(r => r.nombre === saved)) {
          setResidenciaActiva(list[0].nombre);
          localStorage.setItem('pref_residencia_reservas', list[0].nombre);
        }
      }
    } catch {}
  }, []);

  useEffect(() => {
    void cargar();
    void cargarResidencias();
  }, [cargar, cargarResidencias]);

  const handleSort = (field: 'numeroOrden' | 'fechaEntrada' | 'fechaSalida') => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const reservasFiltradas = reservas.filter(r => {
    if (r.esBloqueo || !r.huespedId) return false;

    const b = buscar.toLowerCase();
    const matchBuscar =
      !buscar ||
      r.huespedNombreCompleto?.toLowerCase().includes(b) ||
      r.huespedDni?.toLowerCase().includes(b) ||
      (r.habitacionNumero || '').includes(b) ||
      (r.residenciaNombre || '').toLowerCase().includes(b);
    const matchEstado = !filtroEstado || r.estado === filtroEstado;
    const matchResidencia = r.residenciaNombre === residenciaActiva;

    return matchBuscar && matchEstado && matchResidencia;
  });

  const reservasOrdenadas = [...reservasFiltradas].sort((a, b) => {
    let comparison = 0;
    if (sortField === 'numeroOrden') {
      comparison = a.numeroOrden - b.numeroOrden;
    } else if (sortField === 'fechaEntrada') {
      comparison = a.fechaEntrada.localeCompare(b.fechaEntrada);
    } else if (sortField === 'fechaSalida') {
      comparison = a.fechaSalida.localeCompare(b.fechaSalida);
    }
    return sortAsc ? comparison : -comparison;
  });

  const totalPaginas = Math.ceil(reservasOrdenadas.length / filasPorPagina) || 1;
  const paginatedReservas = reservasOrdenadas.slice((pagina - 1) * filasPorPagina, pagina * filasPorPagina);

  return {
    reservas,
    loading,
    buscar, setBuscar,
    filtroEstado, setFiltroEstado,
    statusMsg, setStatusMsg,
    residencias,
    residenciaActiva, setResidenciaActiva,
    sortField, sortAsc, handleSort,
    pagina, setPagina,
    filasPorPagina,
    totalPaginas,
    paginatedReservas,
    reservasFiltradas,
    reservasOrdenadas,
    cargar,
  };
}
