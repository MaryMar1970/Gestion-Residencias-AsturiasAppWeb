import { useState, useEffect, useCallback } from 'react';
import { apiFetch } from '../api';
import { Reserva, Residencia } from '../types';

export function useBloqueosData() {
  const [bloqueos, setBloqueos] = useState<Reserva[]>([]);
  const [loading, setLoading] = useState(true);
  const [buscar, setBuscar] = useState('');
  const [statusMsg, setStatusMsg] = useState('');

  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [residenciaActiva, setResidenciaActiva] = useState<string>(
    () => localStorage.getItem('pref_residencia_bloqueos') || 'Todas'
  );

  const [sortField, setSortField] = useState<'numeroOrden' | 'fechaEntrada' | 'fechaSalida'>('numeroOrden');
  const [sortAsc, setSortAsc] = useState(true);

  const cargar = useCallback(async () => {
    setLoading(true);
    try {
      const data = await apiFetch<Reserva[]>('/api/reservas');
      const filtrados = data.filter(r => !r.huespedId);
      setBloqueos(filtrados);
    } catch {} finally {
      setLoading(false);
    }
  }, []);

  const cargarResidencias = useCallback(async () => {
    try {
      const data = await apiFetch<Residencia[]>('/api/residencias');
      setResidencias(data);
    } catch {}
  }, []);

  useEffect(() => {
    void cargar();
    void cargarResidencias();
  }, [cargar, cargarResidencias]);

  const eliminar = async (id: string) => {
    if (!confirm('¿Eliminar/Cancelar este bloqueo/reserva interna?')) return;
    try {
      await apiFetch(`/api/reservas/${id}`, { method: 'DELETE' });
      setStatusMsg('Registro eliminado correctamente.');
      void cargar();
    } catch (e: any) {
      alert(e.message);
    }
  };

  const handleSort = (field: 'numeroOrden' | 'fechaEntrada' | 'fechaSalida') => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const bloqueosFiltrados = bloqueos.filter(r => {
    const b = buscar.toLowerCase();
    const matchBuscar =
      !buscar ||
      (r.habitacionNumero || '').includes(b) ||
      (r.motivoBloqueo && r.motivoBloqueo.toLowerCase().includes(b)) ||
      (r.residenciaNombre || '').toLowerCase().includes(b);
    const matchResidencia = residenciaActiva === 'Todas' || r.residenciaNombre === residenciaActiva;
    return matchBuscar && matchResidencia;
  });

  const bloqueosOrdenados = [...bloqueosFiltrados].sort((a, b) => {
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

  return {
    bloqueos,
    loading,
    buscar, setBuscar,
    statusMsg, setStatusMsg,
    residencias,
    residenciaActiva, setResidenciaActiva,
    sortField, sortAsc, handleSort,
    bloqueosFiltrados,
    bloqueosOrdenados,
    eliminar,
    cargar,
  };
}
