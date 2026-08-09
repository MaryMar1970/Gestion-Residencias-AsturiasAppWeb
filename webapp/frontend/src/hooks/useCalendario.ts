import { useState, useEffect, useRef, useCallback } from 'react';
import { apiFetch, formatFechaDisplay } from '../api';
import { formatFechaCorta, addDays, toDateStr, parseLocal } from '../utils/dateUtils';
import { Residencia, Reserva, CalData, Festivo as FestivoType } from '../types';

export function useCalendario() {
  const hoy = new Date();
  const [inicio, setInicio] = useState<string>(() => {
    const gotoFecha = sessionStorage.getItem('gotoReservaFecha');
    if (gotoFecha) {
      const checkinDate = parseLocal(gotoFecha);
      const startView = addDays(checkinDate, -2);
      return toDateStr(startView);
    }
    const prefInicio = localStorage.getItem('pref_calendario_inicio');
    if (prefInicio) return prefInicio;
    const d = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    return toDateStr(d);
  });

  const [dias, setDias] = useState<number>(() => {
    const gotoId = sessionStorage.getItem('gotoReservaId');
    if (gotoId) return 30;
    const prefDias = localStorage.getItem('pref_calendario_dias');
    if (prefDias) {
      const parsed = parseInt(prefDias, 10);
      if (!isNaN(parsed) && parsed > 0) return parsed;
    }
    return 10;
  });

  const [residenciaFiltro, setResidenciaFiltro] = useState<string>(() => localStorage.getItem('pref_residencia_calendario') || '');
  const [calData, setCalData] = useState<CalData | null>(null);
  const [festivos, setFestivos] = useState<Set<string>>(new Set());
  const [residencias, setResidencias] = useState<Residencia[]>([]);
  const [loading, setLoading] = useState(false);
  const [tooltip, setTooltip] = useState<{ reserva: Reserva; x: number; y: number } | null>(null);
  const [reservaSeleccionada, setReservaSeleccionada] = useState<Reserva | null>(null);
  const [highlightReservaId, setHighlightReservaId] = useState<string | null>(null);
  const [draggedReserva, setDraggedReserva] = useState<Reserva | null>(null);
  const [dragOverTarget, setDragOverTarget] = useState<{ habitacionId: string; fecha: string } | null>(null);
  const [moverStatus, setMoverStatus] = useState<{ title: string; details: string; isError?: boolean } | null>(null);
  const tooltipTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    if (inicio) localStorage.setItem('pref_calendario_inicio', inicio);
  }, [inicio]);

  useEffect(() => {
    if (dias) localStorage.setItem('pref_calendario_dias', dias.toString());
  }, [dias]);

  const cargarResidencias = useCallback(async () => {
    try {
      const data = await apiFetch<Residencia[]>('/api/residencias');
      setResidencias(data);
      const pref = localStorage.getItem('pref_residencia_calendario');
      if (pref && !data.some(r => r.id === pref)) {
        setResidenciaFiltro('');
        localStorage.removeItem('pref_residencia_calendario');
      }
    } catch {}
  }, []);

  const cargarFestivos = useCallback(async () => {
    try {
      const finDate = addDays(parseLocal(inicio), dias);
      const finStr = toDateStr(finDate);
      const data = await apiFetch<FestivoType[]>(`/api/festivos?inicio=${inicio}&fin=${finStr}`);
      setFestivos(new Set(data.map(f => f.fecha)));
    } catch {}
  }, [inicio, dias]);

  const cargarCalData = useCallback(async () => {
    setLoading(true);
    try {
      const finDate = addDays(parseLocal(inicio), dias);
      const finStr = toDateStr(finDate);
      let url = `/api/calendario?fechaInicio=${inicio}&fechaFin=${finStr}&inicio=${inicio}&dias=${dias}`;
      if (residenciaFiltro) url += `&residenciaId=${residenciaFiltro}`;
      console.log("[Calendario] Fetching:", url);
      const data = await apiFetch<CalData>(url);
      console.log("[Calendario] Response:", JSON.stringify(data).substring(0, 500));
      console.log("[Calendario] Habitaciones count:", data?.habitaciones?.length ?? 'N/A');
      setCalData(data);
    } catch (err) {
      console.error("Error al cargar datos del calendario:", err);
    } finally {
      setLoading(false);
    }
  }, [inicio, dias, residenciaFiltro]);

  useEffect(() => {
    void cargarResidencias();
  }, [cargarResidencias]);

  useEffect(() => {
    void cargarCalData();
    void cargarFestivos();
  }, [cargarCalData, cargarFestivos]);

  useEffect(() => {
    const gotoId = sessionStorage.getItem('gotoReservaId');
    if (gotoId) {
      setHighlightReservaId(gotoId);
      sessionStorage.removeItem('gotoReservaId');
      sessionStorage.removeItem('gotoReservaFecha');
      const timer = setTimeout(() => setHighlightReservaId(null), 6000);
      return () => clearTimeout(timer);
    }
  }, []);

  const handleDropReserva = async (res: Reserva, targetHabitacionId: string, targetFecha: string) => {
    setDraggedReserva(null);
    setDragOverTarget(null);

    if (res.esBloqueo) return;

    const habOrigen = calData?.habitaciones.find(h => h.habitacionId === res.habitacionId);
    const habDestino = calData?.habitaciones.find(h => h.habitacionId === targetHabitacionId);

    const origenHabNum = habOrigen?.numero || res.habitacionNumero || '?';
    const destinoHabNum = habDestino?.numero || '?';

    const entradaOrig = parseLocal(res.fechaEntrada);
    const salidaOrig = parseLocal(res.fechaSalida);
    const duracionDias = Math.max(1, Math.round((salidaOrig.getTime() - entradaOrig.getTime()) / (1000 * 3600 * 24)));

    const nuevaEntrada = parseLocal(targetFecha);
    const nuevaSalida = addDays(nuevaEntrada, duracionDias);

    const nuevaEntradaStr = toDateStr(nuevaEntrada);
    const nuevaSalidaStr = toDateStr(nuevaSalida);

    if (res.habitacionId === targetHabitacionId && res.fechaEntrada === nuevaEntradaStr) {
      return;
    }

    const reservaTag = res.numeroOrden ? `#${res.numeroOrden}` : '';

    setMoverStatus({
      title: `Desplazando reserva ${reservaTag}...`,
      details: `Trasladando a Hab/Apto ${destinoHabNum} (${formatFechaCorta(nuevaEntradaStr)} - ${formatFechaCorta(nuevaSalidaStr)})`
    });

    try {
      await apiFetch(`/api/reservas/${res.id}/mover`, {
        method: 'POST',
        body: JSON.stringify({
          habitacionId: targetHabitacionId,
          fechaEntrada: nuevaEntradaStr,
          fechaSalida: nuevaSalidaStr,
        }),
      });

      setMoverStatus({
        title: `Reserva ${reservaTag} trasladada con éxito`,
        details: `De Hab ${origenHabNum} (${formatFechaDisplay(res.fechaEntrada)}) → Hab ${destinoHabNum} (${formatFechaDisplay(nuevaEntradaStr)})`
      });

      await cargarCalData();
      setTimeout(() => setMoverStatus(null), 5000);
    } catch (err: any) {
      setMoverStatus({
        title: `No se pudo trasladar la reserva ${reservaTag}`,
        details: err?.message || 'Conflicto de fechas o error de servidor.',
        isError: true
      });
      setTimeout(() => setMoverStatus(null), 6000);
    }
  };

  return {
    inicio, setInicio,
    dias, setDias,
    residenciaFiltro, setResidenciaFiltro,
    calData,
    festivos,
    residencias,
    loading,
    tooltip, setTooltip,
    reservaSeleccionada, setReservaSeleccionada,
    highlightReservaId, setHighlightReservaId,
    draggedReserva, setDraggedReserva,
    dragOverTarget, setDragOverTarget,
    moverStatus, setMoverStatus,
    tooltipTimer,
    cargarCalData,
    handleDropReserva,
  };
}
