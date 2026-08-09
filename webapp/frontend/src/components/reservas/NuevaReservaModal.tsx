import React, { useState, useEffect, useMemo } from 'react';
import { apiFetch } from '../../api';
import { Residencia, Habitacion, Huesped, Tarifa } from '../../types';
import { toDateStr, addDays, parseLocal } from '../../utils/dateUtils';
import { validarYCorregirDNI } from '../../utils/dniUtils';
import { toTitleCase, formatEmail } from '../../utils/textUtils';
import { CustomDateTimePicker } from '../common/CustomDateTimePicker';

export function NuevaReservaModal({
  habitacionIdInicial, fechaEntradaInicial, onSaved, onCancel, isSolicitud, itemToEdit
}: {
  habitacionIdInicial?: string; fechaEntradaInicial?: string;
  onSaved: () => void; onCancel: () => void;
  isSolicitud?: boolean;
  itemToEdit?: any;
}) {

  const normalizarEmpleo = (val?: string | null) => {
    if (!val) return 'GC';
    const v = val.trim().toUpperCase();
    if (v.includes('ALUMNO')) return 'ALUMNO GC';
    if (v.includes('FUNCIONARIO')) return 'FUNCIONARIO EN GC';
    if (v.includes('MILITAR EN')) return 'MILITAR EN GC';
    if (v.includes('MILITAR NO') || v.includes('MILITAR')) return 'MILITAR NO GC';
    return 'GC';
  };

  // 1. USESTATE HOOKS
  const [habitaciones, setHabitaciones] = useState<Habitacion[]>([]);
  const [residencias, setResidencias]   = useState<Residencia[]>([]);
  const [tarifas, setTarifas]           = useState<Tarifa[]>([]);
  const [residenciaSeleccionadaId, setResidenciaSeleccionadaId] = useState('');

  const tarifasFiltradas = useMemo(() => {
    return tarifas.filter(t => !residenciaSeleccionadaId || t.residenciaId === residenciaSeleccionadaId);
  }, [tarifas, residenciaSeleccionadaId]);

  const [buscarHuespedTexto, setBuscarHuespedTexto] = useState('');
  const [huespedesSugeridos, setHuespedesSugeridos] = useState<Huesped[]>([]);
  const [buscandoHuespedes, setBuscandoHuespedes] = useState(false);
  const [huesped, setHuesped]            = useState<Huesped | null>(null);
  const [huespedNuevo, setHuespedNuevo]  = useState(false);

  const [solapamiento, setSolapamiento]  = useState('');
  const [loading, setLoading]            = useState(false);

  const [showHabitacionesPopup, setShowHabitacionesPopup] = useState(false);
  const [habitacionesDisponiblesPopup, setHabitacionesDisponiblesPopup] = useState<Habitacion[]>([]);
  const [loadingHabitacionesPopup, setLoadingHabitacionesPopup] = useState(false);

  const [alojamientoCantidades, setAlojamientoCantidades] = useState<{ [key: string]: number }>({});
  const [selectedHabitacionIds, setSelectedHabitacionIds] = useState<string[]>(habitacionIdInicial ? [habitacionIdInicial] : []);

  const [huespedForm, setHuespedForm] = useState({
    dni: '', nombre: '', apellidos: '', telefono: '', email: '',
    direccion: '', codigoPostal: '', municipio: '', provincia: '',
    empleo: 'Guardia', situacion: 'Activo', tipoHuesped: 'Externo',
    centroOrigen: '', departamento: '', enListaNegra: false, motivoListaNegra: '', notas: '',
    familiaNumerosa: 'NO', porcentajeDescuento: 0
  });

  const [candidatosModal, setCandidatosModal] = useState<any[] | null>(null);
  const [showAjustesModal, setShowAjustesModal] = useState(false);
  const [edadMaximaNinos, setEdadMaximaNinos] = useState<number>(() => {
    const saved = localStorage.getItem('residencia_edad_maxima_ninos');
    return saved ? (parseInt(saved, 10) || 6) : 6;
  });
  const [tempEdadMaxima, setTempEdadMaxima] = useState<number>(edadMaximaNinos);

  const getLocalDateTimeString = (d = new Date()) => {
    const offset = d.getTimezoneOffset() * 60000;
    return new Date(d.getTime() - offset).toISOString().slice(0, 16);
  };

  const getInitialFechaSolicitud = () => {
    const saved = localStorage.getItem('lastSelectedFechaSolicitud');
    if (saved && !isNaN(new Date(saved).getTime())) {
      return saved;
    }
    return getLocalDateTimeString();
  };

  const manana = toDateStr(addDays(new Date(), 1));

  const [f, setF] = useState({
    habitacionId: habitacionIdInicial ?? '',
    huespedId: '',
    fechaEntrada: fechaEntradaInicial ?? toDateStr(new Date()),
    fechaSalida: fechaEntradaInicial ? toDateStr(addDays(parseLocal(fechaEntradaInicial), 1)) : manana,
    numPersonas: 1 as unknown as number,
    numNinos: 0 as number,
    camasSupletorias: 0,
    familiaNumerosa: 'NO',
    porcentajeDescuento: 0,
    alojamientoSolicitado: 'Cualquier alojamiento disponible',
    esBloqueo: false,
    motivoBloqueo: '',
    tarifaId: '',
    observaciones: '',
    finalidad: 'Otros',
    empleo: normalizarEmpleo(itemToEdit?.empleo),
    rango: itemToEdit?.huespedRango || itemToEdit?.rango || 'Guardia',
    situacion: 'Activo',
    resolucion: itemToEdit?.resolucion || 'SI',
    fechaSolicitud: getInitialFechaSolicitud()
  });

  // SETTERS
  const set = (k: string, v: any) => setF(p => ({ ...p, [k]: v }));
  const setH = (k: string, v: any) => setHuespedForm(p => ({ ...p, [k]: v }));

  // COMPUTED & MEMO
  const tiposHabitacionResidencia = useMemo(() => {
    if (!residenciaSeleccionadaId || !habitaciones || !Array.isArray(habitaciones)) return [];
    const habsRes = habitaciones.filter(h => h && h.residenciaId === residenciaSeleccionadaId);
    const map = new Map<string, { id: string; nombre: string; codigo: string; cama: string; capacidad: number }>();
    for (const h of habsRes) {
      const key = `${h.tipoHabitacionId}_${h.tipoCamaPrincipal || 'IND'}`;
      if (!map.has(key)) {
        map.set(key, {
          id: key,
          nombre: h.tipoNombre,
          codigo: h.tipoCodigo,
          cama: h.tipoCamaPrincipal || 'IND',
          capacidad: h.capacidadPersonas || 1
        });
      }
    }
    const getTipoHabOrden = (t: { nombre: string; cama: string; capacidad: number }) => {
      const nombreUpper = (t.nombre || '').toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      const camaUpper = (t.cama || '').toUpperCase();
      if (nombreUpper.includes('INDIVIDUAL') || (t.capacidad === 1 && camaUpper === 'IND')) return 1;
      if (nombreUpper.includes('DOBLE') && camaUpper !== 'MAT') return 2;
      if (nombreUpper.includes('MATRIMONIO') || (nombreUpper.includes('DOBLE') && camaUpper === 'MAT') || (camaUpper === 'MAT' && t.capacidad === 2)) return 3;
      if (nombreUpper.includes('TRIPLE') && camaUpper === 'MAT') return 4;
      if (nombreUpper.includes('TRIPLE') && camaUpper !== 'MAT') return 5;
      if (nombreUpper.includes('CUADRUPLE') || nombreUpper.includes('CUADRUP')) return 6;
      return 100 + (t.capacidad || 1);
    };
    return Array.from(map.values()).sort((a, b) => getTipoHabOrden(a) - getTipoHabOrden(b));
  }, [habitaciones, residenciaSeleccionadaId]);

  const paxTotalEstimado = useMemo(() => {
    let total = 0;
    tiposHabitacionResidencia.forEach((t: { id: string; nombre: string; codigo: string; cama: string; capacidad: number }) => {
      const qty = alojamientoCantidades[t.id] || 0;
      total += qty * t.capacidad;
    });
    return total;
  }, [alojamientoCantidades, tiposHabitacionResidencia]);

  const totalHabitacionesSolicitadas = useMemo(() => {
    return (Object.values(alojamientoCantidades) as number[]).reduce((a, b) => a + b, 0);
  }, [alojamientoCantidades]);

  const numAdultos = useMemo(() => {
    return Math.max(1, (+f.numPersonas || 1) - (+f.numNinos || 0));
  }, [f.numPersonas, f.numNinos]);

  const capacidadHabitacionesAsignadas = useMemo(() => {
    if (selectedHabitacionIds.length === 0) return paxTotalEstimado;
    let total = 0;
    selectedHabitacionIds.forEach(id => {
      const hab = habitaciones.find(h => h.id === id);
      if (hab) {
        total += (hab.capacidadPersonas ?? 1) + (hab.admiteSupletorias ? (hab.plazasSupletorias ?? 0) : 0);
      }
    });
    return total;
  }, [selectedHabitacionIds, habitaciones, paxTotalEstimado]);

  const haCambiadoDatosHuesped = useMemo(() => {
    if (!huesped) return false;
    return (
      (huespedForm.nombre?.trim() || '') !== (huesped.nombre || '') ||
      (huespedForm.apellidos?.trim() || '') !== (huesped.apellidos || '') ||
      (huespedForm.telefono?.trim() || '') !== (huesped.telefono || '') ||
      (huespedForm.email?.trim() || '') !== (huesped.email || '') ||
      (huespedForm.direccion?.trim() || '') !== (huesped.direccion || '') ||
      (huespedForm.codigoPostal?.trim() || '') !== (huesped.codigoPostal || '') ||
      (huespedForm.municipio?.trim() || '') !== (huesped.municipio || '') ||
      (huespedForm.provincia?.trim() || '') !== (huesped.provincia || '')
    );
  }, [huesped, huespedForm]);

  // EVENT HANDLERS & HELPERS
  const renderRoomLabel = (t: { nombre: string; cama: string; capacidad: number }) => {
    const nombreUpper = (t.nombre || '').toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const camaUpper = (t.cama || '').toUpperCase();

    if (nombreUpper.includes('INDIVIDUAL') || (t.capacidad === 1 && camaUpper === 'IND')) {
      return <strong style={{ fontSize: 11.5 }} title="Habitación Individual (1 plaza)">INDIVIDUAL</strong>;
    }
    if (nombreUpper.includes('DOBLE') && camaUpper !== 'MAT') {
      return <strong style={{ fontSize: 11.5 }} title="Habitación Doble (2 camas separadas / TWIN)">DOBLE</strong>;
    }
    if (nombreUpper.includes('MATRIMONIO') || (nombreUpper.includes('DOBLE') && camaUpper === 'MAT') || (camaUpper === 'MAT' && t.capacidad === 2)) {
      return <strong style={{ fontSize: 11.5 }} title="Habitación Matrimonio (1 cama de matrimonio)">MATRIMONIO</strong>;
    }
    if (nombreUpper.includes('TRIPLE') && camaUpper === 'MAT') {
      return <span title="Habitación Triple (1 cama matrimonio + 1 individual)"><strong style={{ fontSize: 11.5 }}>TRIPLE</strong><sub style={{ fontSize: '0.75em', fontWeight: 700, marginLeft: 1 }}>MAT</sub></span>;
    }
    if (nombreUpper.includes('TRIPLE') && camaUpper !== 'MAT') {
      return <span title="Habitación Triple (3 camas separadas / TWIN)"><strong style={{ fontSize: 11.5 }}>TRIPLE</strong><sub style={{ fontSize: '0.75em', fontWeight: 700, marginLeft: 1 }}>DOB</sub></span>;
    }
    if (nombreUpper.includes('CUADRUPLE') || nombreUpper.includes('CUADRUP')) {
      return <span title="Habitación Cuádruple (4 plazas / camas separadas / TWIN)"><strong style={{ fontSize: 11.5 }}>CUADRUPLE</strong><sub style={{ fontSize: '0.75em', fontWeight: 700, marginLeft: 1 }}>DOB</sub></span>;
    }

    return <span title={`${t.nombre} (${t.cama})`}><strong style={{ fontSize: 11.5 }}>{t.nombre.toUpperCase()}</strong>{t.cama ? <sub style={{ fontSize: '0.75em', marginLeft: 1 }}>{t.cama}</sub> : null}</span>;
  };

  const updateAlojamientoCantidad = (typeId: string, delta: number) => {
    const current = alojamientoCantidades[typeId] || 0;
    const nextVal = Math.max(0, current + delta);
    const nextMap = { ...alojamientoCantidades, [typeId]: nextVal };

    const newTotalHabitaciones = (Object.values(nextMap) as number[]).reduce((a, b) => a + b, 0);
    if (delta > 0 && newTotalHabitaciones > f.numPersonas) {
      alert(`Incongruencia de Habitaciones vs PAX:\nHas indicado PAX SOLICITUD = ${f.numPersonas}.\n\nEl número de habitaciones solicitadas (${newTotalHabitaciones}) no puede ser mayor que el número de personas (${f.numPersonas}).`);
      return;
    }

    setAlojamientoCantidades(nextMap);

    const summaryItems: string[] = [];
    tiposHabitacionResidencia.forEach((t: { id: string; nombre: string; codigo: string; cama: string; capacidad: number }) => {
      const qty = nextMap[t.id] || 0;
      if (qty > 0) {
        summaryItems.push(`${qty}x ${t.nombre} (${t.cama})`);
      }
    });
    const summaryText = summaryItems.length > 0 ? summaryItems.join(', ') : 'Cualquier alojamiento disponible';
    set('alojamientoSolicitado', summaryText);
  };

  useEffect(() => {
    Promise.all([
      apiFetch<Habitacion[]>('/api/habitaciones'),
      apiFetch<Residencia[]>('/api/residencias'),
      apiFetch<Tarifa[]>('/api/tarifas')
    ]).then(([h, r, t]) => {
      setHabitaciones(h);
      setResidencias(r);
      setTarifas(t);

      let initialResId = '';
      if (habitacionIdInicial) {
        const initialHab = h.find(x => x.id === habitacionIdInicial);
        if (initialHab) {
          initialResId = initialHab.residenciaId;
        }
      }

      if (!initialResId) {
        if (r.length === 1) {
          initialResId = r[0].id;
        } else if (r.length > 1) {
          const last = localStorage.getItem('lastSelectedResidenciaId');
          if (last && r.some(x => x.id === last)) {
            initialResId = last;
          } else {
            initialResId = r[0].id;
          }
        }
      }
      setResidenciaSeleccionadaId(initialResId);
    });
  }, []);

  // Precargar datos si se abre en modo edición (itemToEdit)
  useEffect(() => {
    if (!itemToEdit) return;

    const esNeg = ['NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'].includes((itemToEdit.resolucion || '').toUpperCase());
    const validHabId = esNeg ? '' : (itemToEdit.habitacionId || habitacionIdInicial || '');

    setF({
      habitacionId: validHabId,
      huespedId: itemToEdit.huespedId || '',
      fechaEntrada: itemToEdit.fechaEntrada || fechaEntradaInicial || '',
      fechaSalida: itemToEdit.fechaSalida || '',
      numPersonas: itemToEdit.numPersonas || 1,
      numNinos: itemToEdit.numNinos || 0,
      camasSupletorias: itemToEdit.camasSupletorias || 0,
      familiaNumerosa: itemToEdit.familiaNumerosa || 'NO',
      porcentajeDescuento: itemToEdit.porcentajeDescuento || 0,
      alojamientoSolicitado: itemToEdit.alojamientoSolicitado || 'Cualquier alojamiento disponible',
      esBloqueo: itemToEdit.esBloqueo || false,
      motivoBloqueo: itemToEdit.motivoBloqueo || '',
      tarifaId: itemToEdit.tarifaId || '',
      observaciones: itemToEdit.observaciones || '',
      finalidad: itemToEdit.finalidad || 'Otros',
      empleo: normalizarEmpleo(itemToEdit.empleo),
      rango: itemToEdit.huespedRango || itemToEdit.rango || 'Guardia',
      situacion: itemToEdit.huespedSituacion || 'Activo',
      resolucion: itemToEdit.resolucion || 'SI',
      fechaSolicitud: itemToEdit.fechaSolicitud ? itemToEdit.fechaSolicitud.slice(0, 16) : getLocalDateTimeString()
    });

    if (validHabId) {
      setSelectedHabitacionIds([validHabId]);
      const hab = habitaciones.find(x => x.id === validHabId);
      if (hab) setResidenciaSeleccionadaId(hab.residenciaId);
    } else {
      setSelectedHabitacionIds([]);
    }

    if (itemToEdit.huespedId) {
      apiFetch<Huesped>(`/api/huespedes/${itemToEdit.huespedId}`).then(h => {
        if (h) {
          setHuesped(h);
          setF(p => ({
            ...p,
            empleo: normalizarEmpleo(itemToEdit.empleo || h.empleoCategoria),
            rango: itemToEdit.huespedRango || h.empleo || 'Guardia',
            situacion: itemToEdit.huespedSituacion || h.situacion || 'Activo',
            finalidad: itemToEdit.finalidad || h.finalidad || 'Otros'
          }));
          setHuespedForm({
            dni: h.dni || '',
            nombre: h.nombre || '',
            apellidos: h.apellidos || '',
            telefono: h.telefono || '',
            email: h.email || '',
            direccion: h.direccion || '',
            codigoPostal: h.codigoPostal || '',
            municipio: h.municipio || '',
            provincia: h.provincia || '',
            empleo: h.empleo || itemToEdit.empleo || 'Guardia',
            situacion: h.situacion || 'Activo',
            tipoHuesped: h.tipoHuesped || 'Externo',
            centroOrigen: h.centroOrigen || '',
            departamento: h.departamento || '',
            enListaNegra: h.enListaNegra || false,
            motivoListaNegra: h.motivoListaNegra || '',
            notas: h.notas || '',
            familiaNumerosa: h.familiaNumerosa || 'NO',
            porcentajeDescuento: h.porcentajeDescuento || 0
          });
        }
      }).catch(() => {
        setHuespedForm(p => ({
          ...p,
          dni: itemToEdit.huespedDni || '',
          nombre: itemToEdit.huespedNombre || itemToEdit.huespedNombreCompleto || '',
          apellidos: itemToEdit.huespedApellidos || '',
          telefono: itemToEdit.huespedTelefono || '',
          email: itemToEdit.huespedEmail || ''
        }));
      });
    }
  }, [itemToEdit, habitaciones]);

  const handleResidenciaChange = (resId: string) => {
    setResidenciaSeleccionadaId(resId);
    localStorage.setItem('lastSelectedResidenciaId', resId);
    const currentHab = habitaciones.find(hab => hab.id === f.habitacionId);
    if (currentHab && currentHab.residenciaId !== resId) {
      set('habitacionId', '');
    }
  };

  const handleFamiliaNumerosaChange = (val: string) => {
    const pct = val === 'ESPECIAL' ? 50 : val === 'GENERAL' ? 20 : 0;
    setF(p => ({ ...p, familiaNumerosa: val, porcentajeDescuento: pct }));
    setHuespedForm(p => ({ ...p, familiaNumerosa: val, porcentajeDescuento: pct }));
  };

  // Búsqueda en tiempo real de huéspedes
  useEffect(() => {
    if (!buscarHuespedTexto || buscarHuespedTexto.trim().length < 2) {
      setHuespedesSugeridos([]);
      return;
    }
    const timer = setTimeout(async () => {
      setBuscandoHuespedes(true);
      try {
        const res = await apiFetch<any>(`/api/huespedes?buscar=${encodeURIComponent(buscarHuespedTexto.trim())}`);
        const list: Huesped[] = Array.isArray(res) ? res : (res?.items ?? []);
        setHuespedesSugeridos(list);
      } catch {
        setHuespedesSugeridos([]);
      } finally {
        setBuscandoHuespedes(false);
      }
    }, 300);
    return () => clearTimeout(timer);
  }, [buscarHuespedTexto]);

  const seleccionarHuesped = (h: Huesped) => {
    if (h.enListaNegra) { alert(`⚠️ ATENCIÓN: Este huésped está en LISTA NEGRA.\nMotivo: ${h.motivoListaNegra}`); }
    setHuesped(h);
    set('huespedId', h.id);
    if (h.empleo) set('rango', h.empleo);
    if (h.empleoCategoria) set('empleo', normalizarEmpleo(h.empleoCategoria));
    if (h.situacion) set('situacion', h.situacion);
    if (h.finalidad) set('finalidad', h.finalidad);
    if (h.familiaNumerosa) handleFamiliaNumerosaChange(h.familiaNumerosa);
    setHuespedForm({
      dni: h.dni || '',
      nombre: h.nombre || '',
      apellidos: h.apellidos || '',
      telefono: h.telefono || '',
      email: h.email || '',
      direccion: h.direccion || '',
      codigoPostal: h.codigoPostal || '',
      municipio: h.municipio || '',
      provincia: h.provincia || '',
      empleo: h.empleo || 'Guardia',
      situacion: h.situacion || 'Activo',
      tipoHuesped: h.tipoHuesped || 'Externo',
      centroOrigen: h.centroOrigen || '',
      departamento: h.departamento || '',
      enListaNegra: h.enListaNegra || false,
      motivoListaNegra: h.motivoListaNegra || '',
      notas: h.notas || '',
      familiaNumerosa: h.familiaNumerosa || 'NO',
      porcentajeDescuento: h.porcentajeDescuento || 0
    });
    setBuscarHuespedTexto('');
    setHuespedesSugeridos([]);
  };

  const handleBuscarKeyDown = async (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const textoBusqueda = buscarHuespedTexto.trim();
      if (!textoBusqueda) return;

      setBuscandoHuespedes(true);
      try {
        const res = await apiFetch<any>(`/api/huespedes?buscar=${encodeURIComponent(textoBusqueda)}`);
        const list: Huesped[] = Array.isArray(res) ? res : (res?.items ?? []);

        if (list.length === 1) {
          seleccionarHuesped(list[0]);
        } else if (list.length > 1) {
          setHuespedesSugeridos(list);
        } else {
          // No está en la BD -> abrir automáticamente el alta de nuevo huésped pre-rellenando los datos introducidos sin emergente
          let dniVal = '';
          let nombreVal = '';
          let apellidosVal = '';

          const hasDigits = /\d/.test(textoBusqueda);
          if (hasDigits) {
            dniVal = textoBusqueda.toUpperCase().replace(/\s+/g, '');
          } else {
            const parts = textoBusqueda.split(/\s+/);
            if (parts.length === 1) {
              nombreVal = parts[0];
            } else if (parts.length >= 2) {
              nombreVal = parts[0];
              apellidosVal = parts.slice(1).join(' ');
            }
          }

          setHuesped(null);
          set('huespedId', '');
          setHuespedNuevo(true);
          setHuespedForm({
            dni: dniVal,
            nombre: nombreVal,
            apellidos: apellidosVal,
            telefono: '', email: '',
            direccion: '', codigoPostal: '', municipio: '', provincia: '',
            empleo: f.empleo || 'Guardia', situacion: f.situacion || 'Activo', tipoHuesped: 'Externo',
            centroOrigen: '', departamento: '', enListaNegra: false, motivoListaNegra: '', notas: '',
            familiaNumerosa: f.familiaNumerosa || 'NO', porcentajeDescuento: f.porcentajeDescuento || 0
          });
          setBuscarHuespedTexto('');
          setHuespedesSugeridos([]);
        }
      } catch {
        setHuespedesSugeridos([]);
      } finally {
        setBuscandoHuespedes(false);
      }
    }
  };

  const abrirRegistroNuevoHuesped = () => {
    setHuesped(null);
    set('huespedId', '');
    setHuespedNuevo(true);

    let dniVal = '';
    let nombreVal = '';
    let apellidosVal = '';
    const txt = buscarHuespedTexto.trim();
    if (/^\d{7,8}[A-Za-z]?$/.test(txt)) {
      dniVal = txt.toUpperCase();
    } else if (txt) {
      const parts = txt.split(/\s+/);
      if (parts.length === 1) {
        nombreVal = parts[0];
      } else if (parts.length >= 2) {
        nombreVal = parts[0];
        apellidosVal = parts.slice(1).join(' ');
      }
    }

    setHuespedForm({
      dni: dniVal,
      nombre: nombreVal,
      apellidos: apellidosVal,
      telefono: '', email: '',
      direccion: '', codigoPostal: '', municipio: '', provincia: '',
      empleo: f.empleo || 'Guardia', situacion: f.situacion || 'Activo', tipoHuesped: 'Externo',
      centroOrigen: '', departamento: '', enListaNegra: false, motivoListaNegra: '', notas: '',
      familiaNumerosa: f.familiaNumerosa || 'NO', porcentajeDescuento: f.porcentajeDescuento || 0
    });
    setBuscarHuespedTexto('');
    setHuespedesSugeridos([]);
  };

  const handleBlurDniHuespedForm = async () => {
    const dniVal = huespedForm.dni?.trim();
    if (!dniVal || huesped) return;

    const { esValido, dniFinal } = validarYCorregirDNI(dniVal);
    if (dniFinal && dniFinal !== dniVal) {
      setH('dni', dniFinal);
    }
    if (!esValido) return;

    try {
      const h = await apiFetch<Huesped>(`/api/huespedes/dni/${encodeURIComponent(dniFinal || dniVal)}`);
      if (h) {
        seleccionarHuesped(h);
        setHuespedNuevo(false);
      }
    } catch { }
  };

  const buscarHabitacionesDisponiblesPopup = async () => {
    if (!residenciaSeleccionadaId || !f.fechaEntrada || !f.fechaSalida) {
      alert('Por favor selecciona Residencia, Fecha Entrada y Fecha Salida antes de buscar disponibilidad.');
      return;
    }
    setLoadingHabitacionesPopup(true);
    setShowHabitacionesPopup(true);
    try {
      const query = `/api/reservas/habitaciones-disponibles?residenciaId=${residenciaSeleccionadaId}&fechaEntrada=${f.fechaEntrada}&fechaSalida=${f.fechaSalida}&pax=${(f.numPersonas || 1) + (f.numNinos || 0)}`;
      const lib = await apiFetch<Habitacion[]>(query);
      setHabitacionesDisponiblesPopup(lib);
    } catch (e: any) {
      alert(e.message || 'Error al buscar habitaciones disponibles');
    } finally {
      setLoadingHabitacionesPopup(false);
    }
  };

  const toggleSeleccionHabitacionPopup = (habId: string) => {
    setSelectedHabitacionIds(prev => {
      if (prev.includes(habId)) {
        return prev.filter(x => x !== habId);
      } else {
        const maxPermitidas = totalHabitacionesSolicitadas > 0 ? totalHabitacionesSolicitadas : 1;
        if (prev.length >= maxPermitidas) {
          alert(`Atención: Solo puedes seleccionar hasta ${maxPermitidas} habitación(es) según las cantidades indicadas en ALOJAMIENTO SOLICITADO.`);
          return prev;
        }
        return [...prev, habId];
      }
    });
  };

  const confirmarSeleccionHabitacionesPopup = () => {
    if (selectedHabitacionIds.length > 0) {
      set('habitacionId', selectedHabitacionIds[0]);
    } else {
      set('habitacionId', '');
    }
    setShowHabitacionesPopup(false);
  };

  const handleCpChange = async (cpValue: string) => {
    setH('codigoPostal', cpValue);
    if (cpValue.length >= 4) {
      try {
        const res = await apiFetch<{ codigoPostal: string; municipio: string; provincia: string }[]>(`/api/direcciones/buscar?cp=${cpValue}`);
        if (res.length >= 1) {
          setHuespedForm(p => ({
            ...p,
            codigoPostal: res[0].codigoPostal,
            municipio: res[0].municipio,
            provincia: res[0].provincia
          }));
        }
      } catch { }
    }
  };

  const handleMunicipioChange = async (mValue: string) => {
    setH('municipio', mValue);
  };

  const ejecutarGuardadoEdicion = async (reevaluarCandId: string | null = null, candNombre: string | null = null) => {
    setLoading(true);
    try {
      if (huesped) {
        await apiFetch(`/api/huespedes/${huesped.id}`, {
          method: 'PUT',
          body: JSON.stringify({
            dni: huesped.dni,
            nombre: toTitleCase(huespedForm.nombre),
            apellidos: toTitleCase(huespedForm.apellidos),
            telefono: huespedForm.telefono?.trim() || null,
            email: formatEmail(huespedForm.email) || null,
            direccion: toTitleCase(huespedForm.direccion) || null,
            codigoPostal: huespedForm.codigoPostal?.trim() || null,
            municipio: toTitleCase(huespedForm.municipio) || null,
            provincia: toTitleCase(huespedForm.provincia) || null,
            empleo: f.rango || huesped.empleo || 'Guardia',
            empleoCategoria: f.empleo || huesped.empleoCategoria || 'GC',
            situacion: f.situacion || huesped.situacion || 'Activo',
            finalidad: f.finalidad || huesped.finalidad || 'Otros',
            familiaNumerosa: f.familiaNumerosa || huesped.familiaNumerosa || 'NO',
            tipoHuesped: huespedForm.tipoHuesped || huesped.tipoHuesped || 'Externo'
          })
        });
      }

      await apiFetch(`/api/reservas/${itemToEdit.id}`, {
        method: 'PUT',
        body: JSON.stringify({
          fechaEntrada: f.fechaEntrada,
          fechaSalida: f.fechaSalida,
          numPersonas: +f.numPersonas || 1,
          numNinos: +f.numNinos || 0,
          camasSupletorias: f.camasSupletorias || 0,
          familiaNumerosa: f.familiaNumerosa,
          porcentajeDescuento: f.porcentajeDescuento,
          alojamientoSolicitado: f.alojamientoSolicitado,
          estado: itemToEdit.estado || 'Confirmada',
          esBloqueo: false,
          tarifaId: f.tarifaId || null,
          observaciones: f.observaciones,
          pagado: itemToEdit.pagado || false,
          formaPago: itemToEdit.formaPago || null,
          finalidad: f.finalidad,
          empleo: f.empleo,
          rango: f.rango,
          situacion: f.situacion,
          resolucion: f.resolucion,
          habitacionId: selectedHabitacionIds[0] || f.habitacionId || itemToEdit.habitacionId || null,
          fechaSolicitud: f.fechaSolicitud || null,
          reevaluarCandidatoId: reevaluarCandId
        })
      });

      if (candNombre) {
        alert(`Se ha registrado la RENUNCIA de la reserva y se ha asignado y reevaluado la solicitud a favor de: ${candNombre}`);
      }

      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (loading) return;

    if (!itemToEdit && !f.huespedId && !huespedNuevo) {
      alert('Por favor selecciona un Huésped existente o haz clic en "+ Registrar Nuevo Huésped".');
      return;
    }

    if (itemToEdit) {
      if (f.resolucion === 'RENUNCIA' && candidatosModal === null) {
        try {
          const cands = await apiFetch<any[]>(`/api/reservas/${itemToEdit.id}/candidatos-reevaluacion`);
          if (cands && cands.length > 0) {
            setCandidatosModal(cands);
            return;
          }
        } catch {}
      }

      await ejecutarGuardadoEdicion(null, null);
      return;
    }

    // Validar campos obligatorios (*) al Crear Solicitud
    const camposFaltantes: string[] = [];

    if (!residenciaSeleccionadaId) camposFaltantes.push('RESIDENCIA');
    if (!f.fechaSolicitud?.trim()) camposFaltantes.push('FECHA / HORA SOLICITUD');
    if (!f.empleo) camposFaltantes.push('EMPLEO');
    if (!f.rango) camposFaltantes.push('RANGO');
    if (!f.situacion) camposFaltantes.push('SITUACIÓN');
    if (!f.finalidad) camposFaltantes.push('FINALIDAD / MOTIVO');

    if (!huesped && !huespedNuevo) {
      camposFaltantes.push('HUÉSPED SOLICITANTE (selecciona un huésped o pulsa + Registrar Nuevo Huésped)');
    } else {
      if (!huespedForm.dni?.trim() && !huesped?.dni) camposFaltantes.push('DNI');
      if (!huespedForm.nombre?.trim() && !huesped?.nombre) camposFaltantes.push('Nombre');
      if (!huespedForm.apellidos?.trim() && !huesped?.apellidos) camposFaltantes.push('Apellidos');
      if (!huespedForm.direccion?.trim() && !huesped?.direccion) camposFaltantes.push('DIRECCIÓN');
      if (!huespedForm.codigoPostal?.trim() && !huesped?.codigoPostal) camposFaltantes.push('C. Postal');
      if (!huespedForm.municipio?.trim() && !huesped?.municipio) camposFaltantes.push('Población');
      if (!huespedForm.provincia?.trim() && !huesped?.provincia) camposFaltantes.push('Provincia');
      if (!huespedForm.telefono?.trim() && !huesped?.telefono) camposFaltantes.push('TELÉFONO');
    }

    if (!f.numPersonas || +f.numPersonas < 1) camposFaltantes.push('PAX');
    if (!f.familiaNumerosa) camposFaltantes.push('DTO. F.N.');
    if (!f.fechaEntrada) camposFaltantes.push('FECHA ENTRADA');
    if (!f.fechaSalida) camposFaltantes.push('FECHA SALIDA');

    if (camposFaltantes.length > 0) {
      alert(`⚠️ No se puede crear la solicitud. Faltan campos obligatorios marcados con (*):\n\nPor favor, completa los siguientes campos antes de Crear Solicitud:\n\n• ${camposFaltantes.join('\n• ')}`);
      return;
    }

    if (f.numNinos && +f.numNinos >= +f.numPersonas) {
      alert(`Incongruencia de PAX vs NIÑOS:\nHas indicado PAX = ${f.numPersonas} y NIÑOS = ${f.numNinos}.\n\nEl número de niños debe ser estrictamente menor que el número total de personas para garantizar la presencia de al menos 1 adulto acompañante.`);
      return;
    }

    if (totalHabitacionesSolicitadas > f.numPersonas) {
      alert(`Incongruencia de Habitaciones vs PAX:\nHas solicitado ${totalHabitacionesSolicitadas} habitaciones para solo ${f.numPersonas} persona(s).\n\nEl número de habitaciones no puede ser mayor que el número de personas indicadas.`);
      return;
    }

    if (totalHabitacionesSolicitadas > 0 && capacidadHabitacionesAsignadas < numAdultos) {
      alert(`Capacidad insuficiente de adultos:\nLa capacidad del alojamiento asignado (${capacidadHabitacionesAsignadas} pax) es inferior a los ${numAdultos} adulto(s) indicados (${f.numPersonas} PAX TOTAL con ${f.numNinos || 0} niño(s)).\n\nPor favor, asigna habitaciones de mayor capacidad o reduce los adultos.`);
      return;
    }

    if (totalHabitacionesSolicitadas > 0 && selectedHabitacionIds.length > 0 && selectedHabitacionIds.length !== totalHabitacionesSolicitadas) {
      alert(`Atención: Has solicitado ${totalHabitacionesSolicitadas} habitación(es) (${f.alojamientoSolicitado}) pero se han seleccionado ${selectedHabitacionIds.length}.\n\nDebes asignar las ${totalHabitacionesSolicitadas} habitaciones requeridas o ajustar la cantidad solicitada en ALOJAMIENTO SOLICITADO.`);
      return;
    }

    setLoading(true);
    try {
      let actualHuespedId = f.huespedId;
      let datosHuespedActualizados = false;

      if (huespedNuevo) {
        const nuevoHuesped = await apiFetch<Huesped>('/api/huespedes', {
          method: 'POST',
          body: JSON.stringify({
            dni: (huespedForm.dni || '').trim().toUpperCase(),
            nombre: toTitleCase(huespedForm.nombre),
            apellidos: toTitleCase(huespedForm.apellidos),
            telefono: huespedForm.telefono?.trim() || null,
            email: formatEmail(huespedForm.email) || null,
            direccion: toTitleCase(huespedForm.direccion) || null,
            codigoPostal: huespedForm.codigoPostal?.trim() || null,
            municipio: toTitleCase(huespedForm.municipio) || null,
            provincia: toTitleCase(huespedForm.provincia) || null,
            centroOrigen: huespedForm.centroOrigen?.trim() || null,
            departamento: huespedForm.departamento?.trim() || null,
            motivoListaNegra: huespedForm.motivoListaNegra?.trim() || null,
            notas: huespedForm.notas?.trim() || null,
            tipoHuesped: huespedForm.tipoHuesped || 'Externo',
            enListaNegra: !!huespedForm.enListaNegra,
            empleo: f.rango || null,
            empleoCategoria: f.empleo || null,
            situacion: f.situacion || null,
            finalidad: f.finalidad || null,
            familiaNumerosa: f.familiaNumerosa || 'NO',
            porcentajeDescuento: f.porcentajeDescuento ?? 0
          })
        });
        actualHuespedId = nuevoHuesped.id;
      } else if (huesped && haCambiadoDatosHuesped) {
        await apiFetch(`/api/huespedes/${huesped.id}`, {
          method: 'PUT',
          body: JSON.stringify({
            dni: huesped.dni,
            nombre: huespedForm.nombre.trim(),
            apellidos: huespedForm.apellidos.trim(),
            telefono: huespedForm.telefono?.trim() || null,
            email: huespedForm.email?.trim() || null,
            direccion: huespedForm.direccion?.trim() || null,
            codigoPostal: huespedForm.codigoPostal?.trim() || null,
            municipio: huespedForm.municipio?.trim() || null,
            provincia: huespedForm.provincia?.trim() || null,
            empleo: f.rango || huesped.empleo,
            empleoCategoria: f.empleo || huesped.empleoCategoria,
            situacion: f.situacion || huesped.situacion,
            finalidad: f.finalidad || huesped.finalidad,
            familiaNumerosa: f.familiaNumerosa || huesped.familiaNumerosa,
            tipoHuesped: huespedForm.tipoHuesped || huesped.tipoHuesped || 'Externo'
          })
        });
        datosHuespedActualizados = true;
      }

      const uniqueSelectedIds = Array.from(new Set(selectedHabitacionIds));
      const habIdsToCreate = uniqueSelectedIds.length > 0 ? uniqueSelectedIds : (f.habitacionId ? [f.habitacionId] : [null]);
      let batchNumeroOrden: number | null = null;

      for (const habId of habIdsToCreate) {
        const res: any = await apiFetch<any>('/api/reservas', {
          method: 'POST',
          body: JSON.stringify({
            ...f,
            resolucion: null,
            numeroOrden: batchNumeroOrden,
            numPersonas: +f.numPersonas || 1,
            numNinos: +f.numNinos || 0,
            fechaSolicitud: f.fechaSolicitud || null,
            habitacionId: habId,
            huespedId: actualHuespedId || null,
            tarifaId: f.tarifaId || null,
          })
        });

        if (res?.reserva?.numeroOrden && !batchNumeroOrden) {
          batchNumeroOrden = res.reserva.numeroOrden;
        }
      }

      if (datosHuespedActualizados) {
        alert('✓ Datos del huésped actualizados correctamente.');
      }

      onSaved();
    } catch (e: any) { alert(e.message); } finally { setLoading(false); }
  };

  const listaEmpleos = [
    'GC', 'ALUMNO GC', 'FUNCIONARIO EN GC', 'MILITAR EN GC', 'MILITAR NO GC'
  ];

  const listaRangos = [
    'Guardia', 'Cabo', 'Cabo 1º', 'Cabo Mayor', 'Sargento', 'Sargento 1º', 'Brigada',
    'Subteniente', 'Suboficial Mayor', 'Alférez', 'Teniente', 'Capitán', 'Comandante',
    'Tte. Coronel', 'Coronel', 'General B', 'General D', 'Funcionario', 'Otro'
  ];

  const listaSituaciones = [
    'Activo', 'Viogen', 'Asociación', 'Reserva activo', 'Reserva', 'Excedencia',
    'Especiales', 'Retirado', 'Viuda', 'Huerfano', 'Estudiante'
  ];

  const listaFinalidades = [
    'Otros', 'Comisión', 'Comisión NO indem.', 'Destino', 'Enfermedad', 'Urgencia', 'Sepelio', 'Máx. Estancia'
  ];

  const listaResoluciones = [
    'SI', 'CONCEDIDA', 'REEVALUADA', 'NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'
  ];

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onCancel()}>
      <div className="modal" style={{ maxWidth: 860, maxHeight: '96vh', overflowY: 'auto' }}>
        <div className="modal-header" style={{ padding: '10px 18px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <h3 style={{ fontSize: 15, margin: 0 }}>{itemToEdit ? "📝 Editar Solicitud" : (isSolicitud ? "➕ Nueva solicitud" : "➕ Nueva Reserva / Solicitud")}</h3>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              style={{ padding: '2px 8px', fontSize: 13, border: '1px solid var(--border)' }}
              title="Configuración de parámetros del formulario (ej. Edad máxima de niños)"
              onClick={() => {
                setTempEdadMaxima(edadMaximaNinos);
                setShowAjustesModal(true);
              }}
            >
              ⚙️
            </button>
            <button className="modal-close" onClick={onCancel}>×</button>
          </div>
        </div>

        <form onSubmit={submit}>
          <div className="modal-body modal-body-compact" style={{ padding: '10px 18px', display: 'flex', flexDirection: 'column', gap: 7 }}>
            {solapamiento && <div className="alert alert-danger" style={{ padding: '6px 10px', fontSize: 11.5, marginBottom: 4 }}>⚠️ {solapamiento}</div>}

            {/* AVISOS DE VALIDACIÓN RÁPIDOS */}
            {totalHabitacionesSolicitadas > f.numPersonas && (
              <div className="alert alert-danger" style={{ fontSize: 12, padding: '8px 12px' }}>
                ⚠️ <strong>Incongruencia en la solicitud:</strong> Has solicitado {totalHabitacionesSolicitadas} habitaciones para solo {f.numPersonas} persona(s). El número de habitaciones no puede ser mayor que el número de personas.
              </div>
            )}

            {totalHabitacionesSolicitadas > 0 && capacidadHabitacionesAsignadas < numAdultos && (
              <div className="alert alert-danger" style={{ fontSize: 11.5, padding: '6px 10px' }}>
                ⚠️ <strong>Capacidad Insuficiente:</strong> La capacidad del alojamiento asignado ({capacidadHabitacionesAsignadas} pax) es inferior a los {numAdultos} adulto(s) indicados.
              </div>
            )}

            {totalHabitacionesSolicitadas > 0 && capacidadHabitacionesAsignadas >= numAdultos && capacidadHabitacionesAsignadas < f.numPersonas && (
              <div className="alert alert-info" style={{ fontSize: 11.5, padding: '6px 10px' }}>
                ℹ️ <strong>Co-Alojamiento con Menores:</strong> Solicitud de {f.numPersonas} pax ({numAdultos} adulto(s) + {f.numNinos} niño(s)) en habitación de {capacidadHabitacionesAsignadas} pax (niños comparten habitación con los adultos).
              </div>
            )}

            {itemToEdit && (
              <div style={{
                background: ['NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'].includes((f.resolucion || '').toUpperCase())
                  ? 'rgba(217, 83, 79, 0.08)'
                  : 'rgba(45, 138, 78, 0.08)',
                border: ['NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'].includes((f.resolucion || '').toUpperCase())
                  ? '1px solid rgba(217, 83, 79, 0.25)'
                  : '1px solid rgba(45, 138, 78, 0.25)',
                borderRadius: 8,
                padding: '10px 14px',
                fontSize: 12.5,
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                gap: 12
              }}>
                <div>
                  {['NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'].includes((f.resolucion || '').toUpperCase()) ? (
                    <>
                      <strong style={{ color: 'var(--danger, #d9534f)' }}>📋 Estado Alojamiento:</strong>{' '}
                      <strong>Sin habitación adjudicada (Solicitud DENEGADA / Lista de Espera)</strong>
                      <div style={{ fontSize: 11.5, color: 'var(--text-muted)', marginTop: 2 }}>
                        📅 Fechas solicitadas: del <strong>{f.fechaEntrada || '—'}</strong> al <strong>{f.fechaSalida || '—'}</strong> ({itemToEdit.totalNoches || 0} noches)
                      </div>
                    </>
                  ) : (
                    <>
                      <strong style={{ color: 'var(--primary)' }}>🏠 Alojamiento Adjudicado:</strong>{' '}
                      <strong>{selectedHabitacionIds.length > 0 ? `Habitaciones seleccionadas (${selectedHabitacionIds.length})` : (f.habitacionId ? `Hab. ${itemToEdit.habitacionNumero || ''}` : 'Sin habitación asignada')}</strong>{' '}
                      <span style={{ opacity: 0.8 }}>({itemToEdit.residenciaNombre || 'Residencia'})</span>
                      <div style={{ fontSize: 11.5, color: 'var(--text-muted)', marginTop: 2 }}>
                        📅 Reserva del <strong>{f.fechaEntrada}</strong> al <strong>{f.fechaSalida}</strong> ({itemToEdit.totalNoches || 0} noches)
                      </div>
                    </>
                  )}
                </div>
                <span className={`badge ${['SI', 'CONCEDIDA'].includes((f.resolucion || '').toUpperCase()) ? 'badge-success' : f.resolucion === 'RENUNCIA' ? 'badge-warning' : 'badge-danger'}`} style={{ fontSize: 11 }}>
                  Resolución: {f.resolucion}
                </span>
              </div>
            )}

            {/* FILA 1: RESIDENCIA | FECHA/HORA SOLICITUD | RESOLUCIÓN (en edición) */}
            <div style={{ display: 'grid', gridTemplateColumns: itemToEdit ? '1.2fr 1.2fr 1fr' : '1fr 1fr', gap: 16 }}>
              <div className="form-group">
                <label>RESIDENCIA *</label>
                <select value={residenciaSeleccionadaId} onChange={e => handleResidenciaChange(e.target.value)} required>
                  <option value="">— Seleccione Residencia —</option>
                  {residencias.map(r => (
                    <option key={r.id} value={r.id}>{r.nombre}</option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label>FECHA / HORA SOLICITUD *</label>
                <CustomDateTimePicker value={f.fechaSolicitud} onChange={val => set('fechaSolicitud', val)} />
              </div>
              {itemToEdit && (
                <div className="form-group">
                  <label>RESOLUCIÓN *</label>
                  <select value={f.resolucion} onChange={e => {
                    const newRes = e.target.value;
                    set('resolucion', newRes);
                    if (['NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'].includes(newRes.toUpperCase())) {
                      set('habitacionId', '');
                      setSelectedHabitacionIds([]);
                      setSolapamiento('');
                    }
                  }} required>
                    {listaResoluciones.map(res => (
                      <option key={res} value={res}>{res}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>

            {/* FILA 2: PAX TOTAL | NIÑOS | FECHA ENTRADA | FECHA SALIDA */}
            <div style={{ display: 'grid', gridTemplateColumns: '95px 85px 1fr 1fr', gap: 12 }}>
              <div className="form-group">
                <label title="Número total de personas alojadas (Adultos + Niños)">
                  PAX<sub style={{ fontSize: '0.65em', fontWeight: 800, marginLeft: 2 }}>
                    <span style={{ padding: '1px 4px', background: 'rgba(30,58,95,0.1)', borderRadius: 3, border: '1px solid rgba(30,58,95,0.3)', color: 'var(--primary)' }}>
                      TOTAL
                    </span>
                  </sub> *
                </label>
                <input
                  type="number"
                  min={1}
                  max={9}
                  value={f.numPersonas}
                  onChange={e => {
                    const newPax = e.target.value === '' ? '' : +e.target.value;
                    set('numPersonas', newPax);
                    if (typeof newPax === 'number' && f.numNinos >= newPax) {
                      set('numNinos', Math.max(0, newPax - 1));
                    }
                  }}
                  required
                />
              </div>
              <div className="form-group">
                <label title={`Número de menores de edad (niños < ${edadMaximaNinos} años) incluidos en el total de PAX TOTAL. Haz clic en el botón ⚙️ arriba para configurar.`}>
                  NIÑOS<sub style={{ fontSize: '0.65em', fontWeight: 800, marginLeft: 2 }}>
                    <span style={{ padding: '1px 4px', background: 'rgba(30,58,95,0.1)', borderRadius: 3, border: '1px solid rgba(30,58,95,0.3)', color: 'var(--primary)' }}>
                      &lt; {edadMaximaNinos} A.
                    </span>
                  </sub>
                </label>
                <input
                  type="number"
                  min={0}
                  max={f.numPersonas ? Math.max(0, +f.numPersonas - 1) : 0}
                  value={f.numNinos}
                  onChange={e => {
                    const val = e.target.value === '' ? 0 : +e.target.value;
                    const maxPermitido = f.numPersonas ? Math.max(0, +f.numPersonas - 1) : 0;
                    if (f.numPersonas && val >= +f.numPersonas) {
                      alert(`Incongruencia de PAX vs NIÑOS:\nEl número de niños (${val}) debe ser estrictamente MENOR que el total de PAX (${f.numPersonas}).\n\nDebe haber al menos 1 adulto responsable de la reserva (PAX = Nº ADULTOS + Nº NIÑOS).`);
                      set('numNinos', maxPermitido);
                    } else {
                      set('numNinos', val);
                    }
                  }}
                  placeholder="0"
                />
              </div>
              <div className="form-group">
                <label>FECHA ENTRADA *</label>
                <input type="date" value={f.fechaEntrada} onChange={e => set('fechaEntrada', e.target.value)} required />
              </div>
              <div className="form-group">
                <label>FECHA SALIDA *</label>
                <input type="date" value={f.fechaSalida} min={f.fechaEntrada} onChange={e => set('fechaSalida', e.target.value)} required />
              </div>
            </div>

            {/* BLOQUE HUÉSPED SOLICITANTE */}
            <div style={{ background: 'var(--surface-2)', padding: '8px 12px', borderRadius: 8, border: '1px solid var(--border)', display: 'flex', flexDirection: 'column', gap: 7 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <label style={{ fontWeight: 700, fontSize: 12, color: 'var(--primary)', margin: 0 }}>👤 HUÉSPED SOLICITANTE *</label>
                {!itemToEdit && !huespedNuevo && huesped && (
                  <button type="button" className="btn btn-ghost btn-sm" onClick={() => { setHuesped(null); set('huespedId', ''); setBuscarHuespedTexto(''); }}>✕ Cambiar Huésped</button>
                )}
              </div>

              {/* FILA: EMPLEO | RANGO | SITUACIÓN | FINALIDAD */}
              <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1.2fr 1fr 1.2fr', gap: 10 }}>
                <div className="form-group">
                  <label>EMPLEO *</label>
                  <select value={f.empleo} onChange={e => set('empleo', e.target.value)} required>
                    {listaEmpleos.map(emp => (
                      <option key={emp} value={emp}>{emp}</option>
                    ))}
                  </select>
                </div>
                <div className="form-group">
                  <label>RANGO *</label>
                  <select value={f.rango} onChange={e => set('rango', e.target.value)} required>
                    {listaRangos.map(r => (
                      <option key={r} value={r}>{r}</option>
                    ))}
                  </select>
                </div>
                <div className="form-group">
                  <label>SITUACIÓN *</label>
                  <select value={f.situacion} onChange={e => set('situacion', e.target.value)} required>
                    {listaSituaciones.map(sit => (
                      <option key={sit} value={sit}>{sit}</option>
                    ))}
                  </select>
                </div>
                <div className="form-group">
                  <label>FINALIDAD / MOTIVO *</label>
                  <select value={f.finalidad} onChange={e => set('finalidad', e.target.value)} required>
                    {listaFinalidades.map(fin => (
                      <option key={fin} value={fin}>{fin}</option>
                    ))}
                  </select>
                </div>
              </div>

              {!itemToEdit && !huesped && !huespedNuevo && (
                <div style={{ position: 'relative' }}>
                  <div className="search-box" style={{ position: 'relative' }}>
                    <span className="search-icon">🔍</span>
                    <input
                      placeholder="Busca por DNI, Nombre y/o Apellidos"
                      value={buscarHuespedTexto}
                      onChange={e => setBuscarHuespedTexto(e.target.value)}
                      onKeyDown={handleBuscarKeyDown}
                    />
                    {buscandoHuespedes && (
                      <span style={{ position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)', fontSize: 11, color: 'var(--text-muted)', pointerEvents: 'none' }}>
                        ⏳ Buscando...
                      </span>
                    )}
                  </div>

                  {huespedesSugeridos.length > 0 && (
                    <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: '#ffffff', border: '1px solid var(--border)', borderRadius: 8, zIndex: 1000, maxHeight: 200, overflowY: 'auto', boxShadow: '0 8px 24px rgba(0,0,0,0.25)', marginTop: 4 }}>
                      {huespedesSugeridos.map(h => (
                        <div key={h.id} style={{ padding: '9px 12px', cursor: 'pointer', borderBottom: '1px solid var(--border-light)', fontSize: 12, background: '#ffffff', color: 'var(--text)' }}
                          onMouseDown={e => e.preventDefault()}
                          onClick={() => seleccionarHuesped(h)}>
                          <strong>{h.nombreCompleto}</strong> — <span className="badge badge-secondary">{h.dni}</span> <span className="text-muted">({h.empleo || 'Sin Rango'} / {h.situacion || 'Sin Situación'})</span>
                        </div>
                      ))}
                    </div>
                  )}

                  <div style={{ marginTop: 10, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: 11, color: 'var(--text-muted)' }}>Si el huésped no está registrado, puedes darlo de alta en el acto:</span>
                    <button type="button" className="btn btn-ghost btn-sm" onClick={() => abrirRegistroNuevoHuesped()}>+ Registrar Nuevo Huésped</button>
                  </div>
                </div>
              )}

              {huesped && (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 12px', background: 'rgba(45,138,78,0.1)', border: '1px solid rgba(45,138,78,0.3)', borderRadius: 6, fontSize: 12 }}>
                  <div>
                    <strong style={{ fontSize: 13 }}>{huespedForm.nombre} {huespedForm.apellidos}</strong>
                    <span style={{ marginLeft: 10, color: 'var(--text-muted)' }}>DNI: <strong>{huesped.dni}</strong></span>
                  </div>
                  <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                    {haCambiadoDatosHuesped && (
                      <span className="badge badge-warning">✏️ Datos modificados</span>
                    )}
                    <span className="badge badge-success">✓ Seleccionado</span>
                  </div>
                </div>
              )}

              {(itemToEdit || huespedNuevo || huesped) && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10, fontSize: 12 }}>
                  {!itemToEdit && huespedNuevo && (
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 2 }}>
                      <span style={{ fontSize: 12, fontWeight: 700, color: 'var(--warning)' }}>📝 Registrando Nuevo Huésped:</span>
                      <button type="button" className="btn btn-ghost btn-sm" style={{ padding: '2px 8px' }} onClick={() => setHuespedNuevo(false)}>✕ Cancelar alta nuevo</button>
                    </div>
                  )}

                  {/* DNI - NOMBRE - APELLIDOS - DTO. F.N. */}
                  <div style={{ display: 'grid', gridTemplateColumns: '100px 1.2fr 1.6fr 100px', gap: 10 }}>
                    <div className="form-group">
                      <label>DNI *</label>
                      <input
                        value={huespedForm.dni}
                        maxLength={9}
                        onChange={e => setH('dni', e.target.value.toUpperCase())}
                        onBlur={handleBlurDniHuespedForm}
                        placeholder="12345678Z"
                        required={!huesped}
                        readOnly={!huespedNuevo && !!huesped}
                      />
                    </div>
                    <div className="form-group">
                      <label>Nombre *</label>
                      <input value={huespedForm.nombre} onChange={e => setH('nombre', e.target.value)} onBlur={e => setH('nombre', toTitleCase(e.target.value))} spellCheck={true} lang="es" placeholder="Nombre" required={!huesped} />
                    </div>
                    <div className="form-group">
                      <label>Apellidos *</label>
                      <input value={huespedForm.apellidos} onChange={e => setH('apellidos', e.target.value)} onBlur={e => setH('apellidos', toTitleCase(e.target.value))} spellCheck={true} lang="es" placeholder="Apellidos" required={!huesped} />
                    </div>
                    <div className="form-group">
                      <label>DTO. F.N. *</label>
                      <select value={f.familiaNumerosa} onChange={e => handleFamiliaNumerosaChange(e.target.value)} required>
                        <option value="NO">NO (0%)</option>
                        <option value="GENERAL">General (20%)</option>
                        <option value="ESPECIAL">Especial (50%)</option>
                      </select>
                    </div>
                  </div>

                  {/* DIRECCIÓN - C. POSTAL - POBLACIÓN - PROVINCIA */}
                  <div style={{ display: 'grid', gridTemplateColumns: '2.5fr 80px 1.5fr 1.2fr', gap: 10 }}>
                    <div className="form-group">
                      <label>DIRECCIÓN *</label>
                      <input value={huespedForm.direccion} onChange={e => setH('direccion', e.target.value)} onBlur={e => setH('direccion', toTitleCase(e.target.value))} spellCheck={true} lang="es" placeholder="Calle / Avda..." required />
                    </div>
                    <div className="form-group" style={{ position: 'relative' }}>
                      <label>C. Postal *</label>
                      <input value={huespedForm.codigoPostal} maxLength={6} onChange={e => handleCpChange(e.target.value)} required />
                    </div>
                    <div className="form-group" style={{ position: 'relative' }}>
                      <label>Población *</label>
                      <input value={huespedForm.municipio} onChange={e => handleMunicipioChange(e.target.value)} onBlur={e => setH('municipio', toTitleCase(e.target.value))} spellCheck={true} lang="es" required />
                    </div>
                    <div className="form-group">
                      <label>Provincia *</label>
                      <input value={huespedForm.provincia} onChange={e => setH('provincia', e.target.value)} onBlur={e => setH('provincia', toTitleCase(e.target.value))} spellCheck={true} lang="es" required />
                    </div>
                  </div>

                  {/* TELÉFONO - EMAIL */}
                  <div style={{ display: 'grid', gridTemplateColumns: '100px 1.5fr', gap: 10 }}>
                    <div className="form-group">
                      <label>TELÉFONO *</label>
                      <input value={huespedForm.telefono} maxLength={9} onChange={e => setH('telefono', e.target.value)} placeholder="612345678" required />
                    </div>
                    <div className="form-group">
                      <label>EMAIL</label>
                      <input type="email" value={huespedForm.email} onChange={e => setH('email', e.target.value)} onBlur={e => setH('email', formatEmail(e.target.value))} placeholder="ejemplo@correo.es" />
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* ALOJAMIENTO SOLICITADO */}
            <div className="form-group" style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 6 }}>
                <label style={{ margin: 0 }}>ALOJAMIENTO SOLICITADO *</label>
                <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                  <span style={{
                    fontSize: 11, fontWeight: 700,
                    color: capacidadHabitacionesAsignadas < f.numPersonas ? 'var(--danger)' : 'var(--primary)',
                    background: capacidadHabitacionesAsignadas < f.numPersonas ? 'rgba(220,53,69,0.12)' : 'rgba(45,138,78,0.12)',
                    border: capacidadHabitacionesAsignadas < f.numPersonas ? '1px solid rgba(220,53,69,0.3)' : '1px solid rgba(45,138,78,0.3)',
                    padding: '2px 10px', borderRadius: 12
                  }}>
                    👥 Capacidad: {capacidadHabitacionesAsignadas} pax (PAX: {f.numPersonas})
                  </span>
                  {totalHabitacionesSolicitadas > 0 && (
                    <span style={{
                      fontSize: 11, fontWeight: 700,
                      color: totalHabitacionesSolicitadas > f.numPersonas ? 'var(--danger)' : 'var(--primary)',
                      background: totalHabitacionesSolicitadas > f.numPersonas ? 'rgba(220,53,69,0.12)' : 'rgba(45,138,78,0.12)',
                      border: totalHabitacionesSolicitadas > f.numPersonas ? '1px solid rgba(220,53,69,0.3)' : '1px solid rgba(45,138,78,0.3)',
                      padding: '2px 10px', borderRadius: 12
                    }}>
                      🛏️ Habitaciones: {totalHabitacionesSolicitadas}
                    </span>
                  )}
                </div>
              </div>
              {!residenciaSeleccionadaId ? (
                <div style={{ fontSize: 12, color: 'var(--text-muted)', fontStyle: 'italic', padding: 8, background: 'var(--surface-2)', borderRadius: 6 }}>
                  (Seleccione primero una Residencia arriba)
                </div>
              ) : (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 8, background: 'var(--surface-2)', padding: '8px 10px', borderRadius: 8, border: '1px solid var(--border)' }}>
                  {tiposHabitacionResidencia.map((t: { id: string; nombre: string; codigo: string; cama: string; capacidad: number }) => {
                    const qty = alojamientoCantidades[t.id] || 0;
                    return (
                      <div
                        key={t.id}
                        onClick={() => updateAlojamientoCantidad(t.id, 1)}
                        style={{
                          display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 6, padding: '6px 10px',
                          background: qty > 0 ? 'rgba(45,138,78,0.12)' : 'var(--bg)',
                          border: qty > 0 ? '1px solid var(--primary)' : '1px solid var(--border)',
                          borderRadius: 6, fontSize: 11.5, cursor: 'pointer', userSelect: 'none'
                        }}>
                        <span style={{ whiteSpace: 'nowrap', fontSize: 11.5, fontWeight: 600 }}>
                          {renderRoomLabel(t)}
                        </span>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 3, flexShrink: 0 }} onClick={e => e.stopPropagation()}>
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ padding: '0px 5px', fontSize: 11, minWidth: 20, height: 22, lineHeight: '1', borderRadius: 4 }}
                            onClick={(e) => { e.stopPropagation(); updateAlojamientoCantidad(t.id, -1); }}
                            disabled={qty <= 0}>-</button>
                          <span style={{ fontWeight: 700, minWidth: 14, textAlign: 'center', fontSize: 11.5 }}>{qty}</span>
                          <button
                            type="button"
                            className="btn btn-secondary"
                            style={{ padding: '0px 5px', fontSize: 11, minWidth: 20, height: 22, lineHeight: '1', borderRadius: 4 }}
                            onClick={(e) => { e.stopPropagation(); updateAlojamientoCantidad(t.id, 1); }}>+</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* SELECCIÓN / ASISTENTE DE HABITACIÓN */}
            <div style={{ background: 'var(--surface-1)', padding: '8px 12px', borderRadius: 8, border: '1px dashed var(--primary)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  <label style={{ fontWeight: 700, fontSize: 12 }}>🛏️ ASIGNACIÓN DE HABITACIÓN / APARTAMENTO</label>
                  {selectedHabitacionIds.length > 0 ? (
                    <div style={{ fontSize: 11.5, marginTop: 2, color: selectedHabitacionIds.length === totalHabitacionesSolicitadas ? 'var(--primary)' : 'var(--warning)', fontWeight: 600 }}>
                      {selectedHabitacionIds.length === totalHabitacionesSolicitadas ? '✓' : '⚠️'} {selectedHabitacionIds.length} de {totalHabitacionesSolicitadas} habitaciones asignadas:{' '}
                      {selectedHabitacionIds.map(id => {
                        const hab = habitaciones.find(h => h.id === id);
                        return hab ? `Hab. ${hab.numero} (${hab.tipoNombre} - ${hab.tipoCamaPrincipal})` : id;
                      }).join(' · ')}
                    </div>
                  ) : (
                    <div style={{ fontSize: 10.5, color: 'var(--text-muted)', marginTop: 2 }}>
                      {totalHabitacionesSolicitadas > 0
                        ? `⚠️ Pendiente de asignar ${totalHabitacionesSolicitadas} habitación(es) solicitada(s). Haz clic en "Buscar Disponibles".`
                        : 'Puedes seleccionar una habitación libre o buscar las disponibles en esas fechas y PAX.'}
                    </div>
                  )}
                </div>
                <button type="button" className="btn btn-primary btn-sm" onClick={buscarHabitacionesDisponiblesPopup}>
                  🔍 Buscar Disponibles
                </button>
              </div>
            </div>

            {/* TARIFA Y OBSERVACIONES */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', gap: 12 }}>
              <div className="form-group">
                <label>TARIFA APLICABLE</label>
                <select value={f.tarifaId} onChange={e => set('tarifaId', e.target.value)}>
                  <option value="">— Tarifa estándar / Ninguna —</option>
                  {tarifasFiltradas.map(t => (
                    <option key={t.id} value={t.id}>{t.nombreTarifa} ({t.precioNoche}€/noche)</option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label>OBSERVACIONES</label>
                <textarea value={f.observaciones} onChange={e => set('observaciones', e.target.value)} rows={2} placeholder="Peticiones especiales, notas del huésped o incidencias..." />
              </div>
            </div>
          </div>

          <div className="modal-footer" style={{ padding: '8px 18px' }}>
            <button type="button" className="btn btn-ghost btn-sm" onClick={onCancel}>Cancelar</button>
            <button type="submit" className="btn btn-primary btn-sm" disabled={loading || !!solapamiento || (!itemToEdit && !f.huespedId && !huespedNuevo)}>
              {loading ? 'Guardando...' : (itemToEdit ? '💾 Guardar cambios' : '💾 Crear Solicitud')}
            </button>
          </div>
        </form>
      </div>

      {/* POPUP DE AJUSTES Y CONFIGURACIÓN */}
      {showAjustesModal && (
        <div className="modal-overlay" style={{ zIndex: 1200 }} onClick={e => e.target === e.currentTarget && setShowAjustesModal(false)}>
          <div className="modal" style={{ maxWidth: 360, borderRadius: 12, boxShadow: '0 10px 30px rgba(0,0,0,0.3)' }}>
            <div className="modal-header" style={{ padding: '12px 18px', background: 'var(--surface-2)' }}>
              <h3 style={{ fontSize: 14, display: 'flex', alignItems: 'center', gap: 6, margin: 0, color: 'var(--primary)' }}>
                ⚙️ Ajustes del Formulario
              </h3>
              <button className="modal-close" onClick={() => setShowAjustesModal(false)}>×</button>
            </div>
            <div className="modal-body" style={{ padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: 14 }}>
              <div className="form-group" style={{ margin: 0 }}>
                <label style={{ fontSize: 11, fontWeight: 700, textTransform: 'uppercase', color: 'var(--text-muted)', marginBottom: 8 }}>
                  Edad Máxima para Niños (Años)
                </label>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10, padding: '10px 14px', background: 'var(--surface-2)', borderRadius: 8, border: '1px solid var(--border)' }}>
                  <button
                    type="button"
                    className="btn btn-ghost btn-sm"
                    style={{ width: 34, height: 34, fontWeight: 800, fontSize: 16, padding: 0, borderRadius: 6, border: '1px solid var(--border)' }}
                    onClick={() => setTempEdadMaxima(prev => Math.max(1, prev - 1))}
                  >-</button>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', minWidth: 80 }}>
                    <span style={{ fontSize: 19, fontWeight: 800, color: 'var(--primary)' }}>&lt; {tempEdadMaxima} A.</span>
                    <span style={{ fontSize: 10, color: 'var(--text-muted)' }}>menores de {tempEdadMaxima} años</span>
                  </div>
                  <button
                    type="button"
                    className="btn btn-ghost btn-sm"
                    style={{ width: 34, height: 34, fontWeight: 800, fontSize: 16, padding: 0, borderRadius: 6, border: '1px solid var(--border)' }}
                    onClick={() => setTempEdadMaxima(prev => Math.min(17, prev + 1))}
                  >+</button>
                </div>
              </div>
            </div>
            <div className="modal-footer" style={{ padding: '10px 18px', background: 'var(--surface-2)' }}>
              <button type="button" className="btn btn-ghost btn-sm" onClick={() => setShowAjustesModal(false)}>Cancelar</button>
              <button type="button" className="btn btn-primary btn-sm" onClick={() => {
                setEdadMaximaNinos(tempEdadMaxima);
                localStorage.setItem('residencia_edad_maxima_ninos', tempEdadMaxima.toString());
                setShowAjustesModal(false);
              }}>
                💾 Guardar Ajustes
              </button>
            </div>
          </div>
        </div>
      )}

      {/* POPUP DE SELECCIÓN DE HABITACIONES DISPONIBLES */}
      {showHabitacionesPopup && (
        <div className="modal-overlay" style={{ zIndex: 1100 }} onClick={e => e.target === e.currentTarget && setShowHabitacionesPopup(false)}>
          <div className="modal" style={{ maxWidth: 860, width: '92%' }}>
            <div className="modal-header" style={{ padding: '12px 18px' }}>
              <h3 style={{ fontSize: 16 }}>🛏️ Habitaciones Disponibles — Seleccionadas: {selectedHabitacionIds.length} / {totalHabitacionesSolicitadas || 1}</h3>
              <button className="modal-close" onClick={() => setShowHabitacionesPopup(false)}>×</button>
            </div>
            <div className="modal-body" style={{ maxHeight: '72vh', overflowY: 'auto', padding: '12px 18px' }}>
              <div style={{ fontSize: 11.5, marginBottom: 10, color: 'var(--text-muted)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>
                  Libres del <strong>{f.fechaEntrada}</strong> al <strong>{f.fechaSalida}</strong> (Solicitadas: <strong>{totalHabitacionesSolicitadas}</strong>):
                </span>
                <span className={`badge ${selectedHabitacionIds.length === totalHabitacionesSolicitadas ? 'badge-success' : 'badge-warning'}`} style={{ fontSize: 10.5, padding: '2px 8px' }}>
                  {selectedHabitacionIds.length} / {totalHabitacionesSolicitadas} Asignadas
                </span>
              </div>

              {loadingHabitacionesPopup ? (
                <div style={{ padding: 20, textAlign: 'center' }}>Cargando habitaciones libres...</div>
              ) : habitacionesDisponiblesPopup.length === 0 ? (
                <div className="alert alert-warning">
                  ⚠️ No hay ninguna habitación libre para las fechas y capacidad seleccionadas en esta residencia.
                </div>
              ) : (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: '8px' }}>
                  {habitacionesDisponiblesPopup.map(h => {
                    const isSelected = selectedHabitacionIds.includes(h.id);
                    const admiteCoAlojamiento = f.numNinos > 0 && (h.capacidadPersonas ?? 1) < f.numPersonas && (h.capacidadPersonas ?? 1) >= numAdultos;
                    return (
                      <div key={h.id}
                        style={{
                          padding: '6px 10px',
                          border: isSelected ? '2px solid var(--primary)' : '1px solid var(--border)',
                          borderRadius: 6,
                          background: isSelected ? 'rgba(45,138,78,0.12)' : 'var(--surface-2)',
                          cursor: 'pointer',
                          display: 'flex',
                          flexDirection: 'column',
                          justifyContent: 'center',
                          gap: 3,
                          transition: 'all 0.15s ease'
                        }}
                        onClick={() => toggleSeleccionHabitacionPopup(h.id)}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', lineHeight: 1.2 }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                            <span style={{
                              width: 16, height: 16, borderRadius: 4,
                              border: isSelected ? '2px solid var(--primary)' : '1.5px solid var(--text-muted)',
                              background: isSelected ? 'var(--primary)' : 'transparent',
                              color: '#fff', fontSize: 11, fontWeight: 900,
                              display: 'inline-flex', alignItems: 'center', justifyContent: 'center'
                            }}>{isSelected ? '✓' : ''}</span>
                            <strong style={{ fontSize: 13.5 }}>Hab. {h.numero}</strong>
                          </div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                            <span className="badge badge-primary" style={{ padding: '1px 6px', fontSize: 10 }}>{h.tipoCodigo}</span>
                            {isSelected && (
                              <span style={{ fontSize: 10, color: 'var(--primary)', fontWeight: 800 }}>
                                #{selectedHabitacionIds.indexOf(h.id) + 1}
                              </span>
                            )}
                          </div>
                        </div>

                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: 11, color: 'var(--text-muted)', marginTop: 1 }}>
                          <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '55%' }}>{h.tipoNombre}</span>
                          <span>Cama: <strong>{h.tipoCamaPrincipal}</strong> ({h.capacidadPersonas}p)</span>
                        </div>

                        {admiteCoAlojamiento && (
                          <div style={{ fontSize: 10, color: 'var(--primary)', fontWeight: 700, marginTop: 1, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                            👶 Co-alojamiento (niños en hab. {h.capacidadPersonas}p)
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div className="modal-footer" style={{ justifyContent: 'space-between' }}>
              <button type="button" className="btn btn-ghost" onClick={() => setShowHabitacionesPopup(false)}>Cancelar</button>
              <button type="button" className="btn btn-primary" onClick={confirmarSeleccionHabitacionesPopup}>
                ✓ Confirmar Selección ({selectedHabitacionIds.length} / {totalHabitacionesSolicitadas})
              </button>
            </div>
          </div>
        </div>
      )}

      {/* POPUP DE REEVALUACIÓN TRAS RENUNCIA */}
      {candidatosModal && candidatosModal.length > 0 && (
        <div className="modal-overlay" style={{ zIndex: 1250 }} onClick={e => e.target === e.currentTarget && setCandidatosModal(null)}>
          <div className="modal" style={{ maxWidth: 650, borderRadius: 12 }}>
            <div className="modal-header" style={{ background: 'rgba(230,162,60,0.15)', borderBottom: '1px solid rgba(230,162,60,0.3)' }}>
              <h3 style={{ fontSize: 15, margin: 0, color: 'var(--warning-dark, #b88200)', display: 'flex', alignItems: 'center', gap: 6 }}>
                ⚠️ Reevaluación por Renuncia de Reserva
              </h3>
              <button className="modal-close" onClick={() => setCandidatosModal(null)}>×</button>
            </div>
            <div className="modal-body" style={{ padding: '14px 18px', display: 'flex', flexDirection: 'column', gap: 10 }}>
              <p style={{ fontSize: 12.5, margin: 0 }}>
                La reserva ha sido marcada como <strong>RENUNCIA</strong>. Hay <strong>{candidatosModal.length}</strong> solicitud(es) en lista de espera para esas fechas y tipo de habitación.
              </p>
              <p style={{ fontSize: 12, color: 'var(--text-muted)', margin: 0 }}>
                Selecciona la solicitud a la que deseas reasignar automáticamente esta habitación liberada:
              </p>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8, maxHeight: 240, overflowY: 'auto' }}>
                {candidatosModal.map(cand => (
                  <div key={cand.id} style={{ padding: '10px 14px', border: '1px solid var(--border)', borderRadius: 8, background: 'var(--surface-2)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <strong style={{ fontSize: 13 }}>{cand.huespedNombreCompleto || cand.huespedNombre || 'Solicitante'}</strong>
                      <div style={{ fontSize: 11.5, color: 'var(--text-muted)' }}>
                        DNI: {cand.huespedDni} | Pax: {cand.numPersonas} | Entrada: {cand.fechaEntrada}
                      </div>
                    </div>
                    <button
                      type="button"
                      className="btn btn-primary btn-sm"
                      onClick={() => {
                        const candId = cand.id;
                        const candNom = cand.huespedNombreCompleto || cand.huespedNombre;
                        setCandidatosModal(null);
                        ejecutarGuardadoEdicion(candId, candNom);
                      }}>
                      Asignar y Reevaluar
                    </button>
                  </div>
                ))}
              </div>
            </div>
            <div className="modal-footer">
              <button type="button" className="btn btn-ghost btn-sm" onClick={() => setCandidatosModal(null)}>Cancelar</button>
              <button type="button" className="btn btn-secondary btn-sm" onClick={() => {
                setCandidatosModal(null);
                ejecutarGuardadoEdicion(null, null);
              }}>
                Guardar solo Renuncia sin Reasignar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}