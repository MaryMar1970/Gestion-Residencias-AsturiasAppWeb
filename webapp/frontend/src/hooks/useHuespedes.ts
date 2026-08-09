import { useState, useCallback } from 'react';
import { apiFetch } from '../api';
import { Huesped } from '../types';
import { validarYCorregirDNI } from '../utils/dniUtils';

export function useHuespedes() {
  const [buscarHuespedTexto, setBuscarHuespedTexto] = useState('');
  const [huespedesSugeridos, setHuespedesSugeridos] = useState<Huesped[]>([]);
  const [buscandoHuespedes, setBuscandoHuespedes] = useState(false);
  const [huesped, setHuesped] = useState<Huesped | null>(null);
  const [huespedNuevo, setHuespedNuevo] = useState(false);

  const [huespedForm, setHuespedForm] = useState({
    dni: '',
    nombre: '',
    apellidos: '',
    telefono: '',
    email: '',
    direccion: '',
    codigoPostal: '',
    municipio: '',
    provincia: '',
    empleo: 'Guardia',
    situacion: 'Activo',
    tipoHuesped: 'Externo',
    centroOrigen: '',
    departamento: '',
    enListaNegra: false,
    motivoListaNegra: '',
    notas: '',
    familiaNumerosa: 'NO',
    porcentajeDescuento: 0,
  });

  const buscarHuespedesAPI = useCallback(async (texto: string) => {
    if (!texto.trim()) {
      setHuespedesSugeridos([]);
      return;
    }
    setBuscandoHuespedes(true);
    try {
      const q = encodeURIComponent(texto.trim());
      const res = await apiFetch<any>(`/api/huespedes?buscar=${q}`);
      setHuespedesSugeridos(Array.isArray(res) ? res : (res?.items ?? []));
    } catch {
      setHuespedesSugeridos([]);
    } finally {
      setBuscandoHuespedes(false);
    }
  }, []);

  const seleccionarHuesped = useCallback((h: Huesped) => {
    setHuesped(h);
    setHuespedNuevo(false);
    setHuespedesSugeridos([]);
    setBuscarHuespedTexto(`${h.dni} - ${h.nombre} ${h.apellidos || ''}`);
  }, []);

  const abrirRegistroNuevoHuesped = useCallback((dniInicial = '') => {
    setHuesped(null);
    setHuespedNuevo(true);
    setHuespedesSugeridos([]);

    const { dniFinal: dniLimpio } = validarYCorregirDNI(dniInicial.trim().toUpperCase());
    setHuespedForm({
      dni: dniLimpio,
      nombre: '',
      apellidos: '',
      telefono: '',
      email: '',
      direccion: '',
      codigoPostal: '',
      municipio: '',
      provincia: '',
      empleo: 'Guardia',
      situacion: 'Activo',
      tipoHuesped: 'Externo',
      centroOrigen: '',
      departamento: '',
      enListaNegra: false,
      motivoListaNegra: '',
      notas: '',
      familiaNumerosa: 'NO',
      porcentajeDescuento: 0,
    });
  }, []);

  const handleBuscarKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      void buscarHuespedesAPI(buscarHuespedTexto);
    }
  }, [buscarHuespedTexto, buscarHuespedesAPI]);

  const handleBlurDniHuespedForm = useCallback(async () => {
    const dniIngresado = huespedForm.dni.trim().toUpperCase();
    if (!dniIngresado) return;

    const { dniFinal: dniCorregido } = validarYCorregirDNI(dniIngresado);
    if (dniCorregido !== huespedForm.dni) {
      setHuespedForm(prev => ({ ...prev, dni: dniCorregido }));
    }

    try {
      const q = encodeURIComponent(dniCorregido);
      const existentes = await apiFetch<Huesped[]>(`/api/huespedes?buscar=${q}`);
      const coincidenciaExacta = existentes.find(h => h.dni.toUpperCase() === dniCorregido);

      if (coincidenciaExacta) {
        setHuesped(coincidenciaExacta);
        setHuespedNuevo(false);
        setBuscarHuespedTexto(`${coincidenciaExacta.dni} - ${coincidenciaExacta.nombre} ${coincidenciaExacta.apellidos || ''}`);
      }
    } catch {}
  }, [huespedForm.dni]);

  return {
    buscarHuespedTexto, setBuscarHuespedTexto,
    huespedesSugeridos, setHuespedesSugeridos,
    buscandoHuespedes, setBuscandoHuespedes,
    huesped, setHuesped,
    huespedNuevo, setHuespedNuevo,
    huespedForm, setHuespedForm,
    buscarHuespedesAPI,
    seleccionarHuesped,
    abrirRegistroNuevoHuesped,
    handleBuscarKeyDown,
    handleBlurDniHuespedForm,
  };
}
