// ─── Domain Types — Consolidated from Reservas.tsx & App.tsx ─────────────────

export type User = { id: string; email: string; fullName: string; roles: string[] };

export type PaginatedResult<T> = {
  items: T[];
  totalCount: number;
  page: number;
  pageSize: number;
  totalPages: number;
};

export const RESOLUCIONES = {
  SI: 'SI',
  NO: 'NO',
  CONCEDIDA: 'CONCEDIDA',
  DENEGADA: 'DENEGADA',
  DESESTIMADA: 'DESESTIMADA',
  RENUNCIA: 'RENUNCIA',
  REEVALUADA: 'REEVALUADA'
} as const;

export type TipoResolucion = typeof RESOLUCIONES[keyof typeof RESOLUCIONES];

export type Residencia = {
  id: string;
  nombre: string;
  razonSocial?: string;
  cif?: string;
  direccion?: string;
  codigoPostal?: string;
  municipio?: string;
  provincia?: string;
  telefono?: string;
  email?: string;
  serieFactura?: string;
  activa: boolean;
  orden: number;
  totalHabitaciones: number;
};

export type Habitacion = {
  id: string;
  residenciaId: string;
  residenciaNombre: string;
  tipoHabitacionId: string;
  tipoNombre: string;
  tipoCodigo: string;
  numero: string;
  nombre?: string;
  tipoCamaPrincipal?: string;
  capacidadPersonas?: number;
  admiteSupletorias?: boolean;
  plazasSupletorias?: number;
  activa?: boolean;
  notas?: string;
  orden?: number;
};

export type Huesped = {
  id: string;
  dni: string;
  nombre: string;
  apellidos: string;
  nombreCompleto: string;
  telefono?: string;
  email?: string;
  tipoHuesped: string;
  enListaNegra: boolean;
  totalReservas: number;
  direccion?: string;
  codigoPostal?: string;
  municipio?: string;
  provincia?: string;
  centroOrigen?: string;
  departamento?: string;
  empleo?: string;
  situacion?: string;
  finalidad?: string;
  empleoCategoria?: string;
  motivoListaNegra?: string;
  notas?: string;
  familiaNumerosa?: string;
  porcentajeDescuento?: number;
};

export type Reserva = {
  id: string;
  numeroOrden: number;
  habitacionId?: string | null;
  habitacionNumero?: string | null;
  residenciaNombre?: string | null;
  tipoHabitacionNombre?: string | null;
  huespedId?: string;
  huespedNombreCompleto?: string;
  huespedDni?: string;
  huespedNombre?: string;
  huespedApellidos?: string;
  huespedTelefono?: string;
  huespedEmail?: string;
  huespedSituacion?: string;
  huespedRango?: string;
  huespedEmpleoCategoria?: string;
  fechaEntrada: string;
  fechaSalida: string;
  totalNoches: number;
  numPersonas: number;
  numNinos?: number;
  alojamientoSolicitado?: string;
  estado: string;
  estadoInt: number;
  esBloqueo: boolean;
  motivoBloqueo?: string;
  tarifaNombreSnapshot?: string;
  precioNocheAplicado: number;
  importeTotal: number;
  pagado: boolean;
  formaPago?: string;
  observaciones?: string;
  finalidad?: string;
  empleo?: string;
  evaluacion?: string;
};

export type CalHabitacion = {
  habitacionId: string;
  residenciaId?: string;
  numero: string;
  residenciaNombre: string;
  tipoNombre: string;
  tipoCodigo: string;
  activa: boolean;
  orden: number;
  reservas: Reserva[];
};

export type CalData = {
  fechaInicio: string;
  fechaFin: string;
  habitaciones: CalHabitacion[];
};

export type Festivo = {
  id: string;
  fecha: string;
  descripcion: string;
  ambito?: string;
  activo?: boolean;
};

export type TipoHab = {
  id: string;
  nombre: string;
  codigo: string;
  capacidadMaxima: number;
  admiteSupletorias: boolean;
  descripcion?: string;
  activo: boolean;
  orden: number;
};

export type Tarifa = {
  id: string;
  residenciaId: string;
  residenciaNombre: string;
  tipoHabitacionId: string;
  tipoHabitacionNombre: string;
  nombreTarifa: string;
  precioNoche: number;
  precioMes: number;
  porcentajeIva: number;
  descripcion?: string;
  activa: boolean;
};

export type Page = 'dashboard' | 'residencias' | 'alojamientos' | 'tarifas' | 'festivos' | 'evaluacion' | 'calendario' | 'reservas' | 'huespedes' | 'facturas' | 'bloqueos';
