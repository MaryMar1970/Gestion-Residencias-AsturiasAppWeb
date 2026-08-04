// ─── Constants & Selection Lists ────────────────────────────────────────────

export const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

export const DIAS_LONG = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

export const ESTADO_COLOR: Record<string, string> = {
  Confirmada: 'confirmada', Pendiente: 'pendiente',
  CheckIn: 'checkin', CheckOut: 'checkout', Bloqueada: 'bloqueada'
};

export const LISTA_EMPLEOS = [
  'Guardia', 'Cabo', 'Cabo 1º', 'Cabo Mayor', 'Sargento', 'Sargento 1º', 'Brigada',
  'Subteniente', 'Suboficial Mayor', 'Alférez', 'Teniente', 'Capitán', 'Comandante',
  'Tte. Coronel', 'Coronel', 'General B', 'General D', 'Funcionario', 'Estudiante', 'Investigador', 'Externo', 'Otro'
];

export const LISTA_SITUACIONES = [
  'Activo', 'Viogen', 'Asociación', 'Reserva activo', 'Reserva', 'Excedencia',
  'Especiales', 'Retirado', 'Viuda', 'Huerfano', 'Estudiante'
];

export const LISTA_FINALIDADES = [
  'Otros', 'Comisión', 'Comisión NO indem.', 'Destino', 'Enfermedad', 'Urgencia', 'Sepelio', 'Máx. Estancia'
];

export const LISTA_RESOLUCIONES = [
  'SI', 'CONCEDIDA', 'REEVALUADA', 'NO', 'DENEGADA', 'DESESTIMADA', 'RENUNCIA'
];
