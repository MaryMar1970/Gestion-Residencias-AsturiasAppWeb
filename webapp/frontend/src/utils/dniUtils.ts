// ─── ALGORITMO DE VALIDACIÓN Y CORRECCIÓN DE DNI ──────────────────────────────
const LETRAS_DNI = 'TRWAGMYFPDXBNJZSQVHLCKE';

export function validarYCorregirDNI(dniInput: string): { esValido: boolean; dniFinal: string } {
  // 1. Comprueba si el DNI está vacío. Si está vacío, no realiza ninguna validación y lo considera válido.
  if (!dniInput) {
    return { esValido: true, dniFinal: '' };
  }

  // 2. Normaliza el texto: Elimina espacios al principio y al final, convierte la letra a mayúsculas.
  const normalizado = dniInput.trim().toUpperCase();
  if (normalizado === '') {
    return { esValido: true, dniFinal: '' };
  }

  // 2.1 Caso especial: 8 dígitos exactos sin letra -> calcular y ofrecer/añadir la letra correspondiente
  if (/^\d{8}$/.test(normalizado)) {
    const numero = parseInt(normalizado, 10);
    const resto = numero % 23;
    const letraCalculada = LETRAS_DNI[resto];
    const dniCompleto = `${normalizado}${letraCalculada}`;

    const aceptaCompletar = window.confirm(
      `El DNI introducido (${normalizado}) contiene 8 números sin la letra de control.\n\n` +
      `La letra calculada que le corresponde al número ${normalizado} es "${letraCalculada}".\n` +
      `El DNI completo es: ${dniCompleto}.\n\n` +
      `¿Deseas añadir automáticamente la letra correspondiente (${dniCompleto})?`
    );

    if (aceptaCompletar) {
      return { esValido: true, dniFinal: dniCompleto };
    } else {
      return { esValido: false, dniFinal: normalizado };
    }
  }

  // 3. Verifica el formato: Comprueba que tenga exactamente 8 dígitos seguidos de una letra.
  const regexDNI = /^\d{8}[A-Z]$/;
  if (!regexDNI.test(normalizado)) {
    alert(`Aviso de formato: El DNI "${normalizado}" no tiene un formato válido.\n\nDebe constar de exactamente 8 números seguidos de una letra (ejemplo: 12345678Z).`);
    return { esValido: false, dniFinal: normalizado };
  }

  // 4. Calcula la letra correcta
  const numero = parseInt(normalizado.substring(0, 8), 10);
  const letraIntroducida = normalizado.substring(8, 9);
  const resto = numero % 23;
  const letraCalculada = LETRAS_DNI[resto];

  // 5. Compara la letra introducida con la calculada
  if (letraIntroducida === letraCalculada) {
    return { esValido: true, dniFinal: normalizado };
  }

  // 6. Si no coinciden, construye el DNI correcto y ofrece corregir
  const dniCorrecto = `${normalizado.substring(0, 8)}${letraCalculada}`;
  const aceptaCorreccion = window.confirm(
    `El DNI introducido (${normalizado}) tiene una letra incorrecta.\n\n` +
    `La letra calculada correspondiente al número ${normalizado.substring(0, 8)} es "${letraCalculada}".\n` +
    `El DNI correcto debería ser: ${dniCorrecto}.\n\n` +
    `¿Deseas sustituir automáticamente la letra por la correcta (${dniCorrecto})?`
  );

  if (aceptaCorreccion) {
    return { esValido: true, dniFinal: dniCorrecto };
  } else {
    return { esValido: false, dniFinal: normalizado };
  }
}
