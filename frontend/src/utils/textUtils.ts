// ─── Text Utilities & Normalization ──────────────────────────────────────────
import diccionarioCompleto from './diccionarioEspañaCompleto.json';

export function removeAccents(str?: string | null): string {
  if (!str) return '';
  return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

export function normalizeKey(str?: string | null): string {
  if (!str) return '';
  return removeAccents(str).toLowerCase().replace(/\s+/g, ' ').trim();
}

// Partículas y conectores de nombres/apellidos/direcciones que deben permanecer en minúscula (salvo al inicio)
const MINUSCULAS_CONECTORES = new Set([
  'de', 'del', 'la', 'las', 'el', 'los', 'y', 'e', 'o', 'u', "d'", "l'"
]);

// Diccionario ortográfico oficial INE (Nombres, Apellidos, Municipios y Provincias de España)
const DICCIONARIO_ORTOGRAFIA: Record<string, string> = {
  ...(diccionarioCompleto as Record<string, string>),
  // Garantizar correcciones primarias con tildes iniciales
  maria: 'María',
  jose: 'José',
  jesus: 'Jesús',
  angel: 'Ángel',
  oscar: 'Óscar',
  adrian: 'Adrián',
  ruben: 'Rubén',
  ramon: 'Ramón',
  raul: 'Raúl',
  monica: 'Mónica',
  veronica: 'Verónica',
  alvaro: 'Álvaro',
  lucia: 'Lucía',
  sofia: 'Sofía',
  andres: 'Andrés',
  tomas: 'Tomás',
  victor: 'Víctor',
  hector: 'Héctor',
  joaquin: 'Joaquín',
  agustin: 'Agustín',
  fermin: 'Fermín',
  german: 'Germán',
  ivan: 'Iván',
  julian: 'Julián',
  nicolas: 'Nicolás',
  sebastian: 'Sebastián',
  inigo: 'Íñigo'
};

export function toTitleCase(str?: string | null): string {
  if (!str) return '';
  const words = str.trim().split(/\s+/);
  return words
    .map((word, index) => {
      if (!word) return '';
      const cleanWord = removeAccents(word).toLowerCase();

      // Si es una partícula/conector y NO es la primera palabra del campo, se mantiene en minúscula
      if (index > 0 && MINUSCULAS_CONECTORES.has(cleanWord)) {
        return cleanWord;
      }

      if (DICCIONARIO_ORTOGRAFIA[cleanWord]) {
        return DICCIONARIO_ORTOGRAFIA[cleanWord];
      }

      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

export function formatEmail(email?: string | null): string {
  if (!email) return '';
  return email.trim().toLowerCase();
}
