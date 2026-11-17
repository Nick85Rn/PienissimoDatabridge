// Formato del file di importazione clienti di Pienissimo (vedi "esempio_importazioni.xlsx").
// Logica pura, senza React né SheetJS, così si può testare da sola.

export type Category = 'Anagrafica' | 'Contatti' | 'Indirizzo' | 'Dettagli' | 'Consensi';
export type CleanType = 'text' | 'phone' | 'email' | 'date' | 'flag' | 'cap' | 'province' | 'none';

export interface ImportColumn {
  id: string;
  /** Intestazione esatta richiesta dal file di importazione */
  label: string;
  category: Category;
  requiredGroup?: string;
  description?: string;
  cleanType: CleanType;
}

// Ordine e intestazioni identici al modello di importazione Pienissimo.
export const IMPORT_COLUMNS: ImportColumn[] = [
  { id: 'nome', label: 'nome', category: 'Anagrafica', description: 'Nome (o Nome Completo se attivi Split)', cleanType: 'text' },
  { id: 'cognome', label: 'cognome', category: 'Anagrafica', description: 'Cognome (lascia vuoto se usi Split)', cleanType: 'text' },
  { id: 'email', label: 'email', category: 'Contatti', requiredGroup: 'contatto', description: 'Email personale/aziendale', cleanType: 'email' },
  { id: 'telefono', label: 'telefono', category: 'Contatti', requiredGroup: 'contatto', description: 'Cellulare o fisso, con prefisso (es. +39…)', cleanType: 'phone' },
  { id: 'indirizzo', label: 'indirizzo', category: 'Indirizzo', description: 'Via e numero civico', cleanType: 'text' },
  { id: 'citta', label: 'citta', category: 'Indirizzo', description: 'Città di residenza', cleanType: 'text' },
  { id: 'cap', label: 'cap', category: 'Indirizzo', description: 'Codice postale (5 cifre per l\'Italia)', cleanType: 'cap' },
  { id: 'provincia', label: 'provincia', category: 'Indirizzo', description: 'Provincia (es. RN)', cleanType: 'province' },
  { id: 'nazione', label: 'nazione', category: 'Indirizzo', description: 'Nazione (es. Italia)', cleanType: 'text' },
  { id: 'note', label: 'note', category: 'Dettagli', description: 'Note libere aggiuntive', cleanType: 'none' },
  { id: 'data_nascita', label: 'data di nascita', category: 'Anagrafica', description: 'Convertita in AAAA-MM-GG (accetta GG/MM/AAAA, date Excel…)', cleanType: 'date' },
  { id: 'consenso_marketing', label: 'consenso_marketing', category: 'Consensi', description: 'Consenso marketing: 1 = sì, 0 = no (vuoto = 0)', cleanType: 'flag' },
  { id: 'profilazione', label: 'profilazione', category: 'Consensi', description: 'Consenso profilazione: 1 = sì, 0 = no (vuoto = 0)', cleanType: 'flag' },
  { id: 'tags', label: 'tags', category: 'Dettagli', description: 'Etichette separate da virgola', cleanType: 'none' },
];

export type Mapping = Record<string, string>;
export type Row = Record<string, unknown>;

export interface ProcessOptions {
  cleaningEnabled: boolean;
  splitNameEnabled: boolean;
  deduplicateEnabled: boolean;
  globalTag: string;
  defaultPrefix: string;
}

export interface ImportStats {
  totale: number;
  duplicatiRimossi: number;
  senzaContatto: number;
  dateNonValide: number;
  consensiNonRiconosciuti: number;
  emailNonValide: number;
}

// --- DATE ---------------------------------------------------------------------

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const pad = (n: number, l = 2) => String(n).padStart(l, '0');

function validYmd(y: number, m: number, d: number): string | null {
  if (y < 1900 || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const t = Date.UTC(y, m - 1, d);
  const dt = new Date(t);
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  if (t > Date.now()) return null; // data di nascita nel futuro
  return `${y}-${pad(m)}-${pad(d)}`;
}

/** Converte un valore di data in "AAAA-MM-GG". Restituisce null se non è una data riconoscibile. */
export function parseDate(value: unknown): string | null {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    return validYmd(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }

  if (typeof value === 'number') {
    // Numero seriale di Excel (giorni dal 30/12/1899)
    if (value < 1 || value > 2958465) return null;
    const dt = new Date(EXCEL_EPOCH + Math.floor(value) * 864e5);
    return validYmd(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
  }

  const s = String(value).trim();
  let m: RegExpMatchArray | null;

  // AAAA-MM-GG (anche con / o . e con orario dopo)
  if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:$|[T\s])/))) {
    return validYmd(+m[1], +m[2], +m[3]);
  }
  // GG/MM/AAAA o GG-MM-AA (formato italiano)
  if ((m = s.match(/^(\d{1,2})[-/.\s](\d{1,2})[-/.\s](\d{4}|\d{2})(?:$|[T\s])/))) {
    let y = +m[3];
    if (m[3].length === 2) {
      const cur = new Date().getFullYear() % 100;
      y += y > cur ? 1900 : 2000;
    }
    return validYmd(y, +m[2], +m[1]);
  }
  // AAAAMMGG
  if ((m = s.match(/^(\d{4})(\d{2})(\d{2})$/))) {
    return validYmd(+m[1], +m[2], +m[3]);
  }
  // Seriale Excel finito in un CSV come testo
  if (/^\d{4,5}(\.\d+)?$/.test(s)) return parseDate(Number(s));
  return null;
}

/** Numero seriale Excel per una data "AAAA-MM-GG" (per scrivere celle data nel file xlsx). */
export function dateToExcelSerial(ymd: string): number {
  const [y, m, d] = ymd.split('-').map(Number);
  return (Date.UTC(y, m - 1, d) - EXCEL_EPOCH) / 864e5;
}

// --- CONSENSI (1/0) -------------------------------------------------------------

const YES = new Set(['1', 'si', 'sì', 's', 'yes', 'y', 'true', 'vero', 'x', 'ok', 'accettato', 'consenso', 'acconsento']);
const NO = new Set(['0', 'no', 'n', 'false', 'falso', 'non accettato', 'rifiutato', 'negato']);

/** 1 / 0 per i consensi. Vuoto = 0 (nessun consenso esplicito). null = valore non riconosciuto. */
export function parseFlag(value: unknown): 0 | 1 | null {
  if (value === null || value === undefined) return 0;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'number') return value === 1 ? 1 : value === 0 ? 0 : null;
  const s = String(value).trim().toLowerCase();
  if (s === '') return 0;
  if (YES.has(s)) return 1;
  if (NO.has(s)) return 0;
  return null;
}

// --- PULIZIA CAMPI ------------------------------------------------------------

const LOWER_PARTICLES = new Set(['de', 'di', 'da', 'lo', 'la', 'le', 'del', 'della', 'delle', 'dei', 'degli', 'van', 'von']);

export function cleanText(str: string): string {
  return str.toLowerCase().split(/\s+/).map((w, i) => {
    if (i > 0 && LOWER_PARTICLES.has(w)) return w;
    return w.replace(/(?:^|['-])\S/g, (a) => a.toUpperCase());
  }).join(' ');
}

export function cleanPhone(str: string, defaultPrefix: string): string {
  let p = str.replace(/[^0-9+]/g, '');
  if (p.startsWith('00')) p = '+' + p.substring(2);
  if (p.startsWith('+')) return p;
  // Numero italiano già con 39 davanti ma senza "+" (es. 393331234567)
  if (defaultPrefix === '+39' && p.startsWith('39') && (p.length === 11 || p.length === 12)) return '+' + p;
  // Prefisso solo se la lunghezza è plausibile (9-10 cifre): interni e numeri parziali restano invariati.
  if (p.length >= 9 && p.length <= 10) return defaultPrefix + p;
  return p;
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
export const isValidEmail = (s: string) => EMAIL_RE.test(s);

const ITALY = /^(it|ita|italia|italy)$/i;
/** CAP italiani letti da Excel come numeri perdono gli zeri iniziali (00100 -> 100): li ripristiniamo. */
export function normalizeCap(cap: string, nazione: string): string {
  const c = cap.replace(/\s+/g, '');
  if (/^\d{3,4}$/.test(c) && (nazione.trim() === '' || ITALY.test(nazione.trim()))) return c.padStart(5, '0');
  return c;
}

const toStr = (v: unknown): string => {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString().slice(0, 10);
  return String(v).trim();
};

// --- MAPPATURA AUTOMATICA -------------------------------------------------------

const norm = (h: string) => h.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[_\-.]+/g, ' ').replace(/\s+/g, ' ').trim();

const FUZZY: Record<string, (h: string) => boolean> = {
  nome: h => /^(first ?name|nominativo|nome cliente)$/.test(h),
  cognome: h => /^(last ?name|surname)$/.test(h),
  email: h => /(^|\s)e ?mail|mail/.test(h),
  telefono: h => /(^|\s)(tel|cell|phone|mobile|whatsapp)/.test(h),
  indirizzo: h => /(^|\s)(via|indirizzo|address)(\s|$)/.test(h) && !/mail/.test(h),
  citta: h => /^(citta|comune|city|localita)/.test(h),
  cap: h => /(^|\s)cap(\s|$)|zip|postale|postal/.test(h),
  provincia: h => /^(prov|provincia|province)/.test(h),
  nazione: h => /^(nazione|paese|stato|country)/.test(h),
  note: h => /^(note|notes|commento|commenti)/.test(h),
  data_nascita: h => /nascita|birth|compleanno/.test(h),
  consenso_marketing: h => /marketing|newsletter|consenso privacy|consenso commerciale/.test(h),
  profilazione: h => /profilazion|profiling/.test(h),
  tags: h => /^(tag|tags|etichett|gruppo|segmento)/.test(h),
};

export function autoMatch(headers: string[]): Mapping {
  const mapping: Mapping = {};
  const used = new Set<string>();
  // 1° giro: corrispondenze esatte (intestazione o id). 2° giro: somiglianze, senza riusare colonne.
  for (const col of IMPORT_COLUMNS) {
    const exact = headers.find(h => !used.has(h) && (norm(h) === norm(col.label) || norm(h) === norm(col.id)));
    if (exact) { mapping[col.id] = exact; used.add(exact); }
  }
  for (const col of IMPORT_COLUMNS) {
    if (mapping[col.id]) continue;
    const test = FUZZY[col.id];
    const fuzzy = test && headers.find(h => !used.has(h) && test(norm(h)));
    if (fuzzy) { mapping[col.id] = fuzzy; used.add(fuzzy); }
  }
  return mapping;
}

// --- ELABORAZIONE ---------------------------------------------------------------

export type OutRow = Record<string, string | number>;

export function isBlankRow(row: Row): boolean {
  return Object.values(row).every(v => toStr(v) === '');
}

export function processRows(rows: Row[], mapping: Mapping, opt: ProcessOptions): { rows: OutRow[]; stats: ImportStats } {
  const stats: ImportStats = { totale: 0, duplicatiRimossi: 0, senzaContatto: 0, dateNonValide: 0, consensiNonRiconosciuti: 0, emailNonValide: 0 };

  const processed = rows.filter(r => !isBlankRow(r)).map(row => {
    const out: OutRow = {};

    let split: { nome: string; cognome: string } | null = null;
    if (opt.splitNameEnabled && mapping['nome'] && !mapping['cognome']) {
      const parts = toStr(row[mapping['nome']]).split(/\s+/);
      split = { nome: parts[0] || '', cognome: parts.slice(1).join(' ') };
    }

    for (const col of IMPORT_COLUMNS) {
      let raw: unknown = split && (col.id === 'nome' || col.id === 'cognome')
        ? split[col.id]
        : (mapping[col.id] ? row[mapping[col.id]] : '');

      if (col.id === 'tags' && opt.globalTag.trim()) {
        const tags = toStr(raw).split(',').map(t => t.trim()).filter(Boolean);
        if (!tags.includes(opt.globalTag.trim())) tags.push(opt.globalTag.trim());
        raw = tags.join(', ');
      }

      // Data e consensi vanno sempre convertiti: il file di importazione li richiede in quel formato.
      if (col.cleanType === 'date') {
        const d = parseDate(raw);
        if (d) out[col.label] = d;
        else { out[col.label] = toStr(raw); if (toStr(raw)) stats.dateNonValide++; }
        continue;
      }
      if (col.cleanType === 'flag') {
        const f = parseFlag(raw);
        if (f === null) { stats.consensiNonRiconosciuti++; out[col.label] = 0; }
        else out[col.label] = f;
        continue;
      }

      const s = toStr(raw);
      if (!s || !opt.cleaningEnabled) { out[col.label] = s; continue; }
      switch (col.cleanType) {
        case 'text': out[col.label] = cleanText(s); break;
        case 'email': out[col.label] = s.toLowerCase().replace(/^mailto:/, ''); break;
        case 'phone': out[col.label] = cleanPhone(s, opt.defaultPrefix); break;
        // Sigla di 2 lettere in maiuscolo (rm -> RM), nome esteso con iniziali maiuscole
        case 'province': out[col.label] = /^[a-z]{2}$/i.test(s) ? s.toUpperCase() : cleanText(s); break;
        default: out[col.label] = s;
      }
    }

    // Il CAP dipende dalla nazione, quindi si sistema a riga completa.
    if (opt.cleaningEnabled && out['cap']) out['cap'] = normalizeCap(String(out['cap']), String(out['nazione'] ?? ''));
    if (out['email'] && !isValidEmail(String(out['email']))) stats.emailNonValide++;
    return out;
  });

  stats.totale = processed.length;
  let result = processed;

  if (opt.deduplicateEnabled) {
    const unique = new Map<string, OutRow>();
    let noContact = 0;
    for (const row of processed) {
      const key = String(row['email'] || row['telefono'] || '');
      if (key.length > 5) {
        if (!unique.has(key)) unique.set(key, row);
        else stats.duplicatiRimossi++;
      } else {
        unique.set(`__no_contact_${noContact++}`, row);
        stats.senzaContatto++;
      }
    }
    result = Array.from(unique.values());
  }

  return { rows: result, stats };
}

/** Celle per il foglio Excel di importazione: testo come testo (mantiene + e zeri iniziali), date come date, consensi come numeri. */
export type XlsxCell = { t: 's'; v: string; z: '@' } | { t: 'n'; v: number; z?: string };

export function toXlsxCells(rows: OutRow[]): XlsxCell[][] {
  const header: XlsxCell[] = IMPORT_COLUMNS.map(c => ({ t: 's', v: c.label, z: '@' }));
  const body = rows.map(r => IMPORT_COLUMNS.map((c): XlsxCell => {
    const v = r[c.label];
    if (c.cleanType === 'flag') return { t: 'n', v: Number(v) || 0 };
    if (c.cleanType === 'date' && typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
      return { t: 'n', v: dateToExcelSerial(v), z: 'yyyy-mm-dd' };
    }
    return { t: 's', v: v === undefined ? '' : String(v), z: '@' };
  }));
  return [header, ...body];
}
