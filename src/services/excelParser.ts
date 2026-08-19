import ExcelJS from 'exceljs';

export type Layout = 'CONSINCO' | 'INFOR' | 'UNKNOWN';

// explicit expected column keys for each layout (header codes or common header texts)
const CONSINCO_KEYS = {
  T: ['t', 'turno'],
  U: ['u', 'carga', 'numero carga', 'nº carga'],
  Z: ['z', 'tipo atividade', 'atividade', 'tipo de atividade'],
  FA: ['fa', 'volume', 'volume auditado', 'volume_auditado'],
  FB: ['fb', 'itens', 'itens verificadas', 'itens_verificados'],
  FD: ['fd', 'diverg', 'divergências', 'divergencias', 'diver'],
  EX: ['ex', 'data', 'data_inspecao', 'data inspeção']
};

const INFOR_KEYS = {
  T: ['t', 'turno'],
  Z: ['z', 'tipo atividade', 'atividade'],
  MJ: ['mj', 'data', 'data_inspecao', 'data inspeção'],
  MM: ['mm', 'carga', 'numero carga', 'nº carga'],
  MN: ['mn', 'volume', 'volume auditado'],
  MO: ['mo', 'itens', 'itens_verificados'],
  MR: ['mr', 'diverg', 'divergências', 'divergencias']
};

function normalizeHeader(value: any) {
  if (value === null || value === undefined) return '';
  return String(value).toLowerCase().trim();
}

function findColumnIndexByCandidates(headers: string[], candidates: string[]) {
  // try exact equal match first
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    for (const c of candidates) {
      if (h === c) return i + 1; // ExcelJS columns are 1-based
    }
  }
  // try contains
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    for (const c of candidates) {
      if (c && h.includes(c)) return i + 1;
    }
  }
  return -1;
}

export function detectLayout(sheet: ExcelJS.Worksheet): Layout {
  // read header row (first non-empty row)
  let headerRowIndex = 1;
  for (let r = 1; r <= Math.min(5, sheet.rowCount); r++) {
    const row = sheet.getRow(r);
    const isEmpty = row.values.every(v => v === null || v === undefined || v === '');
    if (!isEmpty) { headerRowIndex = r; break; }
  }

  const headerRow = sheet.getRow(headerRowIndex);
  const headers = [] as string[];
  for (let c = 1; c <= headerRow.cellCount; c++) {
    headers.push(normalizeHeader(headerRow.getCell(c).value));
  }

  // check for CONSINCO keys
  const consMatches = Object.values(CONSINCO_KEYS).reduce((acc, cand) => {
    const idx = findColumnIndexByCandidates(headers, cand);
    return acc + (idx > 0 ? 1 : 0);
  }, 0);

  const inforMatches = Object.values(INFOR_KEYS).reduce((acc, cand) => {
    const idx = findColumnIndexByCandidates(headers, cand);
    return acc + (idx > 0 ? 1 : 0);
  }, 0);

  if (consMatches >= 3 && consMatches > inforMatches) return 'CONSINCO';
  if (inforMatches >= 3 && inforMatches > consMatches) return 'INFOR';

  // fallback: look for specific code-like headers (fd, mr, fa, mj etc)
  const joined = headers.join(' ');
  if (joined.includes('fd') || joined.includes('fa') || joined.includes('fb')) return 'CONSINCO';
  if (joined.includes('mr') || joined.includes('mj') || joined.includes('mm')) return 'INFOR';

  return 'UNKNOWN';
}

export function parseRowsByLayout(sheet: ExcelJS.Worksheet, layout: Layout) {
  // determine header row
  let headerRowIndex = 1;
  for (let r = 1; r <= Math.min(5, sheet.rowCount); r++) {
    const row = sheet.getRow(r);
    const isEmpty = row.values.every(v => v === null || v === undefined || v === '');
    if (!isEmpty) { headerRowIndex = r; break; }
  }

  const headerRow = sheet.getRow(headerRowIndex);
  const headers: string[] = [];
  for (let c = 1; c <= headerRow.cellCount; c++) {
    headers.push(normalizeHeader(headerRow.getCell(c).value));
  }

  // build mapping depending on detected layout
  const mapping: Record<string, number> = {};
  const keys = layout === 'CONSINCO' ? CONSINCO_KEYS : INFOR_KEYS;

  for (const [key, candidates] of Object.entries(keys)) {
    const idx = findColumnIndexByCandidates(headers, candidates as string[]);
    if (idx > 0) mapping[key] = idx;
  }

  const rows: any[] = [];
  for (let r = headerRowIndex + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    // skip empty rows
    const isEmpty = row.values.every(v => v === null || v === undefined || v === '');
    if (isEmpty) continue;

    // extract fields according to mapping; if mapping missing a column, fallback to best-effort
    const get = (keyName: string) => {
      const idx = mapping[keyName];
      if (!idx) return null;
      const val = row.getCell(idx).value;
      return val === null || val === undefined ? null : val;
    };

    // Normalized object shape expected by worker
    const data_inspecao_raw = get(layout === 'CONSINCO' ? 'EX' : 'MJ');
    const data_inspecao = parseExcelDate(data_inspecao_raw);

    const numero_carga = String(get(layout === 'CONSINCO' ? 'U' : 'MM') ?? '').trim();
    const turno = String(get('T') ?? '').trim();
    const tipo_atividade = String(get('Z') ?? '').trim();
    const volume = get(layout === 'CONSINCO' ? 'FA' : 'MN');
    const itens = get(layout === 'CONSINCO' ? 'FB' : 'MO');
    const divergencia_raw = get(layout === 'CONSINCO' ? 'FD' : 'MR');

    rows.push({
      data_inspecao,
      cd: extractCdFromRow(row, headers),
      turno,
      tipo_atividade,
      numero_carga,
      volume,
      itens,
      divergencia_raw,
      colaborador: extractCellByHeader(row, headers, ['colaborador', 'operador', 'funcionário', 'colab']),
      produto: extractCellByHeader(row, headers, ['produto', 'produto descricao', 'descricao produto', 'sku']),
      quantidade: numberOrZero(get(layout === 'CONSINCO' ? 'FA' : 'MN'))
    });
  }

  return rows;
}

function extractCellByHeader(row: ExcelJS.Row, headers: string[], candidates: string[]) {
  const idx = findColumnIndexByCandidates(headers, candidates);
  if (idx > 0) return row.getCell(idx).value ?? null;
  return null;
}

function extractCdFromRow(row: ExcelJS.Row, headers: string[]) {
  const idx = findColumnIndexByCandidates(headers, ['cd', 'centro', 'cd destino', 'cd_origem']);
  if (idx > 0) return String(row.getCell(idx).value ?? '').trim();
  return '';
}

function numberOrZero(v: any) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[,.]/g, m => (m === ',' ? '.' : '')));
  return isNaN(n) ? 0 : n;
}

function parseExcelDate(v: any): Date {
  if (!v) return new Date();
  if (v instanceof Date) return v;
  // ExcelJS serial date possibility
  if (typeof v === 'number') {
    // Excel serial date (days since 1899-12-31), using ExcelJS util would be better but quick approximation:
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const days = Math.floor(v);
    const ms = days * 24 * 60 * 60 * 1000;
    return new Date(epoch.getTime() + ms);
  }
  const parsed = new Date(String(v));
  if (!isNaN(parsed.getTime())) return parsed;
  return new Date();
}

export function explodeDivergencias(row: any) {
  const raw = (row.divergencia_raw ?? '') as any;
  let text = '';
  if (raw === null || raw === undefined) text = '';
  else if (typeof raw === 'string') text = raw;
  else if (typeof raw === 'number') text = String(raw);
  else if (raw?.richText) text = raw.richText.map((t: any) => t.text).join('');
  else text = String(raw);

  if (!text) return [];

  const parts = text.split(/[,;\n]+/).map((p: string) => p.trim()).filter(Boolean);
  const result: any[] = [];

  for (const p of parts) {
    const tipo = mapTipo(p);
    result.push({
      data_inspecao: row.data_inspecao,
      cd: row.cd,
      turno: row.turno,
      tipo_atividade: row.tipo_atividade,
      numero_carga: row.numero_carga,
      tipo_divergencia: tipo,
      colaborador: row.colaborador ?? null,
      produto: row.produto ?? null,
      quantidade: row.quantidade ?? 1
    });
  }

  return result;
}

function mapTipo(text: string) {
  const t = String(text).toLowerCase();
  if (t.includes('falta')) return 'FALTA';
  if (t.includes('sobra')) return 'SOBRA';
  if (t.includes('invers')) return 'INVERSÃO';
  if (t.includes('avaria')) return 'AVARIA';
  // also try numeric negative or positive hints
  return 'FALTA';
}
