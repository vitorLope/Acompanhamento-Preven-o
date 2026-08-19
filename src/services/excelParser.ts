import ExcelJS from 'exceljs';

export type Layout = 'CONSINCO' | 'INFOR' | 'UNKNOWN';

export function detectLayout(sheet: ExcelJS.Worksheet): Layout {
  // Read first non-empty row(s) and check for header keywords
  const headerRows: string[] = [];
  for (let r = 1; r <= Math.min(5, sheet.rowCount); r++) {
    const row = sheet.getRow(r);
    const texts = row.values?.slice(1).map(v => (v || '').toString().toLowerCase());
    if (texts && texts.length) headerRows.push(...texts);
  }

  const joined = headerRows.join(' ');
  if (joined.includes('diverg') || joined.includes('difer')) {
    // heurística simples
    if (joined.includes('fd') || joined.includes('fa') || joined.includes('fb')) return 'CONSINCO';
    if (joined.includes('mr') || joined.includes('mj') || joined.includes('mm')) return 'INFOR';
    // fallback based on keywords
    if (joined.includes('numero') && joined.includes('carga')) return 'CONSINCO';
  }

  // sample data-based heuristic
  for (let r = 6; r <= Math.min(20, sheet.rowCount); r++) {
    const row = sheet.getRow(r);
    const values = row.values?.slice(1).map(v => (v || '').toString().toLowerCase());
    if (!values) continue;
    const joinedRow = values.join(' ');
    if (joinedRow.includes('consinco')) return 'CONSINCO';
    if (joinedRow.includes('infor')) return 'INFOR';
  }

  return 'UNKNOWN';
}

export function parseRowsByLayout(sheet: ExcelJS.Worksheet, layout: Layout) {
  // Very lightweight parser: assumes first row is header, maps common column titles
  const headerRow = sheet.getRow(1);
  const headers: string[] = [];
  for (let c = 1; c <= headerRow.cellCount; c++) {
    const v = headerRow.getCell(c).value;
    headers.push((v || '').toString().trim().toLowerCase());
  }

  const rows: any[] = [];
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    if (row.values.every(v => v === null || v === undefined || v === '')) continue;
    const obj: any = {};
    for (let c = 1; c <= headers.length; c++) {
      const key = headers[c - 1] || `col_${c}`;
      obj[key] = row.getCell(c).value;
    }

    // Normalize common fields into a predictable shape
    rows.push({
      raw: obj,
      // try to map minimal fields used by worker
      data_inspecao: extractDate(obj),
      cd: String(obj['cd'] || obj['cd destino'] || obj['cd destino'] || obj['centro'] || ''),
      turno: String(obj['turno'] || obj['t'] || ''),
      tipo_atividade: String(obj['tipo atividade'] || obj['atividade'] || ''),
      numero_carga: String(obj['carga'] || obj['numero carga'] || obj['nº carga'] || obj['n'] || ''),
      divergencia_raw: String(obj['divergências'] || obj['diver'] || obj['div'] || obj['d'] || ''),
      colaborador: String(obj['colaborador'] || obj['nome'] || ''),
      produto: String(obj['produto'] || obj['produto descricao'] || ''),
      quantidade: Number(obj['volume'] || obj['volume auditado'] || obj['quantidade'] || 0) || 0
    });
  }

  return rows;
}

function extractDate(obj: any): Date {
  const v = obj['data'] || obj['data_inspecao'] || obj['data inspeção'] || obj['data inspecao'];
  if (!v) return new Date();
  if (v instanceof Date) return v;
  const parsed = new Date(String(v));
  if (!isNaN(parsed.getTime())) return parsed;
  return new Date();
}

export function explodeDivergencias(row: any) {
  // row.divergencia_raw might contain multiple items separated by comma/; or be empty
  const raw = (row.divergencia_raw || '').toString();
  if (!raw) return [];

  const parts = raw.split(/[,;\n]+/).map(p => p.trim()).filter(Boolean);
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
      colaborador: row.colaborador,
      produto: row.produto,
      quantidade: row.quantidade || 1
    });
  }
  return result;
}

function mapTipo(text: string) {
  const t = text.toLowerCase();
  if (t.includes('falta')) return 'FALTA';
  if (t.includes('sobra')) return 'SOBRA';
  if (t.includes('invers')) return 'INVERSÃO';
  if (t.includes('avaria')) return 'AVARIA';
  // fallback: if contains '-' or negative words assume FALTA
  return 'FALTA';
}
