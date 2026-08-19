import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { detectLayout, parseRowsByLayout, explodeDivergencias } from '../src/services/excelParser';

describe('excelParser (CONSINCO layout)', () => {
  it('detects CONSINCO layout and parses rows', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['T', 'U', 'Z', 'FA', 'FB', 'FD', 'EX', 'CD', 'Colaborador', 'Produto']);
    ws.addRow(['1º Turno', 'CARGA-001', 'Separação', 10, 10, 'Falta', '2026-08-18', 'CD001', 'João', 'Produto A']);
    ws.addRow(['2º Turno', 'CARGA-002', 'Separação', 8, 8, 'Sobra', '2026-08-18', 'CD001', 'Maria', 'Produto B']);

    const layout = detectLayout(ws);
    expect(layout).toBe('CONSINCO');

    const rows = parseRowsByLayout(ws, layout);
    expect(rows.length).toBe(2);

    const occ0 = explodeDivergencias(rows[0]);
    expect(occ0.length).toBeGreaterThanOrEqual(1);
    expect(occ0[0].tipo_divergencia).toBe('FALTA');
  });
});

describe('excelParser (INFOR layout)', () => {
  it('detects INFOR layout and parses rows', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.addRow(['T', 'Z', 'MJ', 'MM', 'MN', 'MO', 'MR', 'CD', 'Colaborador', 'Produto']);
    ws.addRow(['1º Turno', 'Separação', '2026-08-18', 'CARGA-010', 12, 12, 'Falta', 'CD002', 'Carlos', 'Produto X']);
    ws.addRow(['3º Turno', 'Separação', '2026-08-18', 'CARGA-011', 5, 5, 'Sobra', 'CD002', 'Ana', 'Produto Y']);

    const layout = detectLayout(ws);
    expect(layout).toBe('INFOR');

    const rows = parseRowsByLayout(ws, layout);
    expect(rows.length).toBe(2);

    const occ0 = explodeDivergencias(rows[0]);
    expect(occ0.length).toBeGreaterThanOrEqual(1);
    expect(occ0[0].tipo_divergencia).toBe('FALTA');
  });
});
