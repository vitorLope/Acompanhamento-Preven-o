#!/usr/bin/env node

// Generates two sample Excel files (CONSINCO and INFOR) in ./samples for manual testing
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function generateConsinco(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  // headers using codes indicated in the specification
  ws.addRow(['T', 'U', 'Z', 'FA', 'FB', 'FD', 'EX', 'CD', 'Colaborador', 'Produto']);
  ws.addRow(['1º Turno', 'CARGA-001', 'Separação', 10, 10, 'Falta', new Date(), 'CD001', 'João', 'Produto A']);
  ws.addRow(['2º Turno', 'CARGA-002', 'Separação', 8, 8, 'Sobra', new Date(), 'CD001', 'Maria', 'Produto B']);
  await wb.xlsx.writeFile(filePath);
}

async function generateInfor(filePath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.addRow(['T', 'Z', 'MJ', 'MM', 'MN', 'MO', 'MR', 'CD', 'Colaborador', 'Produto']);
  ws.addRow(['1º Turno', 'Separação', new Date(), 'CARGA-010', 12, 12, 'Falta', 'CD002', 'Carlos', 'Produto X']);
  ws.addRow(['3º Turno', 'Separação', new Date(), 'CARGA-011', 5, 5, 'Sobra', 'CD002', 'Ana', 'Produto Y']);
  await wb.xlsx.writeFile(filePath);
}

(async () => {
  const samplesDir = path.resolve(process.cwd(), 'samples');
  if (!fs.existsSync(samplesDir)) fs.mkdirSync(samplesDir);

  const cPath = path.join(samplesDir, 'sample-consinco.xlsx');
  const iPath = path.join(samplesDir, 'sample-infor.xlsx');

  await generateConsinco(cPath);
  await generateInfor(iPath);

  console.log('Samples generated:', cPath, iPath);
})();
