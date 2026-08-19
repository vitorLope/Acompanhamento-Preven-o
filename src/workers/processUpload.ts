import { Worker, Job } from 'bullmq';
import ExcelJS from 'exceljs';
import prisma from '../prisma';
import { detectLayout, parseRowsByLayout, explodeDivergencias } from '../services/excelParser';
import { redisClient } from '../lib/queue';

const connection = redisClient; // reuse ioredis connection

const BATCH_SIZE = 1000;

const worker = new Worker('default', async (job: Job) => {
  const { uploadId, filepath } = job.data as { uploadId: number; filepath: string };

  await prisma.tbUpload.update({ where: { id: uploadId }, data: { status: 'PROCESSANDO' } });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filepath);
    const sheet = workbook.worksheets[0];

    const layout = detectLayout(sheet);
    const parsed = parseRowsByLayout(sheet, layout);

    let totalOcorrencias = 0;
    const cargasSet = new Set<string>();
    const occurrences: any[] = [];

    for (const row of parsed) {
      if (row.numero_carga) cargasSet.add(String(row.numero_carga));
      const divergencias = explodeDivergencias(row);
      for (const d of divergencias) {
        occurrences.push({
          data_inspecao: d.data_inspecao,
          cd: d.cd,
          turno: d.turno,
          tipo_atividade: d.tipo_atividade,
          numero_carga: d.numero_carga,
          tipo_divergencia: d.tipo_divergencia,
          colaborador: d.colaborador,
          produto: d.produto,
          quantidade: d.quantidade,
          upload_id: uploadId,
          created_at: new Date()
        });
        totalOcorrencias++;

        // flush batch
        if (occurrences.length >= BATCH_SIZE) {
          await prisma.tbOcorrencia.createMany({ data: occurrences });
          occurrences.length = 0;
        }
      }
    }

    if (occurrences.length > 0) {
      await prisma.tbOcorrencia.createMany({ data: occurrences });
    }

    await prisma.tbUpload.update({ where: { id: uploadId }, data: {
      status: 'PROCESSADO',
      total_lines: parsed.length,
      total_cargas: cargasSet.size,
      total_ocorrencias: totalOcorrencias,
      processed_at: new Date()
    }});

    // Invalidate simple cache pattern - production should use tag-based or targeted keys
    try {
      await connection.publish('invalidate:dashboard', 'all');
    } catch (e) {
      // ignore
    }

  } catch (err) {
    await prisma.tbUpload.update({ where: { id: uploadId }, data: { status: 'ERRO', error_message: String(err) } });
    console.error('Worker error', err);
    throw err;
  }
});

worker.on('completed', job => {
  console.log('Job completed', job.id);
});

worker.on('failed', (job, err) => {
  console.error('Job failed', job?.id, err);
});

export default worker;
