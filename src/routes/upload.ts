import express from 'express';
import multer from 'multer';
import prisma from '../prisma';
import { queue } from '../lib/queue';

const router = express.Router();
const upload = multer({ dest: process.env.UPLOAD_TEMP_DIR || '/tmp/uploads' });

router.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    if (!file) return res.status(400).json({ error: 'Arquivo não enviado' });
    if (!/\.(xlsx|xls)$/i.test(file.originalname)) {
      return res.status(400).json({ error: 'Extensão inválida' });
    }

    const created = await prisma.tbUpload.create({
      data: {
        filename: file.originalname,
        filepath: file.path,
        status: 'RECEBIDO'
      }
    });

    await queue.add('process-upload', { uploadId: created.id, filepath: file.path });

    res.status(202).json({ uploadId: created.id, status: 'RECEBIDO' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Erro interno' });
  }
});

export default router;
