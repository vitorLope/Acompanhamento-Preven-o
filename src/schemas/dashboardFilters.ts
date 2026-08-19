import { z } from 'zod';

export const DashboardFiltersSchema = z.object({
  dataInicial: z.string().optional(),
  dataFinal: z.string().optional(),
  cd: z.string().optional(),
  turno: z.string().optional(),
  tipoAtividade: z.string().optional(),
  page: z.coerce.number().positive().default(1),
  pageSize: z.coerce.number().positive().max(100).default(20),
});
