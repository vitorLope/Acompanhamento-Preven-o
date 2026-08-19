### Test / sample instructions

Added a small script to generate sample Excel files for local testing of the parser and worker.

To generate samples:

1. node scripts/generate-samples.js

This will create ./samples/sample-consinco.xlsx and ./samples/sample-infor.xlsx which you can upload via the POST /api/upload endpoint to exercise the worker.

Notes:
- Ensure dependencies (exceljs, bullmq, ioredis, prisma client) are installed.
- Run worker and api as described in README.
