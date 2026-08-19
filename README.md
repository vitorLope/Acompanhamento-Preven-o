# Backend skeleton for Acompanhamento-Preven-o

Scripts and basic dependencies to start the backend. This commit creates the core files: Prisma schema, queue, worker, upload route and a simple Excel parser to detect CONSINCO/INFOR layouts.

Run locally:

1. copy .env.example to .env and set DATABASE_URL and REDIS_URL
2. docker-compose up -d
3. pnpm install (or npm install)
4. npx prisma migrate dev --name init
5. pnpm build && pnpm start

