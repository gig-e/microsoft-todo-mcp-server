FROM node:22-slim AS builder
WORKDIR /app

RUN corepack enable

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
RUN CI=true pnpm install --frozen-lockfile

COPY . .
RUN pnpm run build

FROM node:22-slim AS runtime
WORKDIR /app
ENV NODE_ENV=production

RUN corepack enable

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
RUN CI=true pnpm install --frozen-lockfile --prod

COPY --from=builder /app/dist ./dist

# Container Apps injects PORT; src/http-server.ts reads it (defaults to 8080 locally).
CMD ["node", "dist/http-server.js"]
