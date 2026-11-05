FROM node:22-bookworm-slim
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --no-audit --no-fund
COPY src ./src
RUN mkdir -p /app/data && chown node:node /app/data
USER node
ENV NODE_ENV=production
CMD ["node", "src/index.js"]
