FROM mcr.microsoft.com/playwright:v1.58.0-noble

WORKDIR /app

ENV MSTEAMS_BROWSER_CHANNEL=chromium

COPY package*.json ./
RUN npm ci

COPY tsconfig.json ./
COPY src ./src
RUN npm run build && npm prune --omit=dev

ENV NODE_ENV=production

RUN mkdir -p /home/pwuser/.teams-mcp-server \
  && chown -R pwuser:pwuser /home/pwuser/.teams-mcp-server

USER pwuser
VOLUME ["/home/pwuser/.teams-mcp-server"]

ENTRYPOINT ["node", "dist/index.js"]
