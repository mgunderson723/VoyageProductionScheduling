FROM node:18-alpine

WORKDIR /app

COPY package.json ./
RUN npm install --production --no-package-lock

COPY server.js ./
COPY scripts/ ./scripts/

EXPOSE 3000

CMD ["node", "server.js"]
