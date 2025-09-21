FROM node:18-alpine
RUN npm install -g npm@latest
WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

EXPOSE 3000

RUN addgroup app && adduser -S -G app app
USER app

CMD ["node", "index.js"]
