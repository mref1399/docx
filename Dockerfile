FROM node:18-alpine

WORKDIR /app

COPY package*.json ./
RUN npm ci --only=production

COPY . .

EXPOSE 3000

RUN addgroup app && adduser -S -G app app
USER app

CMD ["node", "index.js"]
