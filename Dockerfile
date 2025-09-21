FROM node:20-alpine

WORKDIR /app

# ارتقای npm (اینجا سازگار خواهد بود)
RUN npm install -g npm@latest

COPY package*.json ./
RUN npm install

COPY . .

CMD ["node", "index.js"]
