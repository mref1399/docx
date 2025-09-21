FROM node:20-alpine

WORKDIR /app

COPY package*.json ./

# نصب دقیق نسخه docx سازگار
RUN npm install docx@8.5.0 --production

COPY . .

EXPOSE 3000

CMD ["node", "index.js"]

