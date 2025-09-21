FROM node:20-alpine

# ساخت پوشه اپ
WORKDIR /app

# نصب dependency‌ها
COPY package*.json ./
RUN npm ci --only=production

# اضافه کردن سورس
COPY . .

# پورتی که اپ گوش می‌ده
EXPOSE 3000

# دستور اجرا
CMD ["node", "src/webhook.js"]
