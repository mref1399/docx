# استفاده از Node.js 18 Alpine
FROM node:18-alpine

# نصب dependencies سیستمی
RUN apk add --no-cache python3 make g++

# تنظیم working directory
WORKDIR /app

# کپی فایل‌های package
COPY package*.json ./

# نصب dependencies (بجای npm ci از npm install استفاده می‌کنیم)
RUN npm install --only=production

# کپی کد
COPY . .

# تنظیم متغیرهای محیطی
ENV NODE_ENV=production
ENV PORT=3000

# expose کردن پورت
EXPOSE 3000

# اجرای سرور
CMD ["node", "server.js"]
