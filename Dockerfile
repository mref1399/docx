# استفاده از Node.js نسخه 18 alpine (سبک‌تر)
FROM node:18-alpine

# تنظیم working directory
WORKDIR /app

# کپی کردن package.json و package-lock.json
COPY package*.json ./

# نصب dependencies
RUN npm install --only=production

# کپی کردن باقی فایل‌ها
COPY . .

# expose کردن پورت
EXPOSE 3000

# تنظیم متغیر محیطی برای پورت
ENV PORT=3000

# دستور اجرا
CMD ["npm", "start"]
