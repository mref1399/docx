FROM node:18-alpine

# نصب build tools
RUN apk add --no-cache python3 make g++

WORKDIR /app

# کپی package files
COPY package*.json ./

# نصب dependencies
RUN npm ci --only=production

# کپی کد
COPY . .

# تنظیم متغیرهای محیطی
ENV NODE_ENV=production
ENV PORT=3000

# expose port
EXPOSE 3000

# دستور اجرا (مهم!)
CMD ["node", "server.js"]
