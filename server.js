const express = require('express');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const cors = require('cors');
const fs = require('fs').promises;
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use('/downloads', express.static(path.join(__dirname, 'downloads')));

// ایجاد پوشه downloads
const ensureDownloadsDir = async () => {
  const downloadsDir = path.join(__dirname, 'downloads');
  try {
    await fs.access(downloadsDir);
  } catch {
    await fs.mkdir(downloadsDir);
  }
};

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({
    status: 'OK',
    uptime: process.uptime(),
    timestamp: new Date().toISOString()
  });
});

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    message: 'DOCX Converter API',
    version: '1.0.0',
    endpoints: {
      health: '/health',
      webhook: '/webhook (POST)',
      download: '/downloads/:filename'
    }
  });
});

// Webhook endpoint برای تبدیل متن به DOCX
app.post('/webhook', async (req, res) => {
  try {
    const { text } = req.body;

    if (!text) {
      return res.status(400).json({
        success: false,
        error: 'Text field is required'
      });
    }

    console.log('📝 Converting text to DOCX...');

    // ایجاد سند DOCX
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: text,
                size: 24,
                font: 'Arial'
              })
            ]
          })
        ]
      }]
    });

    // تولید نام فایل یکتا
    const timestamp = new Date().toISOString().replace(/[-:.]/g, '').slice(0, 15);
    const filename = `document_${timestamp}.docx`;
    const filepath = path.join(__dirname, 'downloads', filename);

    // تبدیل سند به buffer
    const buffer = await Packer.toBuffer(doc);

    // ذخیره فایل
    await fs.writeFile(filepath, buffer);

    console.log(`✅ DOCX file created: ${filename}`);

    // پاسخ موفقیت‌آمیز
    res.json({
      success: true,
      message: 'DOCX file generated successfully',
      filename: filename,
      downloadUrl: `${req.protocol}://${req.get('host')}/downloads/${filename}`,
      size: buffer.length
    });

  } catch (error) {
    console.error('❌ Error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to generate DOCX file',
      details: error.message
    });
  }
});

// شروع سرور
const startServer = async () => {
  try {
    await ensureDownloadsDir();
    
    app.listen(PORT, '0.0.0.0', () => {
      console.log('🚀 DOCX Converter Server started successfully!');
      console.log(`📍 Server running on: http://0.0.0.0:${PORT}`);
      console.log(`🔗 Webhook URL: http://0.0.0.0:${PORT}/webhook`);
      
      // Heartbeat
      setInterval(() => {
        console.log(`💓 Server heartbeat - Uptime: ${Math.floor(process.uptime())}s`);
      }, 60000);
    });
  } catch (error) {
    console.error('❌ Failed to start server:', error);
    process.exit(1);
  }
};

startServer();
