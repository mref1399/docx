const express = require('express');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// Graceful shutdown handling
process.on('SIGTERM', () => {
    console.log('SIGTERM received, shutting down gracefully');
    server.close(() => {
        console.log('Process terminated');
    });
});

process.on('SIGINT', () => {
    console.log('SIGINT received, shutting down gracefully');
    server.close(() => {
        console.log('Process terminated');
    });
});

// Health check route
app.get('/health', (req, res) => {
    res.status(200).json({ 
        status: 'OK',
        timestamp: new Date().toISOString(),
        uptime: process.uptime(),
        memory: process.memoryUsage()
    });
});

// Root route
app.get('/', (req, res) => {
    console.log('Root endpoint accessed');
    res.json({ 
        message: 'DOCX Converter API is running!',
        status: 'OK',
        timestamp: new Date().toISOString(),
        endpoints: {
            health: '/health',
            webhook: '/webhook',
            convert: '/convert-to-word'
        }
    });
});

// Webhook endpoint برای دریافت از n8n
app.post('/webhook', async (req, res) => {
    try {
        console.log('📨 Webhook received from n8n at:', new Date().toISOString());
        console.log('📊 Request body:', JSON.stringify(req.body, null, 2));

        // استخراج متن از درخواست n8n
        let text = '';
        let filename = '';

        // چندین روش برای استخراج متن
        if (req.body.text) {
            text = req.body.text;
        } else if (req.body.data && req.body.data.text) {
            text = req.body.data.text;
        } else if (req.body.content) {
            text = req.body.content;
        } else if (req.body.message) {
            text = req.body.message;
        } else {
            // اگر متن در جای دیگری باشد، کل body رو string می‌کنیم
            text = JSON.stringify(req.body, null, 2);
        }

        // استخراج نام فایل
        if (req.body.filename) {
            filename = req.body.filename;
        } else if (req.body.name) {
            filename = req.body.name;
        } else {
            filename = `document_${Date.now()}.docx`;
        }

        if (!text || text.trim() === '') {
            console.log('❌ Empty text received');
            return res.status(400).json({ 
                error: 'متن ارسالی خالی است',
                success: false,
                received_data: req.body
            });
        }

        console.log('✅ Text extracted, length:', text.length);
        console.log('📄 Filename:', filename);

        // ساخت سند Word
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: text,
                                size: 24,
                                font: "Arial"
                            })
                        ]
                    })
                ]
            }]
        });

        console.log('📄 Creating Word document...');
        const buffer = await Packer.toBuffer(doc);

        // تنظیم headers برای دانلود
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Length', buffer.length);

        console.log(`✅ Document created: ${filename}, size: ${buffer.length} bytes`);
        
        // ارسال فایل
        res.send(buffer);

    } catch (error) {
        console.error('❌ Error in webhook:', error);
        res.status(500).json({ 
            error: 'خطا در پردازش درخواست',
            details: error.message,
            success: false 
        });
    }
});

// Convert route (برای سازگاری با قبل)
app.post('/convert-to-word', async (req, res) => {
    try {
        console.log('📨 Convert request received at:', new Date().toISOString());
        const { text, filename } = req.body;

        if (!text) {
            console.log('❌ Empty text received');
            return res.status(400).json({ 
                error: 'متن ارسالی خالی است',
                success: false 
            });
        }

        console.log('✅ Creating Word document for text length:', text.length);
        
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: text,
                                size: 24,
                                font: "Arial"
                            })
                        ]
                    })
                ]
            }]
        });

        console.log('📄 Generating buffer...');
        const buffer = await Packer.toBuffer(doc);

        const fileName = filename || `document_${Date.now()}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Length', buffer.length);

        console.log(`✅ Document created: ${fileName}, size: ${buffer.length} bytes`);
        res.send(buffer);

    } catch (error) {
        console.error('❌ Error converting text to Word:', error);
        res.status(500).json({ 
            error: 'خطا در تبدیل متن به ورد',
            details: error.message,
            success: false 
        });
    }
});

// Start server
const server = app.listen(port, '0.0.0.0', () => {
    console.log(`🚀 DOCX Converter Server started successfully!`);
    console.log(`📍 Server running on: http://0.0.0.0:${port}`);
    console.log(`🔗 Webhook URL: http://0.0.0.0:${port}/webhook`);
    console.log(`🕐 Started at: ${new Date().toISOString()}`);
    
    // Keep alive
    setInterval(() => {
        console.log(`💓 Server heartbeat - Uptime: ${Math.floor(process.uptime())}s`);
    }, 60000);
});

module.exports = app;

### قدم ۲: تنظیم n8n Workflow

حالا در n8n این workflow رو بساز:

1. **شروع**: `Manual Trigger` یا هر trigger دیگه
2. **HTTP Request Node**: 
   - **Method**: `POST`
   - **URL**: `https://docx-api-[شناسه-تو].darkube.app/webhook`
   - **Body**:
```json
   {
"text": "{{ $json.your_text_field }}",
"filename": "my-document.docx"
   }
   
### قدم ۳: مثال کامل n8n workflow

```json
{
  "name": "DOCX Converter Webhook",
  "nodes": [
    {
      "parameters": {},
      "name": "Manual Trigger",
      "type": "n8n-nodes-base.manualTrigger",
      "typeVersion": 1,
      "position": [240, 300]
    },
    {
      "parameters": {
        "values": {
          "string": [
            {
              "name": "text",
              "value": "سلام! این یک متن نمونه برای تست تبدیل به ورد است.\n\nاین سرویس توسط n8n و DarkubeCCE راه‌اندازی شده است."
            },
            {
              "name": "filename", 
              "value": "test-document.docx"
            }
          ]
        }
      },
      "name": "Set Text Data",
      "type": "n8n-nodes-base.set",
      "typeVersion": 1,
      "position": [460, 300]
    },
    {
      "parameters": {
        "url": "https://docx-api-[شناسه-تو].darkube.app/webhook",
        "options": {
          "bodyContentType": "json",
          "body": {
            "text": "={{ $json.text }}",
            "filename": "={{ $json.filename }}"
          }
        }
      },
      "name": "Send to DOCX API",
      "type": "n8n-nodes-base.httpRequest", 
      "typeVersion": 1,
      "position": [680, 300]
    }
  ],
  "connections": {
    "Manual Trigger": {
      "main": [
        [
          {
            "node": "Set Text Data",
            "type": "main",
            "index": 0
          }
        ]
      ]
    },
    "Set Text Data": {
      "main": [
        [
          {
            "node": "Send to DOCX API", 
            "type": "main",
            "index": 0
          }
        ]
      ]
    }
  }
}

### قدم ۴: تست با curl

bash
curl -X POST https://docx-api-[شناسه-تو].darkube.app/webhook \
  -H "Content-Type: application/json" \
  -d '{
    "text": "سلام! این متن از n8n ارسال شده است.",
    "filename": "test-from-n8n.docx"
  }' \
  --output test-file.docx

حالا فایل‌ها رو commit کن و بهم بگو تا مرحله بعدی رو راهنماییت کنم! 🚀
