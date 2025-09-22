const express = require('express');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));

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

// Error handling
process.on('uncaughtException', (err) => {
    console.error('Uncaught Exception:', err);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection:', reason);
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
            convert: '/convert-to-word'
        }
    });
});

// Convert route
app.post('/convert-to-word', async (req, res) => {
    try {
        console.log('Conversion request received at:', new Date().toISOString());
        const { text, filename } = req.body;

        if (!text) {
            console.log('Empty text received');
            return res.status(400).json({ 
                error: 'متن ارسالی خالی است',
                success: false 
            });
        }

        console.log('Creating Word document for text length:', text.length);
        
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

        console.log('Generating buffer...');
        const buffer = await Packer.toBuffer(doc);

        const fileName = filename || `document_${Date.now()}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Length', buffer.length);

        console.log(`Document created successfully: ${fileName}, size: ${buffer.length} bytes`);
        res.send(buffer);

    } catch (error) {
        console.error('Error converting text to Word:', error);
        res.status(500).json({ 
            error: 'خطا در تبدیل متن به ورد',
            details: error.message,
            success: false 
        });
    }
});

// Start server and keep it running
const server = app.listen(port, '0.0.0.0', () => {
    console.log(`🚀 DOCX Converter Server started successfully!`);
    console.log(`📍 Server running on: http://0.0.0.0:${port}`);
    console.log(`🕐 Started at: ${new Date().toISOString()}`);
    console.log(`📊 Memory usage: ${JSON.stringify(process.memoryUsage())}`);
    console.log(`⚡ Node version: ${process.version}`);
    console.log(`🔧 Environment: ${process.env.NODE_ENV || 'development'}`);
    
    // Keep alive mechanism
    setInterval(() => {
        console.log(`💓 Server heartbeat - Uptime: ${Math.floor(process.uptime())}s - Memory: ${Math.round(process.memoryUsage().rss / 1024 / 1024)}MB`);
    }, 60000); // هر دقیقه یک بار
});

// Export for testing
module.exports = app;
