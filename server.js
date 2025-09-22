const express = require('express');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));

// Route for health check
app.get('/', (req, res) => {
    res.json({ 
        message: 'Text to Word Converter API is running!',
        status: 'OK',
        timestamp: new Date().toISOString()
    });
});

// Route for converting text to Word
app.post('/convert-to-word', async (req, res) => {
    try {
        const { text, filename } = req.body;

        if (!text) {
            return res.status(400).json({ 
                error: 'متن ارسالی خالی است',
                success: false 
            });
        }

        // Create a new Word document
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: text,
                                size: 24, // 12pt font size
                                font: "Arial"
                            })
                        ]
                    })
                ]
            }]
        });

        // Generate the document buffer
        const buffer = await Packer.toBuffer(doc);

        // Set response headers for file download
        const fileName = filename || `document_${Date.now()}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Length', buffer.length);

        // Send the buffer
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

// Start server
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
