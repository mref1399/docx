const express = require('express');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const cors = require('cors');
const fs = require('fs').promises;
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use('/downloads', express.static(path.join(__dirname, 'downloads')));

const ensureDownloadsDir = async () => {
  const downloadsDir = path.join(__dirname, 'downloads');
  try {
    await fs.access(downloadsDir);
  } catch {
    await fs.mkdir(downloadsDir);
  }
};

app.get('/health', (req, res) => {
  res.json({
    status: 'OK',
    uptime: process.uptime()
  });
});

app.get('/', (req, res) => {
  res.json({
    message: 'DOCX Converter API',
    version: '1.0.0'
  });
});

app.post('/webhook', async (req, res) => {
  try {
    const { text } = req.body;
    
    if (!text) {
      return res.status(400).json({
        success: false,
        error: 'Text field is required'
      });
    }

    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: text,
                size: 24
              })
            ]
          })
        ]
      }]
    });

    const timestamp = Date.now();
    const filename = `document_${timestamp}.docx`;
    const filepath = path.join(__dirname, 'downloads', filename);
    const buffer = await Packer.toBuffer(doc);
    
    await fs.writeFile(filepath, buffer);

    res.json({
      success: true,
      message: 'DOCX file generated',
      filename: filename,
      downloadUrl: `${req.protocol}://${req.get('host')}/downloads/${filename}`
    });

  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

const startServer = async () => {
  await ensureDownloadsDir();
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
  });
};

startServer();
