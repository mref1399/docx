const express = require('express');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(express.json());

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({
    status: 'OK',
    message: 'DOCX Converter API',
    version: '1.0.0',
    uptime: process.uptime()
  });
});

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    message: 'DOCX Converter API',
    version: '1.0.0'
  });
});

// Create uploads directory
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir);
}

// Webhook endpoint - now returns JSON
app.post('/webhook', async (req, res) => {
  try {
    const { text } = req.body;

    if (!text) {
      return res.status(400).json({ 
        error: 'Text is required',
        success: false
      });
    }

    // Create DOCX document
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
              }),
            ],
          }),
        ],
      }],
    });

    // Generate unique filename
    const fileName = `document_${Date.now()}.docx`;
    const filePath = path.join(uploadsDir, fileName);

    // Save file
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    // Return JSON response with download link
    res.json({
      success: true,
      downloadUrl: `https://docx.darkube.app/download/${fileName}`,
      fileName: fileName,
      message: "DOCX file created successfully",
      fileSize: buffer.length
    });

  } catch (error) {
    console.error('Error creating DOCX file:', error);
    res.status(500).json({ 
      error: 'Error creating file',
      success: false,
      details: error.message
    });
  }
});

// Download file route
app.get('/download/:filename', (req, res) => {
  try {
    const fileName = req.params.filename;
    const filePath = path.join(uploadsDir, fileName);

    // Check if file exists
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({
        error: 'File not found',
        success: false
      });
    }

    // Send file for download
    res.download(filePath, fileName, (err) => {
      if (err) {
        console.error('Error sending file:', err);
      } else {
        // Delete file after download (optional)
        setTimeout(() => {
          try {
            fs.unlinkSync(filePath);
            console.log(`File ${fileName} deleted`);
          } catch (deleteError) {
            console.error('Error deleting file:', deleteError);
          }
        }, 60000); // Delete after 1 minute
      }
    });

  } catch (error) {
    console.error('Error downloading file:', error);
    res.status(500).json({
      error: 'Error downloading file',
      success: false
    });
  }
});

// List available files (for debugging)
app.get('/files', (req, res) => {
  try {
    const files = fs.readdirSync(uploadsDir);
    res.json({
      files: files,
      count: files.length
    });
  } catch (error) {
    res.status(500).json({
      error: 'Error reading files',
      details: error.message
    });
  }
});

// Start server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
  console.log(`URL: http://localhost:${port}`);
});
