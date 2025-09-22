const express = require('express');
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = require('docx');
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

// Function to detect if text is a heading (starts with #)
function isHeading(text) {
  return text.trim().startsWith('#');
}

// Function to get heading level
function getHeadingLevel(text) {
  const match = text.match(/^#+/);
  if (match) {
    return Math.min(match[0].length, 6);
  }
  return 0;
}

// Function to clean heading text (remove # symbols)
function cleanHeadingText(text) {
  return text.replace(/^#+\s*/, '');
}

// Function to parse text and create paragraphs
function parseTextToParagraphs(text) {
  const lines = text.split('\n');
  const paragraphs = [];

  for (let line of lines) {
    line = line.trim();
    
    if (line === '') {
      // Empty line with minimal spacing
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: '' })],
          spacing: { after: 0 }
        })
      );
      continue;
    }

    if (isHeading(line)) {
      const level = getHeadingLevel(line);
      const headingText = cleanHeadingText(line);
      
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: headingText,
              bold: true,
              size: 32 - (level * 2),
              font: {
                ascii: 'Times New Roman',
                eastAsia: 'Times New Roman',
                hansi: 'Times New Roman',
                cs: 'B Nazanin'
              }
            })
          ],
          alignment: AlignmentType.JUSTIFIED,
          rightToLeft: true,
          bidirectional: true,
          spacing: {
            before: 200,
            after: 100,
            line: 360,
            lineRule: 'auto'
          },
          indent: {
            firstLine: 708
          },
          heading: level === 1 ? HeadingLevel.HEADING_1 : 
                  level === 2 ? HeadingLevel.HEADING_2 :
                  level === 3 ? HeadingLevel.HEADING_3 :
                  level === 4 ? HeadingLevel.HEADING_4 :
                  level === 5 ? HeadingLevel.HEADING_5 :
                  HeadingLevel.HEADING_6
        })
      );
    } else {
      // Normal paragraph with explicit RTL and minimal spacing
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              size: 28,
              font: {
                ascii: 'Times New Roman',
                eastAsia: 'Times New Roman', 
                hansi: 'Times New Roman',
                cs: 'B Nazanin'
              }
            })
          ],
          style: 'Normal',
          alignment: AlignmentType.JUSTIFIED,
          rightToLeft: true,
          bidirectional: true,
          spacing: {
            line: 240,
            lineRule: 'auto',
            after: 0,      // No spacing after paragraph
            before: 0      // No spacing before paragraph
          },
          indent: {
            firstLine: 708
          }
        })
      );
    }
  }

  return paragraphs;
}

// Webhook endpoint
app.post('/webhook', async (req, res) => {
  try {
    const { text } = req.body;

    if (!text) {
      return res.status(400).json({ 
        error: 'Text is required',
        success: false
      });
    }

    const paragraphs = parseTextToParagraphs(text);

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440, 
              bottom: 1440,
              left: 1440
            }
          },
          type: 'nextPage',
          bidirectional: true
        },
        children: paragraphs
      }],
      styles: {
        default: {
          document: {
            run: {
              size: 28,
              font: {
                ascii: 'Times New Roman',
                eastAsia: 'Times New Roman',
                hansi: 'Times New Roman', 
                cs: 'B Nazanin'
              }
            },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
              rightToLeft: true,
              bidirectional: true,
              spacing: {
                line: 240,
                lineRule: 'auto',
                after: 0,
                before: 0
              },
              indent: {
                firstLine: 708
              }
            }
          }
        },
        paragraphStyles: [
          {
            id: 'Normal',
            name: 'Normal',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              size: 28,
              font: {
                ascii: 'Times New Roman',
                eastAsia: 'Times New Roman',
                hansi: 'Times New Roman',
                cs: 'B Nazanin'
              }
            },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
              rightToLeft: true,
              bidirectional: true,
              spacing: {
                line: 240,
                lineRule: 'auto',
                after: 0,     // Zero spacing between paragraphs
                before: 0
              },
              indent: {
                firstLine: 708
              }
            }
          }
        ]
      }
    });

    const fileName = `document_${Date.now()}.docx`;
    const filePath = path.join(uploadsDir, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    res.json({
      success: true,
      downloadUrl: `https://docx.darkube.app/download/${fileName}`,
      fileName: fileName,
      message: "DOCX file with proper RTL, Normal style and minimal paragraph spacing",
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

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({
        error: 'File not found',
        success: false
      });
    }

    res.download(filePath, fileName, (err) => {
      if (err) {
        console.error('Error sending file:', err);
      } else {
        setTimeout(() => {
          try {
            fs.unlinkSync(filePath);
            console.log(`File ${fileName} deleted`);
          } catch (deleteError) {
            console.error('Error deleting file:', deleteError);
          }
        }, 60000);
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

// List available files
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
