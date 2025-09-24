const express = require('express');
const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, HeadingLevel, Math, MathRun
} = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;
app.use(express.json());

app.get('/health', (req, res) => {
  res.json({ status: 'OK', message: 'DOCX Converter API', version: '1.0.0', uptime: process.uptime() });
});

const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);

function isHeading(text) { return text.trim().startsWith('#'); }
function getHeadingLevel(text) {
  const match = text.match(/^#+/);
  return match ? Math.min(match[0].length, 6) : 0;
}
function cleanHeadingText(text) { return text.replace(/^#+\s*/, ''); }

function baseFont() {
  return {
    ascii: 'Times New Roman',
    hansi: 'Times New Roman',
    cs: 'B Nazanin'
  };
}

function parseTextToParagraphs(text) {
  const lines = text.split('\n');
  const paragraphs = [];

  for (let line of lines) {
    line = line.trim();

    if (line === '') {
      paragraphs.push(new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 0 } }));
      continue;
    }

    // مثال: خطوطی که با $$ شروع شوند فرمول ریاضی هستند
    if (line.startsWith('$$')) {
      const formula = line.replace(/^\$\$\s*/, '');
      paragraphs.push(new Paragraph({
        children: [
          new Math({
            children: [new MathRun(formula)]
          })
        ],
        alignment: AlignmentType.JUSTIFIED,
        rightToLeft: true,
        spacing: { line: 240 }
      }));
      continue;
    }

    if (isHeading(line)) {
      const level = getHeadingLevel(line);
      const headingText = cleanHeadingText(line);
      paragraphs.push(new Paragraph({
        children: [
          new TextRun({
            text: headingText,
            bold: true,
            size: 28,
            font: baseFont()
          })
        ],
        alignment: AlignmentType.JUSTIFIED,
        rightToLeft: true,
        spacing: { before: 200, after: 100, line: 240 },
        indent: { firstLine: 708 },
        heading: level === 1 ? HeadingLevel.HEADING_1 :
                 level === 2 ? HeadingLevel.HEADING_2 :
                 level === 3 ? HeadingLevel.HEADING_3 :
                 level === 4 ? HeadingLevel.HEADING_4 :
                 level === 5 ? HeadingLevel.HEADING_5 : HeadingLevel.HEADING_6
      }));
    } else {
      paragraphs.push(new Paragraph({
        children: [
          new TextRun({
            text: line,
            size: 28,
            font: baseFont()
          })
        ],
        style: 'Normal',
        alignment: AlignmentType.JUSTIFIED,
        rightToLeft: true,
        spacing: { line: 240, after: 0, before: 0 },
        indent: { firstLine: 708 }
      }));
    }
  }

  return paragraphs;
}

app.post('/webhook', async (req, res) => {
  try {
    const { text } = req.body;
    if (!text) return res.status(400).json({ error: 'Text is required', success: false });

    const paragraphs = parseTextToParagraphs(text);

    const doc = new Document({
      sections: [{
        properties: {
          page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
          bidirectional: true
        },
        children: paragraphs
      }],
      styles: {
        default: {
          document: {
            run: { size: 28, font: baseFont() },
            paragraph: {
              alignment: AlignmentType.JUSTIFIED,
              rightToLeft: true,
              spacing: { line: 240, after: 0, before: 0 },
              indent: { firstLine: 708 }
            }
          }
        },
        paragraphStyles: [{
          id: 'Normal',
          name: 'Normal',
          run: { size: 28, font: baseFont() },
          paragraph: {
            alignment: AlignmentType.JUSTIFIED,
            rightToLeft: true,
            spacing: { line: 240, after: 0, before: 0 },
            indent: { firstLine: 708 }
          }
        }]
      }
    });

    const fileName = `document_${Date.now()}.docx`;
    const filePath = path.join(uploadsDir, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    res.json({
      success: true,
      downloadUrl: `https://docx.darkube.app/download/${fileName}`,
      fileName,
      fileSize: buffer.length
    });

  } catch (error) {
    console.error('Error creating DOCX file:', error);
    res.status(500).json({ error: 'Error creating file', success: false, details: error.message });
  }
});

app.get('/download/:filename', (req, res) => {
  try {
    const fileName = req.params.filename;
    const filePath = path.join(uploadsDir, fileName);
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found', success: false });

    res.download(filePath, fileName, (err) => {
      if (!err) {
        setTimeout(() => { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); }, 60000);
      }
    });
  } catch (error) {
    res.status(500).json({ error: 'Error downloading file', success: false });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
