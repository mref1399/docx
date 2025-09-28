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

const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir);

// --- هدینگ بر اساس ستاره ---
function isHeading(text) {
    return /^\*{2,6}\s*/.test(text.trim());
}

function getHeadingLevelByStars(text) {
    const match = text.match(/^\*{2,6}/);
    if (!match) return 0;
    const starCount = match[0].length;
    const map = { 2: 1, 3: 2, 4: 3, 5: 4, 6: 5 };
    return map[starCount] || 0;
}

function cleanHeadingText(text) {
    return text.replace(/^\*{2,6}\s*/, '');
}

// --- ساخت TextRun با تشخیص فونت فارسی/انگلیسی ---
function createRunsWithAutoFontSwitch(line) {
    const runs = [];
    let buffer = '';
    let currentScript = null;
    let isFirstRun = true;

    const flushBuffer = () => {
        if (!buffer) return;
        const isPersian = currentScript === 'fa';
        runs.push(new TextRun({
            text: buffer,
            size: 28,
            font: isPersian
                ? { ascii: 'Times New Roman', hansi: 'Times New Roman', cs: 'B Nazanin' }
                : { ascii: 'Times New Roman', hansi: 'Times New Roman', cs: 'Times New Roman' }
        }));
        buffer = '';
    };

    let i = 0;
    while (i < line.length) {
        const char = line[i];
        const code = char.charCodeAt(0);
        let script;

        if (
            (code >= 0x0600 && code <= 0x06FF) ||
            (code >= 0x0750 && code <= 0x077F) ||
            (code >= 0xFB50 && code <= 0xFDFF) ||
            (code >= 0xFE70 && code <= 0xFEFF)
        ) {
            script = 'fa';
        } else {
            script = 'lat';
        }

        if (script !== currentScript) {
            flushBuffer();
            currentScript = script;
            if (isFirstRun && currentScript === 'fa') {
                buffer += '\u200F';
            }
            isFirstRun = false;
        }

        buffer += char;
        i++;
    }
    flushBuffer();
    return runs;
}

// --- پارس متن به پاراگراف‌ها ---
function parseTextToParagraphs(text) {
    const lines = text.split('\n');
    const paragraphs = [];

    for (let line of lines) {
        line = line.trim();
        if (line === '') {
            paragraphs.push(new Paragraph({
                children: [new TextRun({ text: '' })],
                spacing: { after: 0 }
            }));
            continue;
        }

        if (line.startsWith('$$')) {
            const formula = line.replace(/^\$\$\s*/, '');
            paragraphs.push(new Paragraph({
                children: [new Math({ children: [new MathRun(formula)] })],
                alignment: AlignmentType.JUSTIFIED,
                rightToLeft: true,
                bidirectional: true,
                spacing: { line: 240 }
            }));
            continue;
        }

        if (isHeading(line)) {
            const level = getHeadingLevelByStars(line);
            const headingText = cleanHeadingText(line);
            const runs = createRunsWithAutoFontSwitch(headingText).map(run => {
                run.bold(); // عنوان هدینگ همیشه بولد
                return run;
            });

            paragraphs.push(new Paragraph({
                children: runs,
                alignment: AlignmentType.JUSTIFIED,
                rightToLeft: true,
                bidirectional: true,
                spacing: { line: 240 },
                heading: HeadingLevel[`HEADING_${level}`] || HeadingLevel.HEADING_5
            }));
        } else {
            paragraphs.push(new Paragraph({
                children: createRunsWithAutoFontSwitch(line),
                style: 'Normal',
                alignment: AlignmentType.JUSTIFIED,
                rightToLeft: true,
                bidirectional: true,
                spacing: { line: 240, after: 0, before: 0 },
                indent: { firstLine: 708 }
            }));
        }
    }
    return paragraphs;
}

// --- API وبهوک ---
app.post('/webhook', async (req, res) => {
    try {
        const { text } = req.body;
        if (!text) return res.status(400).json({ error: 'Text is required', success: false });

        const paragraphs = parseTextToParagraphs(text);

        const doc = new Document({
            sections: [{
                properties: { bidirectional: true },
                children: paragraphs
            }],
            styles: {
                default: {
                    document: {
                        run: {
                            size: 28,
                            font: { ascii: 'Times New Roman', hansi: 'Times New Roman', cs: 'B Nazanin' }
                        },
                        paragraph: {
                            alignment: AlignmentType.JUSTIFIED,
                            rightToLeft: true,
                            bidirectional: true,
                            spacing: { line: 240, after: 0, before: 0 },
                            indent: { firstLine: 708 }
                        }
                    }
                },
                heading1: { run: { bold: true } },
                heading2: { run: { bold: true } },
                heading3: { run: { bold: true } },
                heading4: { run: { bold: true } },
                heading5: { run: { bold: true } }
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
        console.error(error);
        res.status(500).json({ error: 'Error creating file', success: false });
    }
});

app.get('/download/:filename', (req, res) => {
    const fileName = req.params.filename;
    const filePath = path.join(uploadsDir, fileName);
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found', success: false });

    res.download(filePath);
});

app.listen(port, () => console.log(`Server running on port ${port}`));
