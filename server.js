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

// امن: شناسایی سطح هدینگ بر اساس تعداد ستاره‌ها
function getHeadingLevelByStars(text) {
    if (!text) return 0;
    const match = text.match(/^\*+/);
    if (!match) return 0;
    const level = match[0].length - 1; // ** => Heading 1, *** => Heading 2 ...
    return level >= 1 && level <= 6 ? level : 0;
}

// حذف ستاره‌ها از متن عنوان
function cleanHeadingTextStars(text) {
    if (!text) return '';
    return text.replace(/^\*+\s*/, '');
}

// ایجاد TextRun با پشتیبانی فارسی/انگلیسی و بولد بین **
function createRunsWithAutoFontSwitch(line) {
    const runs = [];
    if (!line) return runs;

    let buffer = '';
    let currentScript = null;
    let isFirstRun = true;
    let boldMode = false;

    const flushBuffer = () => {
        if (!buffer) return;
        const isPersian = currentScript === 'fa';
        runs.push(
            new TextRun({
                text: buffer,
                bold: boldMode,
                size: 28,
                font: isPersian
                    ? { ascii: 'Times New Roman', hansi: 'Times New Roman', cs: 'B Nazanin' }
                    : { ascii: 'Times New Roman', hansi: 'Times New Roman', cs: 'Times New Roman' }
            })
        );
        buffer = '';
    };

    let i = 0;
    while (i < line.length) {
        if (line.startsWith('**', i)) {
            flushBuffer();
            boldMode = !boldMode;
            i += 2;
            continue;
        }

        const char = line[i];
        const code = char.charCodeAt(0);
        const script =
            (code >= 0x0600 && code <= 0x06FF) ||
            (code >= 0x0750 && code <= 0x077F) ||
            (code >= 0xFB50 && code <= 0xFDFF) ||
            (code >= 0xFE70 && code <= 0xFEFF)
                ? 'fa'
                : 'lat';

        if (script !== currentScript) {
            flushBuffer();
            currentScript = script;
            if (isFirstRun && currentScript === 'fa') buffer += '\u200F';
            isFirstRun = false;
        }

        buffer += char;
        i++;
    }
    flushBuffer();
    return runs;
}

// پردازش ورودی به آرایه Paragraph
function parseTextToParagraphs(text) {
    const lines = text.split('\n');
    const paragraphs = [];

    for (let rawLine of lines) {
        const line = (rawLine || '').trim();

        // خط خالی
        if (line === '') {
            paragraphs.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
            continue;
        }

        // فرمول
        if (line.startsWith('$$')) {
            const formula = line.replace(/^\$\$\s*/, '');
            try {
                paragraphs.push(
                    new Paragraph({
                        children: [new Math({ children: [new MathRun(formula)] })],
                        alignment: AlignmentType.JUSTIFIED,
                        rightToLeft: true,
                        bidirectional: true
                    })
                );
            } catch (err) {
                console.warn('Math parse error, fallback to text:', formula);
                paragraphs.push(
                    new Paragraph({
                        children: createRunsWithAutoFontSwitch(formula),
                        alignment: AlignmentType.JUSTIFIED,
                        rightToLeft: true,
                        bidirectional: true
                    })
                );
            }
            continue;
        }

        // عنوان
        const headingLevel = getHeadingLevelByStars(line);
        if (headingLevel > 0) {
            const headingText = cleanHeadingTextStars(line);
            paragraphs.push(
                new Paragraph({
                    children: createRunsWithAutoFontSwitch(headingText).map(run => {
                        run.bold();
                        return run;
                    }),
                    heading: HeadingLevel[`HEADING_${headingLevel}`],
                    alignment: AlignmentType.JUSTIFIED,
                    rightToLeft: true,
                    bidirectional: true
                })
            );
        } else {
            // متن معمولی
            paragraphs.push(
                new Paragraph({
                    children: createRunsWithAutoFontSwitch(line),
                    style: 'Normal',
                    alignment: AlignmentType.JUSTIFIED,
                    rightToLeft: true,
                    bidirectional: true,
                    spacing: { line: 240, after: 0, before: 0 },
                    indent: { firstLine: 708 }
                })
            );
        }
    }

    // اگر هیچ پاراگرافی تولید نشد، پاراگراف پیش‌فرض اضافه کن
    if (paragraphs.length === 0) {
        paragraphs.push(
            new Paragraph({
                children: [new TextRun({ text: 'Empty document', bold: true })],
                alignment: AlignmentType.JUSTIFIED,
                rightToLeft: true,
                bidirectional: true
            })
        );
    }

    return paragraphs;
}

// مسیر webhook
app.post('/webhook', async (req, res) => {
    try {
        const { text } = req.body;
        if (!text || typeof text !== 'string') {
            console.warn('Invalid text input => fallback used');
        }

        const safeText = typeof text === 'string' ? text : 'Empty document';
        const paragraphs = parseTextToParagraphs(safeText);

        const doc = new Document({
            sections: [
                {
                    properties: { bidirectional: true },
                    children: paragraphs
                }
            ],
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
                }
            }
        });

        const fileName = `document_${Date.now()}.docx`;
        const filePath = path.join(uploadsDir, fileName);
        const buffer = await Packer.toBuffer(doc);

        fs.writeFileSync(filePath, buffer);

        res.json({
            success: true,
            downloadUrl: `/download/${fileName}`,
            fileName,
            fileSize: buffer.length
        });
    } catch (error) {
        console.error('Error creating file =>', error);
        res.status(500).json({ error: 'Error creating file', success: false });
    }
});

// مسیر دانلود
app.get('/download/:filename', (req, res) => {
    const filePath = path.join(uploadsDir, req.params.filename);
    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'File not found', success: false });
    }
    res.download(filePath);
});

app.listen(port, () => console.log(`Server running on port ${port}`));
