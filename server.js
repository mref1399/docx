const express = require('express');
const {
    Document, Packer, Paragraph, TextRun,
    AlignmentType, HeadingLevel
} = require('docx');

const app = express();
const port = process.env.PORT || 3000;

app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ limit: '100mb', extended: true }));

const LRM = '\u200E'; // Left-to-right mark
const RLM = '\u200F'; // Right-to-left mark

function isHeading(text) { return text.trim().startsWith('#'); }
function getHeadingLevel(text) {
    const match = text.match(/^#+/);
    return match ? Math.min(match[0].length, 6) : 0;
}
function cleanHeadingText(text) {
    return text.replace(/^#+\s*/, '').replace(/\s*#+$/, '');
}

function isParagraphRTL(line) {
    const persianCount = (line.match(/[\u0600-\u06FF]/g) || []).length;
    const latinCount = (line.match(/[A-Za-z0-9]/g) || []).length;
    return persianCount >= latinCount;
}

function detectScript(char) {
    const code = char.charCodeAt(0);
    if (
        (code >= 0x0600 && code <= 0x06FF) ||
        (code >= 0x0750 && code <= 0x077F) ||
        (code >= 0xFB50 && code <= 0xFDFF) ||
        (code >= 0xFE70 && code <= 0xFEFF)
    ) {
        return 'fa';
    }
    return 'lat';
}

function createRunsWithAutoFontSwitch(line) {
    const runs = [];
    let buffer = '';
    let currentScript = null;
    let boldMode = false;

    const flushBuffer = () => {
        if (!buffer) return;
        const isPersian = currentScript === 'fa';
        runs.push(new TextRun({
            text: (isPersian ? RLM : LRM) + buffer,
            bold: boldMode,
            size: 28,
            font: isPersian
                ? { ascii: 'Times New Roman', hAnsi: 'Times New Roman', cs: 'B Nazanin' }
                : { ascii: 'Times New Roman', hAnsi: 'Times New Roman', cs: 'Times New Roman' },
            rightToLeft: isPersian,
            bidirectional: true
        }));
        buffer = '';
    };

    const pushScriptedRun = (text, type) => {
        if (!text) return;
        const script = detectScript(text[0]);
        const isPersian = script === 'fa';
        runs.push(new TextRun({
            text: (isPersian ? RLM : LRM) + text,
            bold: boldMode,
            size: 28,
            font: isPersian
                ? { ascii: 'Times New Roman', hAnsi: 'Times New Roman', cs: 'B Nazanin' }
                : { ascii: 'Times New Roman', hAnsi: 'Times New Roman', cs: 'Times New Roman' },
            rightToLeft: isPersian,
            bidirectional: true,
            superScript: type === 'super',
            subScript: type === 'sub'
        }));
        currentScript = null;
    };

    let i = 0;
    while (i < line.length) {
        const char = line[i];

        // toggle bold when encountering double asterisks
        if (char === '*') {
            let starCount = 0;
            while (line[i] === '*') { starCount++; i++; }
            if (starCount >= 2) {
                flushBuffer();
                boldMode = !boldMode;
                continue;
            } else {
                buffer += '*'.repeat(starCount);
                continue;
            }
        }

        // super/sub-script detection (^... or _...)
        if (char === '^' || char === '_') {
            const type = char === '^' ? 'super' : 'sub';
            flushBuffer();
            i++;

            let value = '';
            if (line[i] === '{') {
                i++;
                while (i < line.length && line[i] !== '}') {
                    value += line[i];
                    i++;
                }
                if (line[i] === '}') i++; // skip closing brace
            } else if (i < line.length) {
                value = line[i];
                i++;
            }

            pushScriptedRun(value, type);
            continue;
        }

        // normal character handling with script switching
        const script = detectScript(char);
        if (script !== currentScript) {
            flushBuffer();
            currentScript = script;
        }

        buffer += char;
        i++;
    }

    flushBuffer();
    return runs;
}

function parseTextToParagraphs(text) {
    const lines = text.split('\n');
    const paragraphs = [];

    for (let rawLine of lines) {
        let line = rawLine.trim();

        if (line === '') {
            paragraphs.push(new Paragraph({ children: [new TextRun('')], spacing: { after: 0 } }));
            continue;
        }

        if (line.startsWith('$$')) {
            const formula = line.replace(/^\$\$\s*/, '');
            paragraphs.push(new Paragraph({
                children: [
                    new TextRun({
                        text: `${LRM}${formula}`,
                        size: 28,
                        font: { ascii: 'Cambria Math', hAnsi: 'Cambria Math' }
                    })
                ],
                alignment: AlignmentType.CENTER,
                rightToLeft: false,
                bidirectional: false,
                spacing: { before: 120, after: 120 }
            }));
            continue;
        }

        if (isHeading(line)) {
            const level = getHeadingLevel(line);
            const headingText = cleanHeadingText(line);
            const rtl = isParagraphRTL(headingText);
            const alignment = rtl ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;

            paragraphs.push(new Paragraph({
                children: createRunsWithAutoFontSwitch(headingText).map(run => run.bold()),
                alignment,
                heading: HeadingLevel[`HEADING_${level}`] || HeadingLevel.HEADING_6,
                rightToLeft: rtl,
                bidirectional: true,
                spacing: { line: 240 }
            }));
            continue;
        }

        const rtl = isParagraphRTL(line);
        const alignment = rtl ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;
        const indent = rtl ? { firstLine: 708 } : { left: 0 };

        paragraphs.push(new Paragraph({
            children: createRunsWithAutoFontSwitch(line),
            alignment,
            rightToLeft: rtl,
            bidirectional: true,
            spacing: { line: 240, after: 0, before: 0 },
            indent
        }));
    }

    return paragraphs;
}

app.post('/', async (req, res) => {
    try {
        const { text } = req.body;
        if (!text) return res.status(400).send('Text is required');

        const paragraphs = parseTextToParagraphs(text);

        const doc = new Document({
            sections: [{
                properties: { bidirectional: true },
                children: paragraphs
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        res.setHeader('Content-Disposition', 'attachment; filename=document.docx');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buffer);
    } catch (error) {
        console.error(error);
        res.status(500).send('Error generating file');
    }
});

app.listen(port, () => console.log(`✅ Server running on port ${port}`));
