const express = require('express');
const {
    Document, Packer, Paragraph, TextRun,
    AlignmentType, HeadingLevel, Math, MathRun
} = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ limit: '100mb', extended: true }));

function isHeading(text) { return text.trim().startsWith('#'); }
function getHeadingLevel(text) {
    const match = text.match(/^#+/);
    return match ? Math.min(match[0].length, 6) : 0;
}
function cleanHeadingText(text) {
    return text.replace(/^#+\s*/, '').replace(/\s*#+$/, '');
}

function reverseBrackets(str) {
    return str.replace(/\[|\]/g, match => match === '[' ? ']' : '[');
}

function createRunsWithAutoFontSwitch(line) {
    const runs = [];
    let buffer = '';
    let currentScript = null;
    let isFirstRun = true;
    let boldMode = false;

    const flushBuffer = () => {
        if (!buffer) return;
        const processedText = reverseBrackets(buffer);
        const isPersian = currentScript === 'fa';
        runs.push(new TextRun({
            text: processedText,
            bold: boldMode,
            size: 28,
            font: isPersian
                ? { ascii: 'Times New Roman', hansi: 'Times New Roman', cs: 'B Nazanin' }
                : { ascii: 'Times New Roman', hansi: 'Times New Roman', cs: 'Times New Roman' },
            rightToLeft: true,
            bidirectional: true
        }));
        buffer = '';
    };

    let i = 0;
    while (i < line.length) {
        if (line[i] === '*') {
            let starCount = 0;
            while (line[i] === '*') {
                starCount++;
                i++;
            }
            if (starCount >= 2) {
                flushBuffer();
                boldMode = !boldMode;
                continue;
            } else {
                buffer += '*';
                continue;
            }
        }

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

function parseTextToParagraphs(text) {
    const lines = text.split('\n');
    const paragraphs = [];

    for (let line of lines) {
        line = line.trim();
        if (line === '') {
            paragraphs.push(new Paragraph({
                children: [new TextRun({ text: '', rightToLeft: true, bidirectional: true })],
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
            const level = getHeadingLevel(line);
            const headingText = cleanHeadingText(line);
            paragraphs.push(new Paragraph({
                children: createRunsWithAutoFontSwitch(headingText).map(run => {
                    run.bold();
                    return run;
                }),
                alignment: AlignmentType.JUSTIFIED,
                rightToLeft: true,
                bidirectional: true,
                spacing: { line: 240 },
                heading: HeadingLevel[`HEADING_${level}`] || HeadingLevel.HEADING_6
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

// Direct DOCX output on POST /
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
        res.status(500).send('Error creating file');
    }
});

app.listen(port, () => console.log(`Server running on port ${port}`));
