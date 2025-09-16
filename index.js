import express from "express";
import bodyParser from "body-parser";
import { Document, Packer, Paragraph, TextRun, AlignmentType, FootnoteReferenceRun, Footnote, Footnotes } from "docx";

const app = express();
app.use(bodyParser.json({ limit: "10mb" }));

// تابع تشخیص فارسی
function hasPersian(text) {
    return /[\u0600-\u06FF]/.test(text);
}

// تابع تشخیص انگلیسی
function hasEnglish(text) {
    return /[A-Za-z]/.test(text);
}

// تابع ساخت پاورقی‌ها
function createFootnotes(englishWords) {
    return englishWords.map((entry, i) =>
        new Footnote({
            id: i + 1,
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `${entry.farsi} = ${entry.english}`,
                            font: "B Nazanin",
                            size: 28
                        })
                    ]
                })
            ]
        })
    );
}

app.post("/to-docx", async (req, res) => {
    try {
        const { text } = req.body;
        if (!text) {
            return res.status(400).send({ error: "متن ارسال نشده است" });
        }

        const paragraphs = [];
        const englishRefs = [];
        let footnoteIndex = 1;

        for (let line of text.split("\n")) {
            const trimmed = line.trim();
            let alignment = AlignmentType.JUSTIFIED;

            // وسط چین کردن عناوین جدول یا شکل
            if (/^(جدول\s+\d+\s*-)/.test(trimmed) || /^(شکل\s+\d+\s*-)/.test(trimmed)) {
                alignment = AlignmentType.CENTER;
            }

            const runs = [];
            const words = trimmed.split(" ");

            for (let i = 0; i < words.length; i++) {
                const w = words[i];
                if (hasPersian(w)) {
                    runs.push(new TextRun({
                        text: w + " ",
                        font: "B Nazanin",
                        size: 28
                    }));
                } else if (hasEnglish(w)) {
                    const prevWord = words[i - 1] || "";
                    englishRefs.push({ farsi: prevWord, english: w });
                    runs.push(new TextRun({
                        text: w,
                        font: "Times New Roman",
                        size: 28
                    }));
                    runs.push(new FootnoteReferenceRun(footnoteIndex));
                    footnoteIndex++;
                    runs.push(new TextRun(" ")); // فاصله بعد از انگلیسی
                } else {
                    runs.push(new TextRun(w + " "));
                }
            }

            paragraphs.push(new Paragraph({
                alignment,
                spacing: { line: 276 }, // تقریباً 1.15
                children: runs
            }));
        }

        const doc = new Document({
            sections: [{
                children: paragraphs,
                footnotes: new Footnotes({
                    children: createFootnotes(englishRefs)
                })
            }]
        });

        const buffer = await Packer.toBuffer(doc);

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.setHeader("Content-Disposition", "attachment; filename=output.docx");
        res.send(buffer);

    } catch (error) {
        console.error(error);
        res.status(500).send({ error: "خطا در ساخت فایل DOCX" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));