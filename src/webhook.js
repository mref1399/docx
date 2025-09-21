import express from "express";
import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } from "docx";

const app = express();
app.use(express.json());

app.post("/webhook", async (req, res) => {
    try {
        const { title, content } = req.body;

        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: [
                        new Paragraph({
                            alignment: AlignmentType.RIGHT,
                            heading: HeadingLevel.HEADING_1,
                            children: [
                                new TextRun({
                                    text: title || "عنوان پیش‌فرض",
                                    bold: true,
                                    size: 28,
                                    font: "B Nazanin",
                                }),
                            ],
                        }),
                        new Paragraph({
                            alignment: AlignmentType.JUSTIFIED,
                            spacing: { line: 240 },
                            rightToLeft: true,
                            children: [
                                new TextRun({
                                    text: content || "",
                                    size: 28,
                                    font: "B Nazanin",
                                }),
                            ],
                        }),
                    ],
                },
            ],
        });

        const buffer = await Packer.toBuffer(doc);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.setHeader("Content-Disposition", 'attachment; filename="result.docx"');
        res.send(buffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: "خطا در تولید فایل Word" });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
