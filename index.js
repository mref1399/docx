import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } from "docx";
import fs from "fs";

// متن نمونه
const content = `
این یک پاراگراف نمونه است که شامل حروف یونانی α، β، γ و نمادهای ریاضی ∑، √، π می‌باشد.
این متن برای تست راست‌چین بودن، فونت B Nazanin، اندازه 14 و فاصله خطوط 1 ایجاد شده است.
`;

const doc = new Document({
    sections: [
        {
            properties: {},
            children: [
                // عنوان
                new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    heading: HeadingLevel.HEADING_1,
                    children: [
                        new TextRun({
                            text: "عنوان اصلی",
                            bold: true,
                            font: "B Nazanin",
                            size: 28, // فونت در docx بر حسب half-points است (14 * 2 = 28)
                        }),
                    ],
                }),

                // پاراگراف متن
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED, // تراز دو طرفه
                    spacing: { line: 240 }, // فاصله خطوط 1 (240 یعنی 1 خط)
                    rightToLeft: true, // راست‌چین
                    children: [
                        new TextRun({
                            text: content,
                            font: "B Nazanin",
                            size: 28, // 14 point
                        }),
                    ],
                }),
            ],
        },
    ],
});

// خروجی فایل ورد
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("test.docx", buffer);
    console.log("فایل test.docx با موفقیت ساخته شد");
});
