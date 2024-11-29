import { Document, Paragraph, Table, TableRow, TableCell, WidthType, AlignmentType, Packer } from 'docx';
import * as fs from 'fs';
import * as path from 'path';

// Create document
const doc = new Document({
    sections: [{
        properties: {},
        children: [
            new Paragraph({
                text: "示例文档",
                heading: 'Heading1'
            }),
            new Paragraph({
                text: "这是一个示例文档，用于测试表格提取功能。",
            }),
            new Paragraph({
                text: "",
            }),
            new Paragraph({
                text: "系统配置",
                heading: 'Heading2'
            }),
            new Table({
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE,
                },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                columnSpan: 2,
                                children: [new Paragraph({
                                    text: "基本配置信息",
                                    alignment: AlignmentType.CENTER
                                })],
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph("服务器地址")],
                            }),
                            new TableCell({
                                children: [new Paragraph("http://localhost:8080")],
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph("数据库名称")],
                            }),
                            new TableCell({
                                children: [new Paragraph("test_db")],
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph("用户名")],
                            }),
                            new TableCell({
                                children: [new Paragraph("admin")],
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph("端口号")],
                            }),
                            new TableCell({
                                children: [new Paragraph("5432")],
                            }),
                        ],
                    }),
                ],
            }),
        ],
    }],
});

// Save document
const outputPath = path.join(__dirname, '..', 'sample.docx');
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(outputPath, buffer);
    console.log(`Sample document created at: ${outputPath}`);
});
