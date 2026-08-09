const fs = require('fs');
const { create } = require('xmlbuilder2');
const JSZip = require('jszip');

async function createDocx(data, outputPath) {
    const zip = new JSZip();

    const contentTypesXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' })
            .ele('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }).up()
            .ele('Default', { Extension: 'xml', ContentType: 'application/xml' }).up()
            .ele('Override', { PartName: '/word/document.xml', ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }).up()
            .ele('Override', { PartName: '/word/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' }).up()
        .end({ prettyPrint: true });

    zip.file('[Content_Types].xml', Buffer.from(contentTypesXml, 'utf-8'));

    const relsXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
            .ele('Relationship', { Id: 'rId1', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target: 'word/document.xml' }).up()
        .end({ prettyPrint: true });

    zip.file('_rels/.rels', Buffer.from(relsXml, 'utf-8'));

    const stylesXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('w:styles', { 'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' })
            .ele('w:docDefaults')
                .ele('w:rPrDefault')
                    .ele('w:rPr').up().up()
                .ele('w:pPrDefault')
                    .ele('w:pPr').up().up().up()
            .ele('w:style', { 'w:type': 'paragraph', 'w:styleId': 'Heading1' })
                .ele('w:name', { 'w:val': 'Heading 1' }).up()
                .ele('w:basedOn', { 'w:val': 'Normal' }).up()
                .ele('w:next', { 'w:val': 'Normal' }).up()
                .ele('w:pPr')
                    .ele('w:spacing', { 'w:after': '200', 'w:before': '0' }).up().up()
                .ele('w:rPr')
                    .ele('w:b', { 'w:val': 'true' }).up()
                    .ele('w:sz', { 'w:val': '28' }).up().up().up()
            .ele('w:style', { 'w:type': 'paragraph', 'w:styleId': 'Heading2' })
                .ele('w:name', { 'w:val': 'Heading 2' }).up()
                .ele('w:basedOn', { 'w:val': 'Normal' }).up()
                .ele('w:next', { 'w:val': 'Normal' }).up()
                .ele('w:pPr')
                    .ele('w:spacing', { 'w:after': '150', 'w:before': '0' }).up().up()
                .ele('w:rPr')
                    .ele('w:b', { 'w:val': 'true' }).up()
                    .ele('w:sz', { 'w:val': '24' }).up().up().up()
        .end({ prettyPrint: true });

    zip.file('word/styles.xml', Buffer.from(stylesXml, 'utf-8'));

    const documentXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('w:document', { 'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' })
            .ele('w:body');

    documentXml.ele('w:p')
        .ele('w:r')
        .ele('w:t', { 'xml:space': 'preserve' }).txt('AI 知识产权法律分析报告').up()
        .up()
        .up();

    documentXml.ele('w:p').up();

    const coverData = [
        ['报告编号', data.report_number || ''],
        ['报告日期', data.report_date || ''],
        ['机密等级', data.confidentiality_level || ''],
        ['委托方', data.client || ''],
        ['编制人', data.author || ''],
        ['适用范围', data.scope || '']
    ];

    const coverTable = documentXml.ele('w:tbl');
    coverTable.ele('w:tblPr').ele('w:tblW', { 'w:w': '9000', 'w:type': 'dxa' }).up().up();
    coverTable.ele('w:tblBorders')
        .ele('w:top', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:left', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:bottom', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:right', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideH', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideV', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .up();

    coverData.forEach(row => {
        const tr = coverTable.ele('w:tr');
        row.forEach(cellText => {
            const tc = tr.ele('w:tc');
            tc.ele('w:p')
                .ele('w:r')
                .ele('w:t', { 'xml:space': 'preserve' }).txt(cellText).up()
                .up()
                .up();
        });
    });

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '目录');

    const tocItems = [
        '1. 摘要', '2. 背景事实', '3. 核心争议焦点', '4. 法律分析与适用',
        '5. 法律依据清单', '6. 风险评估矩阵', '7. 合规建议', '8. 结论', '9. 免责声明'
    ];
    tocItems.forEach(item => {
        addParagraph(documentXml, item);
    });

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '1. 摘要（Executive Summary）');

    if (data.summary) {
        addParagraph(documentXml, '结论摘要：' + (data.summary.conclusion || ''));
        if (data.summary.key_recommendations && data.summary.key_recommendations.length > 0) {
            addParagraph(documentXml, '关键建议：');
            data.summary.key_recommendations.forEach((rec, idx) => {
                addParagraph(documentXml, `${idx + 1}. ${rec}`);
            });
        }
    }

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '2. 背景事实（Facts）');

    if (data.background) {
        addParagraph(documentXml, '业务场景描述：' + (data.background.description || ''));
        if (data.background.timeline && data.background.timeline.length > 0) {
            addParagraph(documentXml, '时间线：');
            data.background.timeline.forEach(item => {
                addParagraph(documentXml, `- ${item}`);
            });
        }
        if (data.background.parties && data.background.parties.length > 0) {
            addParagraph(documentXml, '相关方：');
            data.background.parties.forEach(party => {
                addParagraph(documentXml, `- ${party}`);
            });
        }
    }

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '3. 核心争议焦点（Issues）');

    if (data.issues && data.issues.length > 0) {
        createTable(documentXml, ['序号', '争议焦点', '涉及法律领域'], data.issues.map((issue, idx) => [
            (idx + 1).toString(),
            issue.description || '',
            issue.legal_field || ''
        ]));
    } else {
        addParagraph(documentXml, '无');
    }

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '4. 法律分析与适用（Legal Analysis）');

    if (data.legal_analysis && data.legal_analysis.length > 0) {
        data.legal_analysis.forEach((analysis, idx) => {
            addHeading2(documentXml, `4.${idx + 1} ${analysis.title || ''}`);
            addParagraph(documentXml, '法理探讨：' + (analysis.legal_principle || ''));
            addParagraph(documentXml, '相关判例/政策参考：' + (analysis.cases_policy || ''));
            addParagraph(documentXml, '适用分析：' + (analysis.application || ''));
        });
    } else {
        addParagraph(documentXml, '无');
    }

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '5. 法律依据清单（Legal Basis）');

    if (data.legal_basis && data.legal_basis.length > 0) {
        createTable(documentXml, ['序号', '法律依据名称', '条款内容', '适用情形'], data.legal_basis.map((basis, idx) => [
            (idx + 1).toString(),
            basis.name || '',
            basis.clause || '',
            basis.application || ''
        ]));
    } else {
        addParagraph(documentXml, '无');
    }

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '6. 风险评估矩阵（Risk Assessment Matrix）');

    if (data.risks && data.risks.length > 0) {
        createTable(documentXml, ['序号', '风险点', '风险描述', '发生概率', '影响程度', '风险等级', '关联方'], data.risks.map((risk, idx) => [
            (idx + 1).toString(),
            risk.name || '',
            risk.description || '',
            risk.probability || '',
            risk.impact || '',
            risk.level || '',
            risk.related_party || ''
        ]));
        addParagraph(documentXml, '风险等级判定标准：高风险（发生概率高且影响程度高，或任一维度极高）；中风险（发生概率与影响程度中等，或一高一低）；低风险（发生概率低且影响程度低）');
    } else {
        addParagraph(documentXml, '无');
    }

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '7. 合规建议（Compliance Recommendations）');

    if (data.recommendations && data.recommendations.length > 0) {
        createTable(documentXml, ['优先级', '建议内容', '具体措施', '负责部门', '完成时限'], data.recommendations.map(rec => [
            rec.priority || '',
            rec.title || '',
            rec.measures || '',
            rec.responsible || '',
            rec.deadline || ''
        ]));
    } else {
        addParagraph(documentXml, '无');
    }

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '8. 结论（Conclusion）');

    addParagraph(documentXml, data.conclusion || '');

    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    addHeading(documentXml, '9. 免责声明（Disclaimer）');

    const disclaimerText = data.disclaimer || '本报告仅供内部法律分析参考使用，不构成正式的律师法律意见。AI 领域的法律法规与司法实践处于快速发展期，本报告基于当前公开信息编制，相关结论可能因法律更新或事实变化而调整。建议在具体决策前咨询专业知识产权律师。';
    addParagraph(documentXml, disclaimerText);

    documentXml.ele('w:sectPr').up().up().up();

    const docXml = documentXml.end({ prettyPrint: true });
    zip.file('word/document.xml', Buffer.from(docXml, 'utf-8'));

    const docRelsXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
            .ele('Relationship', { Id: 'rId2', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target: 'styles.xml' }).up()
        .end({ prettyPrint: true });

    zip.file('word/_rels/document.xml.rels', Buffer.from(docRelsXml, 'utf-8'));

    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(outputPath, buffer);
    console.log(`报告已成功生成：${outputPath}`);
}

function addHeading(parent, text) {
    const p = parent.ele('w:p');
    p.ele('w:pPr').ele('w:pStyle', { 'w:val': 'Heading1' }).up().up();
    p.ele('w:r').ele('w:t', { 'xml:space': 'preserve' }).txt(text).up().up().up();
}

function addHeading2(parent, text) {
    const p = parent.ele('w:p');
    p.ele('w:pPr').ele('w:pStyle', { 'w:val': 'Heading2' }).up().up();
    p.ele('w:r').ele('w:t', { 'xml:space': 'preserve' }).txt(text).up().up().up();
}

function addParagraph(parent, text) {
    parent.ele('w:p').ele('w:r').ele('w:t', { 'xml:space': 'preserve' }).txt(text).up().up().up();
}

function createTable(parent, headers, rows) {
    const tbl = parent.ele('w:tbl');
    tbl.ele('w:tblPr').ele('w:tblW', { 'w:w': '9000', 'w:type': 'dxa' }).up().up();
    tbl.ele('w:tblBorders')
        .ele('w:top', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:left', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:bottom', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:right', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideH', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideV', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .up();

    const headerRow = tbl.ele('w:tr');
    headers.forEach(header => {
        const tc = headerRow.ele('w:tc');
        tc.ele('w:p').ele('w:r').ele('w:t', { 'xml:space': 'preserve' }).txt(header).up().up().up();
    });

    rows.forEach(row => {
        const tr = tbl.ele('w:tr');
        row.forEach(cell => {
            const tc = tr.ele('w:tc');
            tc.ele('w:p').ele('w:r').ele('w:t', { 'xml:space': 'preserve' }).txt(cell).up().up().up();
        });
    });
}

if (process.argv.length !== 4) {
    console.log('用法: node generate_report.js <input_json> <output_docx>');
    process.exit(1);
}

const inputJson = process.argv[2];
const outputDocx = process.argv[3];

try {
    const data = JSON.parse(fs.readFileSync(inputJson, 'utf-8'));
    createDocx(data, outputDocx);
} catch (err) {
    console.error(`读取输入文件时出错：${err}`);
    process.exit(1);
}