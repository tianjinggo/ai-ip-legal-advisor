const fs = require('fs');

function requireDependency(name) {
    try {
        return require(name);
    } catch (error) {
        if (error.code === 'MODULE_NOT_FOUND') {
            throw new Error(`缺少报告生成依赖 ${name}，请先在 skill 目录执行 npm install。`);
        }
        throw error;
    }
}

async function createDocx(data, outputPath) {
    const { create } = requireDependency('xmlbuilder2');
    const JSZip = requireDependency('jszip');
    const zip = new JSZip();
    const reportTitle = 'AI 知识产权法律分析报告';

    // ===== [Content_Types].xml =====
    const contentTypesXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' })
            .ele('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }).up()
            .ele('Default', { Extension: 'xml', ContentType: 'application/xml' }).up()
            .ele('Override', { PartName: '/word/document.xml', ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }).up()
            .ele('Override', { PartName: '/word/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' }).up()
            .ele('Override', { PartName: '/word/header1.xml', ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml' }).up()
            .ele('Override', { PartName: '/word/footer1.xml', ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' }).up()
        .end({ prettyPrint: true });

    zip.file('[Content_Types].xml', Buffer.from(contentTypesXml, 'utf-8'));

    // ===== _rels/.rels =====
    const relsXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
            .ele('Relationship', { Id: 'rId1', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', Target: 'word/document.xml' }).up()
        .end({ prettyPrint: true });

    zip.file('_rels/.rels', Buffer.from(relsXml, 'utf-8'));

    // ===== word/styles.xml（含中文字体设置） =====
    const stylesXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('w:styles', { 'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' })
            .ele('w:docDefaults')
                .ele('w:rPrDefault')
                    .ele('w:rPr')
                        .ele('w:rFonts', {
                            'w:ascii': 'Times New Roman',
                            'w:hAnsi': 'Times New Roman',
                            'w:eastAsia': '宋体',
                            'w:cs': 'Times New Roman'
                        }).up()
                        .ele('w:sz', { 'w:val': '21' }).up()
                        .ele('w:szCs', { 'w:val': '21' }).up()
                    .up()
                .up()
                .ele('w:pPrDefault')
                    .ele('w:pPr').up()
                .up()
            .up()
            // Normal 样式
            .ele('w:style', { 'w:type': 'paragraph', 'w:default': '1', 'w:styleId': 'Normal' })
                .ele('w:name', { 'w:val': 'Normal' }).up()
                .ele('w:qFormat').up()
                .ele('w:pPr')
                    .ele('w:spacing', { 'w:after': '120', 'w:line': '360', 'w:lineRule': 'auto' }).up()
                .up()
                .ele('w:rPr')
                    .ele('w:rFonts', {
                        'w:ascii': 'Times New Roman',
                        'w:hAnsi': 'Times New Roman',
                        'w:eastAsia': '宋体'
                    }).up()
                    .ele('w:sz', { 'w:val': '21' }).up()
                .up()
            .up()
            // Heading 1 样式（黑体）
            .ele('w:style', { 'w:type': 'paragraph', 'w:styleId': 'Heading1' })
                .ele('w:name', { 'w:val': 'Heading 1' }).up()
                .ele('w:basedOn', { 'w:val': 'Normal' }).up()
                .ele('w:next', { 'w:val': 'Normal' }).up()
                .ele('w:qFormat').up()
                .ele('w:pPr')
                    .ele('w:keepNext').up()
                    .ele('w:spacing', { 'w:before': '360', 'w:after': '200' }).up()
                .up()
                .ele('w:rPr')
                    .ele('w:rFonts', {
                        'w:ascii': 'Arial',
                        'w:hAnsi': 'Arial',
                        'w:eastAsia': '黑体'
                    }).up()
                    .ele('w:b', { 'w:val': 'true' }).up()
                    .ele('w:sz', { 'w:val': '32' }).up()
                    .ele('w:szCs', { 'w:val': '32' }).up()
                .up()
            .up()
            // Heading 2 样式（黑体）
            .ele('w:style', { 'w:type': 'paragraph', 'w:styleId': 'Heading2' })
                .ele('w:name', { 'w:val': 'Heading 2' }).up()
                .ele('w:basedOn', { 'w:val': 'Normal' }).up()
                .ele('w:next', { 'w:val': 'Normal' }).up()
                .ele('w:qFormat').up()
                .ele('w:pPr')
                    .ele('w:keepNext').up()
                    .ele('w:spacing', { 'w:before': '240', 'w:after': '150' }).up()
                .up()
                .ele('w:rPr')
                    .ele('w:rFonts', {
                        'w:ascii': 'Arial',
                        'w:hAnsi': 'Arial',
                        'w:eastAsia': '黑体'
                    }).up()
                    .ele('w:b', { 'w:val': 'true' }).up()
                    .ele('w:sz', { 'w:val': '28' }).up()
                    .ele('w:szCs', { 'w:val': '28' }).up()
                .up()
            .up()
        .end({ prettyPrint: true });

    zip.file('word/styles.xml', Buffer.from(stylesXml, 'utf-8'));

    // ===== word/header1.xml（页眉） =====
    const headerXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('w:hdr', {
            'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w:type': 'default'
        })
            .ele('w:p')
                .ele('w:pPr')
                    .ele('w:pStyle', { 'w:val': 'Header' }).up()
                    .ele('w:jc', { 'w:val': 'right' }).up()
                    .ele('w:pBdr')
                        .ele('w:bottom', {
                            'w:val': 'single', 'w:sz': '6', 'w:space': '1', 'w:color': '999999'
                        }).up()
                    .up()
                .up()
                .ele('w:r')
                    .ele('w:rPr')
                        .ele('w:rFonts', { 'w:eastAsia': '宋体' }).up()
                        .ele('w:sz', { 'w:val': '18' }).up()
                        .ele('w:color', { 'w:val': '666666' }).up()
                    .up()
                    .ele('w:t', { 'xml:space': 'preserve' }).txt(reportTitle).up()
                .up()
            .up()
        .end({ prettyPrint: true });

    zip.file('word/header1.xml', Buffer.from(headerXml, 'utf-8'));

    // ===== word/footer1.xml（页脚，含页码） =====
    const footerXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('w:ftr', {
            'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w:type': 'default'
        })
            .ele('w:p')
                .ele('w:pPr')
                    .ele('w:pStyle', { 'w:val': 'Footer' }).up()
                    .ele('w:jc', { 'w:val': 'center' }).up()
                .up()
                .ele('w:r')
                    .ele('w:rPr')
                        .ele('w:rFonts', { 'w:eastAsia': '宋体' }).up()
                        .ele('w:sz', { 'w:val': '18' }).up()
                        .ele('w:color', { 'w:val': '666666' }).up()
                    .up()
                    .ele('w:t', { 'xml:space': 'preserve' }).txt('第 ').up()
                .up()
                // PAGE 字段
                .ele('w:r')
                    .ele('w:fldChar', { 'w:fldCharType': 'begin' }).up()
                .up()
                .ele('w:r')
                    .ele('w:instrText', { 'xml:space': 'preserve' }).txt(' PAGE ').up()
                .up()
                .ele('w:r')
                    .ele('w:fldChar', { 'w:fldCharType': 'end' }).up()
                .up()
                .ele('w:r')
                    .ele('w:rPr')
                        .ele('w:rFonts', { 'w:eastAsia': '宋体' }).up()
                        .ele('w:sz', { 'w:val': '18' }).up()
                        .ele('w:color', { 'w:val': '666666' }).up()
                    .up()
                    .ele('w:t', { 'xml:space': 'preserve' }).txt(' 页 / 共 ').up()
                .up()
                // NUMPAGES 字段
                .ele('w:r')
                    .ele('w:fldChar', { 'w:fldCharType': 'begin' }).up()
                .up()
                .ele('w:r')
                    .ele('w:instrText', { 'xml:space': 'preserve' }).txt(' NUMPAGES ').up()
                .up()
                .ele('w:r')
                    .ele('w:fldChar', { 'w:fldCharType': 'end' }).up()
                .up()
                .ele('w:r')
                    .ele('w:rPr')
                        .ele('w:rFonts', { 'w:eastAsia': '宋体' }).up()
                        .ele('w:sz', { 'w:val': '18' }).up()
                        .ele('w:color', { 'w:val': '666666' }).up()
                    .up()
                    .ele('w:t', { 'xml:space': 'preserve' }).txt(' 页').up()
                .up()
            .up()
        .end({ prettyPrint: true });

    zip.file('word/footer1.xml', Buffer.from(footerXml, 'utf-8'));

    // ===== word/document.xml =====
    const documentXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('w:document', {
            'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        })
            .ele('w:body');

    // 报告标题
    const titleP = documentXml.ele('w:p');
    titleP.ele('w:pPr')
        .ele('w:jc', { 'w:val': 'center' }).up()
        .up();
    titleP.ele('w:r')
        .ele('w:rPr')
            .ele('w:rFonts', { 'w:eastAsia': '黑体' }).up()
            .ele('w:b', { 'w:val': 'true' }).up()
            .ele('w:sz', { 'w:val': '36' }).up()
        .up()
        .ele('w:t', { 'xml:space': 'preserve' }).txt(reportTitle).up()
        .up();

    documentXml.ele('w:p').up();

    // 封面信息表
    const coverData = [
        ['报告编号', data.report_number || ''],
        ['报告日期', data.report_date || ''],
        ['机密等级', data.confidentiality_level || ''],
        ['委托方', data.client || ''],
        ['编制人', data.author || ''],
        ['适用范围', data.scope || '']
    ];

    const coverTable = documentXml.ele('w:tbl');
    const coverTableProperties = coverTable.ele('w:tblPr');
    coverTableProperties.ele('w:tblW', { 'w:w': '9000', 'w:type': 'dxa' });
    coverTableProperties.ele('w:tblBorders')
        .ele('w:top', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:left', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:bottom', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:right', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideH', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideV', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' });

    coverData.forEach(row => {
        const tr = coverTable.ele('w:tr');
        row.forEach((cellText, idx) => {
            const tc = tr.ele('w:tc');
            if (idx === 0) {
                tc.ele('w:tcPr').ele('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': 'F2F2F2' }).up().up();
            }
            tc.ele('w:p')
                .ele('w:r')
                .ele('w:t', { 'xml:space': 'preserve' }).txt(cellText).up()
                .up()
                .up();
        });
    });

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 目录
    addHeading(documentXml, '目录');

    const tocItems = [
        '1. 摘要', '2. 背景事实', '3. 核心争议焦点', '4. 法律分析与适用',
        '5. 法律依据清单', '6. 风险评估矩阵', '7. 合规建议', '8. 结论', '9. 免责声明'
    ];
    tocItems.forEach(item => {
        addParagraph(documentXml, item);
    });

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 1. 摘要
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

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 2. 背景事实
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

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 3. 核心争议焦点
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

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 4. 法律分析与适用
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

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 5. 法律依据清单
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

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 6. 风险评估矩阵
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

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 7. 合规建议
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

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 8. 结论
    addHeading(documentXml, '8. 结论（Conclusion）');

    addParagraph(documentXml, data.conclusion || '');

    // 分页
    documentXml.ele('w:p').ele('w:r').ele('w:br', { 'w:type': 'page' }).up().up().up();

    // 9. 免责声明
    addHeading(documentXml, '9. 免责声明（Disclaimer）');

    const disclaimerText = data.disclaimer || '本报告仅供内部法律分析参考使用，不构成正式的律师法律意见。AI 领域的法律法规与司法实践处于快速发展期，本报告基于当前公开信息编制，相关结论可能因法律更新或事实变化而调整。建议在具体决策前咨询专业知识产权律师。';
    addParagraph(documentXml, disclaimerText);

    // sectPr（引用页眉页脚）
    const sectPr = documentXml.ele('w:sectPr');
    sectPr.ele('w:headerReference', { 'w:type': 'default', 'r:id': 'rId3' }).up();
    sectPr.ele('w:footerReference', { 'w:type': 'default', 'r:id': 'rId4' }).up();
    sectPr.ele('w:pgSz', { 'w:w': '11906', 'w:h': '16838' }).up();
    sectPr.ele('w:pgMar', {
        'w:top': '1440', 'w:right': '1440', 'w:bottom': '1440', 'w:left': '1440',
        'w:header': '720', 'w:footer': '720', 'w:gutter': '0'
    }).up();

    const docXml = documentXml.end({ prettyPrint: true });
    zip.file('word/document.xml', Buffer.from(docXml, 'utf-8'));

    // ===== word/_rels/document.xml.rels（含 header/footer 关系） =====
    const docRelsXml = create({ version: '1.0', encoding: 'UTF-8' })
        .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
            .ele('Relationship', { Id: 'rId2', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', Target: 'styles.xml' }).up()
            .ele('Relationship', { Id: 'rId3', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', Target: 'header1.xml' }).up()
            .ele('Relationship', { Id: 'rId4', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', Target: 'footer1.xml' }).up()
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
    const tableProperties = tbl.ele('w:tblPr');
    tableProperties.ele('w:tblW', { 'w:w': '9000', 'w:type': 'dxa' });
    tableProperties.ele('w:tblBorders')
        .ele('w:top', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:left', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:bottom', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:right', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideH', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' }).up()
        .ele('w:insideV', { 'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': '000000' });

    // 表头行（灰色底）
    const headerRow = tbl.ele('w:tr');
    headers.forEach(header => {
        const tc = headerRow.ele('w:tc');
        tc.ele('w:tcPr').ele('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': 'D9E2F3' }).up().up();
        tc.ele('w:p')
            .ele('w:r')
            .ele('w:rPr').ele('w:b', { 'w:val': 'true' }).up().up()
            .ele('w:t', { 'xml:space': 'preserve' }).txt(header).up()
            .up()
            .up();
    });

    rows.forEach(row => {
        const tr = tbl.ele('w:tr');
        row.forEach(cell => {
            const tc = tr.ele('w:tc');
            tc.ele('w:p').ele('w:r').ele('w:t', { 'xml:space': 'preserve' }).txt(cell).up().up().up();
        });
    });
}

async function main() {
    if (process.argv.length !== 4) {
        console.log('用法: node generate_report.js <input_json> <output_docx>');
        process.exitCode = 1;
        return;
    }

    const inputJson = process.argv[2];
    const outputDocx = process.argv[3];

    try {
        const data = JSON.parse(fs.readFileSync(inputJson, 'utf-8'));
        await createDocx(data, outputDocx);
    } catch (error) {
        console.error(`生成报告时出错：${error.message}`);
        process.exitCode = 1;
    }
}

if (require.main === module) {
    main();
}

module.exports = { createDocx };
