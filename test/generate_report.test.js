const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');
const JSZip = require('jszip');
const { create } = require('xmlbuilder2');
const { createDocx } = require('../generate_report');

test('生成的 docx 包含中文字体与页眉页脚', async (context) => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'ai-ip-report-'));
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));
    const outputPath = path.join(tempDirectory, 'report.docx');

    await createDocx({
        report_number: 'AI-IP-LR-2026-001',
        report_date: '2026年09月18日',
        client: '测试委托方',
        summary: { conclusion: '测试结论', key_recommendations: ['测试建议'] },
        conclusion: '整体风险可控'
    }, outputPath);

    const archive = await JSZip.loadAsync(fs.readFileSync(outputPath));
    const requiredFiles = [
        '[Content_Types].xml',
        'word/document.xml',
        'word/styles.xml',
        'word/header1.xml',
        'word/footer1.xml',
        'word/_rels/document.xml.rels'
    ];

    for (const fileName of requiredFiles) {
        assert.ok(archive.file(fileName), `缺少 ${fileName}`);
        const xml = await archive.file(fileName).async('string');
        assert.doesNotThrow(() => create(xml), `${fileName} 不是有效 XML`);
    }

    const styles = await archive.file('word/styles.xml').async('string');
    assert.match(styles, /w:eastAsia="宋体"/);
    assert.match(styles, /w:eastAsia="黑体"/);

    const document = await archive.file('word/document.xml').async('string');
    assert.match(document, /xmlns:r="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships"/);
    assert.match(document, /w:headerReference[^>]+r:id="rId3"/);
    assert.match(document, /w:footerReference[^>]+r:id="rId4"/);
    assert.match(document, /<w:tblPr>[\s\S]*?<w:tblBorders>/);
    assert.doesNotMatch(document, /<\/w:tblPr>\s*<w:tblBorders>/);

    const relationships = await archive.file('word/_rels/document.xml.rels').async('string');
    assert.match(relationships, /relationships\/header/);
    assert.match(relationships, /relationships\/footer/);

    const header = await archive.file('word/header1.xml').async('string');
    assert.match(header, /AI 知识产权法律分析报告/);

    const footer = await archive.file('word/footer1.xml').async('string');
    assert.match(footer, /w:instrText[^>]*> PAGE <\/w:instrText>/);
});

test('缺少依赖时提供 npm install 指引', (context) => {
    const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'ai-ip-report-missing-deps-'));
    context.after(() => fs.rmSync(tempDirectory, { recursive: true, force: true }));
    const scriptPath = path.join(tempDirectory, 'generate_report.js');
    const inputPath = path.join(tempDirectory, 'report.json');
    fs.copyFileSync(path.join(__dirname, '..', 'generate_report.js'), scriptPath);
    fs.writeFileSync(inputPath, '{}');

    const result = spawnSync(process.execPath, [scriptPath, inputPath, path.join(tempDirectory, 'report.docx')], {
        encoding: 'utf8'
    });

    assert.equal(result.status, 1);
    assert.match(result.stderr, /缺少报告生成依赖 xmlbuilder2/);
    assert.match(result.stderr, /npm install/);
});
