# AI IP Legal Advisor（AI 知识产权法律分析助手）

面向法务专业人员的 AI 知识产权侵权风险评估与保护咨询 Skill，支持**即时法律咨询**和**一键生成专业 Word 法律分析报告**。

## 功能特性

- **侵权风险评估**：分析 AI 模型训练、内容生成、工具引入等场景的潜在 IP 侵权风险
- **可版权性分析**：依据中国、美国、欧盟等法域的最新法规与判例，评估 AIGC 产出的可版权性
- **合规应对策略**：为企业接入 AI 能力、发布 AI 产品提供可执行的法律建议
- **判例与法规引用**：准确引用标志性判例（北京互联网法院"AI文生图"案、纽约时报诉 OpenAI 案等）
- **双模式输出**：即时咨询模式（四段式回答）+ 正式报告模式（9 章完整法律意见书）

## 环境要求

| 模式 | 依赖 | 说明 |
|------|------|------|
| **咨询模式** | 无外部依赖 | 纯文本四段式回答，任意环境可用 |
| **报告生成模式** | Node.js >= 20 | 生成 .docx Word 报告，需先安装依赖（xmlbuilder2 4.x 要求 Node.js >= 20） |

## 安装与快速开始

```bash
# 1. 克隆仓库
git clone https://github.com/tianjinggo/ai-ip-legal-advisor.git
cd ai-ip-legal-advisor

# 2. 安装依赖（仅报告生成模式需要）
npm install

# 3. 生成法律分析报告
node generate_report.js <input_json> <output_docx>

# 示例
node generate_report.js demo_report.json demo_report.docx

# 4. 运行自检（校验中文字体、页眉页脚与依赖提示）
npm test
```

> **注意**：如果只使用咨询模式（纯文本问答），无需执行 `npm install`，直接在 Agent 中加载本 Skill 即可。

## JSON 数据结构

按 SKILL.md 中定义的报告模板组织 JSON 数据，包含以下 9 个章节：

1. 摘要（Executive Summary）
2. 背景事实（Facts）
3. 核心争议焦点（Issues）
4. 法律分析与适用（Legal Analysis）
5. 法律依据清单（Legal Basis）
6. 风险评估矩阵（Risk Assessment Matrix）
7. 合规建议（Compliance Recommendations）
8. 结论（Conclusion）
9. 免责声明（Disclaimer）

详细模板参见 [SKILL.md](./SKILL.md)。

## 报告特性

- 中文宋体/黑体排版（正文宋体，标题黑体）
- 封面元信息 + 目录
- 页眉显示报告标题，页脚显示页码（第 X 页 / 共 Y 页）
- 风险矩阵、法律依据清单、合规建议表采用专业网格表格（表头灰底加粗）
- 标准法律免责声明

## 适用场景

| 场景 | 说明 |
|------|------|
| AIGC 产品合规评估 | 评估 AI 生成内容的上线合规风险 |
| 大模型训练数据审查 | 分析训练数据来源的 IP 侵权风险 |
| 开源模型引入评估 | 审查开源许可协议的合规性与传染风险 |
| AI 侵权维权策略 | 提供侵权证据固定与维权路径建议 |

## 技术栈

- Node.js >= 20
- [xmlbuilder2](https://github.com/oozcitak/xmlbuilder2) — XML/OpenXML 构建
- [JSZip](https://stuk.github.io/jszip/) — 生成 .docx 压缩包

## 免责声明

本工具仅供法律分析参考，不构成正式律师法律意见。AI 领域法律法规与司法实践处于快速发展期，建议在具体决策前咨询专业知识产权律师。
