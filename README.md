# 成文 chengwen

> 公文格式报告撰写与 docx / Markdown 双轨发布 skill（面向 ZCode / Claude 类编码 Agent）

**一句话**：让 Agent 产出的每一份报告都符合国内公文格式（GB/T 9704-2012 口径），并且 docx 与 Markdown **同源定稿、双轨可用**——md 发布到 Git 仓库供网页阅读，docx 按公文版式投递。

## 它解决什么问题

1. **报告版式失控**——Agent 自由排版的报告（微软雅黑、彩色标题、随意层级）不符合国内公文阅读习惯。本 skill 固化 GB/T 9704-2012 口径：3 号仿宋正文、黑体/楷体层级标题、2 号小标宋文种标题、28 磅固定行距、公文页边距、落款成文日期。
2. **docx 兼容事故**——docx@8.x 生成库把表格百分比宽度写成 `w:w="100%"` 字符串：ECMA 新规范允许，但 **WPS 与部分 Word 版本直接拒绝打开**（本机 XML 校验全绿也发现不了）。本 skill 强制 docx@9+ 并附配套校验器，生成后 exit 0 才允许投递。
3. **双处维护漂移**——md 与 docx 各写各的最后必然口径不一。本 skill 规定 **md 先行定稿为正典，docx 由其渲染**。

## 四步流程

```
① 内容定稿（md 正典）
② 公文格式 docx（validate_docx.py 通过）
③ md 镜像 commit / push 到项目仓库（Gitea / GitHub 网页可查）
④ 邮件投递 + 台账落账（logbook）
```

## 安装

```bash
git clone https://github.com/YKaiXu/chengwen.git ~/.agents/skills/chengwen
```

ZCode 自动发现 `~/.agents/skills/` 下的技能；其他 Agent 工具放入对应 skills 目录即可。触发词：形成报告 / 写报告 / 出报告 / 审查报告 / 发我邮箱，或显式调用 `/chengwen <题目>`。

## 配套校验器

```bash
python3 -I scripts/validate_docx.py your-report.docx
```

纯标准库，检查五类"本机能过、Word/WPS 拒开"的坑：

| 检查项 | 说明 |
|---|---|
| zip 完整性 + XML 良构 | 全部件逐个解析 |
| 宽度整数口径 | `tblW`/`tcW` 出现 `"nn%"` 字符串即 FAIL（docx@8.x 指纹） |
| 空 `<w:pgNumType/>` | WPS 兼容清单要求剥离（WARN） |
| 页码域白名单 | 页眉页脚 instrText 仅允许 PAGE/PAGEREF/TOC 等标准域，防自定义补丁残留 |
| rels 悬挂引用 | 引用不存在的部件（Word 会弹"修复"对话框） |

## 公文格式口径速览

| 元素 | 字体 | 字号 |
|---|---|---|
| 文种标题 | 方正小标宋（回退华文中宋） | 2 号 22pt，居中 |
| 正文 | 仿宋_GB2312 | 3 号 16pt，行距固定 28 磅，首行缩进 2 字符 |
| 一级标题 | 黑体 | 3 号，"一、" |
| 二级标题 | 楷体_GB2312 | 3 号，"（一）" |
| 三/四级 | 仿宋 | 3 号，"1." / "（1）" |

页边距 上 37 / 下 35 / 左 28 / 右 26 mm（公文版心 156×225mm）；层次序号一、→（一）→1.→（1）；成文日期用阿拉伯数字、右空四字；红头仅显式要求时加。完整口径见 [SKILL.md](SKILL.md)。

## 依赖

- Node.js / bun + **docx@9 及以上**（8.x 有宽度字符串缺陷，禁用）
- python3（校验器，纯标准库）
- 可选：documents 类 docx 排版 skill（版式引擎）、`send-mail` 命令（邮件）、Git 远端（md 镜像）

## License

MIT
