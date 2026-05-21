# C-NCAP 2027 AEB 测试场景 — 项目进度

## 项目概述

基于 C-NCAP 2027 管理规则，对 AEB（自动紧急制动）、FCW（前向碰撞预警）、AES（自动紧急转向）等主动安全功能进行测试场景的整理与可视化。

**仓库**: https://github.com/xueli12345/claude_test_260518  
**分支**: `main`  
**最后更新**: 2026-05-20

---

## 当前进度

### 已完成

| 事项 | 说明 |
|------|------|
| PPT 演示文稿 | `AEB_Introduction.pptx` — C-NCAP 2027 AEB 场景总览，含附录P(VRU)和附录Q(C2C)的两栏概览幻灯片 |
| HTML 交互页面 | `CNCAP2027_AEB_Scenarios.html` — 完整的测试场景可视化网页，支持按附录P/Q分Tab浏览、场景卡片展示 |
| HTML 自包含版 | `CNCAP2027_AEB_Scenarios_standalone.html` — 所有50张图片已内嵌为 Base64，单文件拷贝到任何电脑都能正常显示（3.4 MB） |
| 场景图片提取 | `scenario_images/` — 从PDF中提取了 47 张场景示意图（P:/Q: 前缀区分附录） |
| PPTX 幻灯片脚本 | `add_cncap_slide.mjs` — 通过直接操作 XML 向 PPTX 插入新幻灯片的 Node.js 脚本 |
| MinerU 文档提取 Skill | `.claude/skills/mineru-document-extractor/SKILL.md` — 已安装，CLI `mineru-open-api` v0.5.9 可用 |
| 环境配置 | `.claude/settings.local.json` — 权限配置完成（npm/node/python/git/bash等） |

### HTML 页面内容结构

- **附录P Tab** — VRU 弱势道路使用者保护
  - 车对行人 (6项基础 + 6项拓展)
  - 车对二轮车 (6项基础 + 10项拓展)
  - 车对三轮车 (3项基础)
  - 功能覆盖: AEB + FCW + AES
- **附录Q Tab** — C2C 车对车 ADAS 测试
  - 基础场景 (6项必测)
  - 拓展场景 (10项抽测)
  - 功能覆盖: AEB + FCW + AES + LKA + ELK + DMS + BSD + DOW + RCTB + ISLS + AMAP

### 关键数据

- VRU 测试场景总数: **31** 项
- C2C 测试场景: **6 (必测) + 10 (抽测)** 项
- 测试速度范围: **0–120 km/h**
- C-NCAP 版本: **2027**

---

## 源文件清单

| 文件 | 用途 |
|------|------|
| `CNCAP2027/C-NCAP管理规则（2027年版）.pdf` | 主规范文档 |
| `CNCAP2027/附录P  弱势交通参与者保护(主动安全)测试评价规程.pdf` | VRU保护测试规程 |
| `CNCAP2027/附录Q  主动安全先进驾驶辅助系统测试评价规程.pdf` | ADAS测试规程 |
| `AEB_Introduction.pptx` | PPT演示文稿（当前版本） |
| `CNCAP2027_AEB_Scenarios.html` | HTML可视化页面（主版本） |
| `add_cncap_slide.mjs` | PPTX幻灯片注入脚本 |
| `配色.pptx` | 设计配色参考 |

### 生成物（时间戳备份）

- `CNCAP2027_AEB_Scenarios_20260519_144745.html`
- `CNCAP2027_AEB_Scenarios_20260519_163247.html`
- `CNCAP2027_AEB_Scenarios_20260519_171236.html`
- `CNCAP2027_AEB_Scenarios_20260520_111419.html`
- `CNCAP2027_AEB_Scenarios_20260520_113408.html`
- `CNCAP2027_AEB_Scenarios_20260520_132046.html`
- `CNCAP2027_AEB_Scenarios_backup_20260520.html`

---

## 依赖

```json
{
  "@napi-rs/canvas": "^1.0.0",  // PDF渲染/场景图片生成
  "jimp": "^1.6.1",              // 图像处理
  "pdfjs-dist": "^4.0.379"       // PDF解析
}
```

全局 CLI:
- `mineru-open-api` v0.5.9 — 文档提取（PDF/Word/PPT → Markdown）

---

## Git 历史

```
f66e436 Add C-NCAP 2027 AEB test scenarios HTML page
b272a75 Initial commit: AEB introduction PPTX with C-NCAP 2027 test scenarios
```

---

## 下一步 / 待办

- [ ] 补充 PPT 中每个场景的详细说明页
- [ ] HTML 页面增加场景通过标准/评分规则展示
- [ ] 考虑将 HTML 部署为 GitHub Pages
- [ ] 附录P/Q PDF 完整内容提取（可用 MinerU）
