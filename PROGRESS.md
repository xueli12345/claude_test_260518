# C-NCAP 2027 AEB 测试场景 — 项目进度

## 项目概述

基于 C-NCAP 2027 管理规则，对 AEB（自动紧急制动）、FCW（前向碰撞预警）、AES（自动紧急转向）等主动安全功能进行测试场景的整理、可视化与 V2X 赋能分析。

**仓库**: https://github.com/xueli12345/claude_test_260518  
**GitHub Pages**: https://xueli12345.github.io/claude_test_260518/  
**分支**: `main`  
**最后更新**: 2026-05-21

---

## 当前进度

### 已完成

| 事项 | 说明 |
|------|------|
| PPT 演示文稿 | `AEB_Introduction.pptx` — 3 页幻灯片（AEB介绍 / 场景总览 / V2X赋能分析） |
| HTML 交互页面 | `CNCAP2027_AEB_Scenarios.html` — 测试场景可视化网页，按附录P/Q分Tab浏览 |
| GitHub Pages 部署 | `index.html` — 自包含版，50张图片Base64内嵌，单链接可分享任何人 |
| V2X 赋能分析 | PPT Slide 3 — 按 V2P/V2V/V2I 分类，覆盖遮挡穿透、恶劣天气、意图共享 |
| 场景图片提取 | `scenario_images/` — 从PDF中提取了 50 张场景示意图（P:/Q: 前缀区分附录） |
| MinerU 文档提取 | `.claude/skills/mineru-document-extractor/SKILL.md` — CLI `mineru-open-api` v0.5.9 |
| 环境配置 | `.claude/settings.local.json` — 权限配置（npm/node/python/git等） |

### PPT 结构

| Slide | 内容 |
|-------|------|
| 1 | AEB 介绍 — 自动紧急制动系统概述 |
| 2 | C-NCAP 2027 AEB 测试场景总览 — 附录P(VRU) 左栏 + 附录Q(C2C) 右栏 |
| 3 | V2X 赋能分析 — 左栏 V2P/VRU 场景，右栏 V2V/V2I C2C 场景 |

### V2X 赋能分析要点

- **附录P VRU — V2P**: 遮挡穿透（CPFC-MO 等）、转弯/横穿预警（CPTA系列、C2S/C2B）、恶劣环境（CPFAr雨天、CBLAb夜间）
- **附录Q C2C — V2V**: 遮挡/非视距（SCPmo/SCPso、CCFT系列）、恶劣天气（CCRsr/CCRsf）、意图共享（CCRc/Rb/Rh、RCPf）
- **V2I**: 全部路口转弯/横穿场景可通过路侧单元(RSU)获取全局目标
- 约 **70%** 场景可由 V2X 增强，核心优势：NLOS 非视距穿透、全天候

### HTML 页面内容

- **附录P Tab** — VRU 弱势道路使用者保护（31项）
  - 车对行人 (6项基础 + 6项拓展)
  - 车对二轮车 (6项基础 + 10项拓展)
  - 车对三轮车 (3项基础)
- **附录Q Tab** — C2C 车对车 ADAS 测试（16项）
  - 基础场景 (6项必测)
  - 拓展场景 (10项抽测)

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
| `CNCAP2027/附录P ...测试评价规程.pdf` | VRU保护测试规程 |
| `CNCAP2027/附录Q ...测试评价规程.pdf` | ADAS测试规程 |
| `AEB_Introduction.pptx` | PPT演示文稿（3页，含V2X分析） |
| `index.html` | GitHub Pages 自包含版（图片Base64内嵌） |
| `CNCAP2027_AEB_Scenarios.html` | HTML可视化页面（主版本） |
| `scenario_images/` | 50张场景示意图 |
| `add_cncap_slide.mjs` | PPTX幻灯片注入脚本（Slide 2） |
| `add_v2x_slide.py` | PPTX V2X幻灯片正确注入（python-pptx/lxml） |
| `embed_images.mjs` | 图片Base64内嵌脚本 |
| `配色.pptx` | 设计配色参考 |

---

## 依赖

```json
{
  "@napi-rs/canvas": "^1.0.0",
  "jimp": "^1.6.1",
  "pdfjs-dist": "^4.0.379",
  "archiver": "latest",
  "unzipper": "latest"
}
```

Python: `python-pptx`, `lxml`

全局 CLI: `mineru-open-api` v0.5.9

---

## Git 历史

```
d84c469 Fix V2X slide: use python-pptx/lxml for proper PPTX XML handling
acb7b46 Add V2X enablement analysis slide to AEB presentation
753b1a8 Add project files: images, scripts, config, and progress doc
1cb642f Rename standalone HTML to index.html for GitHub Pages short URL
21ceaee Add standalone HTML with embedded images for easy sharing
f66e436 Add C-NCAP 2027 AEB test scenarios HTML page
b272a75 Initial commit: AEB introduction PPTX with C-NCAP 2027 test scenarios
```

---

## 下一步 / 待办

- [ ] 补充 PPT 中每个场景的详细说明页
- [ ] HTML 页面增加场景通过标准/评分规则展示
- [ ] 附录P/Q PDF 完整内容提取（可用 MinerU）
- [ ] PPT V2X 页的内容进一步细化（V2X 通信技术标准、时延指标等）
