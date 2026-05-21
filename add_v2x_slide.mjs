import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const { ZipArchive } = require('archiver');
const unzipper = require('unzipper');

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TEMP = path.join(__dirname, 'pptx_temp');
const PPTX = path.join(__dirname, 'AEB_Introduction.pptx');
const OUT  = path.join(__dirname, 'AEB_Introduction_new.pptx');

const SLIDE_ID = "261";
const RID = "rId10";
const NOTES_RID = "rId3";

// Clean temp
fs.rmSync(TEMP, { recursive: true, force: true });
fs.mkdirSync(TEMP, { recursive: true });

// Unzip
await new Promise((resolve, reject) => {
  fs.createReadStream(PPTX)
    .pipe(unzipper.Extract({ path: TEMP }))
    .on('close', resolve)
    .on('error', reject);
});

// ── V2X slide content ──

function v2xLines(side) {
  // side: 'left' = 附录P VRU V2X, 'right' = 附录Q C2C V2X
  if (side === 'left') {
    const cats = [
      { title: 'V2P 穿透遮挡 — 行人/二轮车感知', color: '1EEF97', items: [
        'CPFC-MO 动态遮挡儿童 — 被挡儿童可被V2P提前感知',
        'CPNC-AI / CPFC-BI 干扰横穿 — 被遮挡弱势道路使用者',
        'CBOFA / CBFA-MO 遮挡斜穿 — 穿透车辆障碍',
        'CBNA-AI / CSFA-BI 干扰近端 — V2P弥补盲区',
        'CSTAO-RN 右转遮挡 — 路侧V2I+车端V2P协同',
      ]},
      { title: 'V2P 转弯/横穿预警', color: '1EEF97', items: [
        'CPTA-LF / CPTA-RF 转弯横穿行人 — 提前感知意图',
        'C2S RCPf / C2B RCPn 转弯横穿二轮 — 路口盲区',
        'CBLAt 转弯切入 / CBLAc 变道切入 — V2P/V2V共享意图',
        'CSFhol 直行对向二轮 — 远距离V2P预警',
      ]},
      { title: '恶劣环境优势 — 天气/光照', color: '442EE0', items: [
        'CPFAr 雨天远端横穿 — 摄像头受限，V2X不受影响',
        'CBLAb 电动自行车夜间 — 弱光下V2P优于视觉',
        'CPFAh / CPLAh 高速场景 — 远距离V2P提前预警',
      ]},
    ];
    let lines = '';
    for (const cat of cats) {
      const c = cat.color;
      lines += `<a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
          <a:r><a:rPr lang="zh-CN" sz="900" b="1" dirty="0"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${cat.title}</a:t></a:r></a:p>`;
      for (const item of cat.items) {
        lines += `<a:p><a:pPr marL="171450" indent="0" algn="l"><a:buNone/></a:pPr>
          <a:r><a:rPr lang="zh-CN" sz="700" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${item}</a:t></a:r></a:p>`;
      }
      lines += `<a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr><a:endParaRPr lang="zh-CN" sz="300" dirty="0"/></a:p>`;
    }
    return lines;
  } else {
    const cats = [
      { title: 'V2V 遮挡/非视距 — 穿透障碍', color: '1EEF97', items: [
        'C2C SCPmo 动态遮挡横穿 — 被挡车辆V2V预警',
        'C2C SCPso 静态遮挡横穿 — 越过建筑物/车辆感知',
        'CCFT / CCFTf 转弯对撞 — 路口建筑遮挡，V2V直连',
        'LCPf / LCPn 左转横穿 — 对向遮挡车辆提前通信',
      ]},
      { title: 'V2V 恶劣天气 — 传感器降级', color: '1EEF97', items: [
        'CCRsr 雨天追尾 — 雷达/摄像头衰减，V2V不受影响',
        'CCRsf 雾天追尾 — 能见度极低，V2V稳定通信',
        'CCFhos 直行对撞 — 雨雾天远距离V2V预警',
      ]},
      { title: 'V2V 意图共享 — 切入/制动', color: '442EE0', items: [
        'CCRc 切入追尾 — 邻车变道意图实时共享',
        'CCRb 制动追尾 — 前车制动信号低延迟广播',
        'CCRh 高速追尾 — 超视距提前获知前车减速',
        'CCRbc 弯道制动 — 弯道遮挡场景V2V补盲',
        'RCPf 右转横穿 — 转向意图+位置共享',
      ]},
      { title: 'V2I 路侧协同', color: '442EE0', items: [
        '所有十字路口转弯/横穿场景均可通过V2I路侧单元(RSU)',
        '获取路口全局目标列表，弥补单车感知盲区',
        '典型: CPTA系列 / CCFT系列 / C2C LCP&SCP系列',
      ]},
    ];
    let lines = '';
    for (const cat of cats) {
      const c = cat.color;
      lines += `<a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
          <a:r><a:rPr lang="zh-CN" sz="900" b="1" dirty="0"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${cat.title}</a:t></a:r></a:p>`;
      for (const item of cat.items) {
        lines += `<a:p><a:pPr marL="171450" indent="0" algn="l"><a:buNone/></a:pPr>
          <a:r><a:rPr lang="zh-CN" sz="700" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${item}</a:t></a:r></a:p>`;
      }
      lines += `<a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr><a:endParaRPr lang="zh-CN" sz="300" dirty="0"/></a:p>`;
    }
    return lines;
  }
}

const slide3Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Slide 3 - V2X">
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="1D0638"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <!-- Top line -->
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="TopLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="54864"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="442EE0"/></a:solidFill><a:ln/></p:spPr></p:sp>
      <!-- Left bar -->
      <p:sp><p:nvSpPr><p:cNvPr id="3" name="LeftBar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="73152" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="442EE0"/></a:solidFill><a:ln/></p:spPr></p:sp>
      <!-- Green dot -->
      <p:sp><p:nvSpPr><p:cNvPr id="4" name="GreenDot"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="548640" y="411480"/><a:ext cx="228600" cy="228600"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:ln/></p:spPr></p:sp>
      <!-- CARIAD label -->
      <p:sp><p:nvSpPr><p:cNvPr id="5" name="CARIAD"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="868680" y="384048"/><a:ext cx="2743200" cy="320040"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="1100" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>C A R I A D</a:t></a:r></a:p></p:txBody></p:sp>
      <!-- Title -->
      <p:sp><p:nvSpPr><p:cNvPr id="6" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="700000"/><a:ext cx="10972800" cy="640000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="3200" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>V2X </a:t></a:r>
            <a:r><a:rPr lang="en-US" sz="3200" b="1" dirty="0"><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>赋能</a:t></a:r>
            <a:r><a:rPr lang="en-US" sz="3200" b="1" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t> C-NCAP 2027 AEB 测试场景</a:t></a:r>
          </a:p></p:txBody></p:sp>
      <!-- Underline -->
      <p:sp><p:nvSpPr><p:cNvPr id="7" name="Underline"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="1340000"/><a:ext cx="2743200" cy="45720"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:ln/></p:spPr></p:sp>
      <!-- Subtitle -->
      <p:sp><p:nvSpPr><p:cNvPr id="8" name="Subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="1450000"/><a:ext cx="10972800" cy="274320"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="1400" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>V2V 车车直连 | V2P 车与弱势道路使用者 | V2I 车路协同 | 非视距穿透 · 恶劣天气抗扰 · 意图共享</a:t></a:r>
          </a:p></p:txBody></p:sp>
      <!-- Vertical divider -->
      <p:sp><p:nvSpPr><p:cNvPr id="9" name="VLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6131451" y="1850000"/><a:ext cx="45719" cy="4200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="442EE0"/></a:solidFill><a:ln/></p:spPr></p:sp>
      <!-- Left header -->
      <p:sp><p:nvSpPr><p:cNvPr id="10" name="LeftHeader"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="1820000"/><a:ext cx="5100000" cy="320040"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1100" b="1" dirty="0"><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>附录P — VRU场景 | V2P / 车与弱势道路使用者</a:t></a:r>
          </a:p></p:txBody></p:sp>
      <!-- Left list -->
      <p:sp><p:nvSpPr><p:cNvPr id="11" name="LeftList"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="2160000"/><a:ext cx="5100000" cy="4000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>${v2xLines('left')}</p:txBody></p:sp>
      <!-- Right header -->
      <p:sp><p:nvSpPr><p:cNvPr id="20" name="RightHeader"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6636028" y="1820000"/><a:ext cx="5100000" cy="320040"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1100" b="1" dirty="0"><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>附录Q — C2C场景 | V2V / V2I 车车与车路协同</a:t></a:r>
          </a:p></p:txBody></p:sp>
      <!-- Right list -->
      <p:sp><p:nvSpPr><p:cNvPr id="21" name="RightList"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6636028" y="2160000"/><a:ext cx="5100000" cy="4000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>${v2xLines('right')}</p:txBody></p:sp>
      <!-- Bottom separator -->
      <p:sp><p:nvSpPr><p:cNvPr id="90" name="BtmLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="5900000"/><a:ext cx="11064240" cy="10973"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="373741"/></a:solidFill><a:ln/></p:spPr></p:sp>
      <!-- Stats row -->
      <p:sp><p:nvSpPr><p:cNvPr id="91" name="Stat1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="5970000"/><a:ext cx="2743200" cy="292608"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1600" b="1" dirty="0"><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>~70%</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="92" name="Stat1Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="6215000"/><a:ext cx="2743200" cy="182880"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="800" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>场景可由V2X增强</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="93" name="Stat2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="3292834" y="5970000"/><a:ext cx="2743200" cy="292608"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1600" b="1" dirty="0"><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>V2V+P+I</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="94" name="Stat2Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="3292834" y="6215000"/><a:ext cx="2743200" cy="182880"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="800" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>V2X通信模式协同</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="95" name="Stat3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6357863" y="5970000"/><a:ext cx="2743200" cy="292608"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1600" b="1" dirty="0"><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>NLOS</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="96" name="Stat3Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="6357863" y="6195000"/><a:ext cx="2743200" cy="182880"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="800" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>非视距穿透是核心优势</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="97" name="Stat4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="9123923" y="5970000"/><a:ext cx="2743200" cy="292608"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="en-US" sz="1600" b="1" dirty="0"><a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Arial" pitchFamily="34" charset="-122"/></a:rPr><a:t>全天候</a:t></a:r></a:p></p:txBody></p:sp>
      <p:sp><p:nvSpPr><p:cNvPr id="98" name="Stat4Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="9123923" y="6195000"/><a:ext cx="2743200" cy="182880"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="800" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>雨雾夜不受影响</a:t></a:r></a:p></p:txBody></p:sp>
      <!-- Footer -->
      <p:sp><p:nvSpPr><p:cNvPr id="99" name="Footer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="526774" y="6500000"/><a:ext cx="4572000" cy="137160"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/></p:spPr>
        <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>
          <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r><a:rPr lang="zh-CN" sz="700" dirty="0"><a:solidFill><a:srgbClr val="888888"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>INTERNAL / CONFIDENTIAL  |  CSD class: 2.5  |  基于 C-NCAP 2027 场景的 V2X 赋能分析</a:t></a:r>
          </a:p></p:txBody></p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;

// ── Write slide files ──
const slidesDir = path.join(TEMP, 'ppt', 'slides');
fs.mkdirSync(path.join(slidesDir, '_rels'), { recursive: true });
fs.writeFileSync(path.join(slidesDir, 'slide3.xml'), slide3Xml, 'utf-8');

// slide3.xml.rels
const slide3Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide3.xml"/>
</Relationships>`;
fs.writeFileSync(path.join(slidesDir, '_rels', 'slide3.xml.rels'), slide3Rels, 'utf-8');

// notesSlide3
fs.mkdirSync(path.join(TEMP, 'ppt', 'notesSlides', '_rels'), { recursive: true });
const notesSlide3 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesSlide xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6858000" cy="12192000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld>
</p:notesSlide>`;
fs.writeFileSync(path.join(TEMP, 'ppt', 'notesSlides', 'notesSlide3.xml'), notesSlide3, 'utf-8');

const notesSlide3Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide3.xml"/>
</Relationships>`;
fs.writeFileSync(path.join(TEMP, 'ppt', 'notesSlides', '_rels', 'notesSlide3.xml.rels'), notesSlide3Rels, 'utf-8');

// ── Update existing files ──

// presentation.xml — add sldId for slide 3
let presXml = fs.readFileSync(path.join(TEMP, 'ppt', 'presentation.xml'), 'utf-8');
presXml = presXml.replace('</p:sldIdLst>', `<p:sldId id="${SLIDE_ID}" r:id="${RID}"/></p:sldIdLst>`);
fs.writeFileSync(path.join(TEMP, 'ppt', 'presentation.xml'), presXml, 'utf-8');

// presentation.xml.rels
let presRels = fs.readFileSync(path.join(TEMP, 'ppt', '_rels', 'presentation.xml.rels'), 'utf-8');
presRels = presRels.replace('</Relationships>', `<Relationship Id="${RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/></Relationships>`);
fs.writeFileSync(path.join(TEMP, 'ppt', '_rels', 'presentation.xml.rels'), presRels, 'utf-8');

// notesMaster.xml.rels
const nmRelsPath = path.join(TEMP, 'ppt', 'notesMasters', '_rels', 'notesMaster1.xml.rels');
if (fs.existsSync(nmRelsPath)) {
  let nmRels = fs.readFileSync(nmRelsPath, 'utf-8');
  nmRels = nmRels.replace('</Relationships>', `<Relationship Id="${NOTES_RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide3.xml"/></Relationships>`);
  fs.writeFileSync(nmRelsPath, nmRels, 'utf-8');
}

// [Content_Types].xml
let ctXml = fs.readFileSync(path.join(TEMP, '[Content_Types].xml'), 'utf-8');
ctXml = ctXml.replace('</Types>',
  '<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
  '<Override PartName="/ppt/notesSlides/notesSlide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' +
  '</Types>');
fs.writeFileSync(path.join(TEMP, '[Content_Types].xml'), ctXml, 'utf-8');

// ── Re-zip ──
if (fs.existsSync(OUT)) fs.unlinkSync(OUT);
const output = fs.createWriteStream(OUT);
const archive = new ZipArchive({ zlib: { level: 9 } });
await new Promise((resolve, reject) => {
  output.on('close', resolve);
  archive.on('error', reject);
  archive.pipe(output);
  archive.directory(TEMP, false);
  archive.finalize();
});

const now = new Date();
const ts = now.getFullYear() +
  String(now.getMonth() + 1).padStart(2, '0') +
  String(now.getDate()).padStart(2, '0') + '_' +
  String(now.getHours()).padStart(2, '0') +
  String(now.getMinutes()).padStart(2, '0') +
  String(now.getSeconds()).padStart(2, '0');
const BACKUP = path.join(__dirname, `AEB_Introduction_${ts}.pptx`);

fs.copyFileSync(PPTX, BACKUP);
fs.unlinkSync(PPTX);
fs.copyFileSync(OUT, PPTX);
fs.unlinkSync(OUT);
fs.rmSync(TEMP, { recursive: true, force: true });

console.log(`Done! Backup: AEB_Introduction_${ts}.pptx`);
