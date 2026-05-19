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
const OUT = path.join(__dirname, 'AEB_Introduction_new.pptx');

// Generate timestamp: YYYYMMDD_HHmmss
const now = new Date();
const ts = now.getFullYear() +
    String(now.getMonth() + 1).padStart(2, '0') +
    String(now.getDate()).padStart(2, '0') + '_' +
    String(now.getHours()).padStart(2, '0') +
    String(now.getMinutes()).padStart(2, '0') +
    String(now.getSeconds()).padStart(2, '0');
const BACKUP = path.join(__dirname, `AEB_Introduction_${ts}.pptx`);

// Clean temp
fs.rmSync(TEMP, { recursive: true, force: true });
fs.mkdirSync(TEMP, { recursive: true });

// Unzip using unzipper
await new Promise((resolve, reject) => {
    fs.createReadStream(PPTX)
        .pipe(unzipper.Extract({ path: TEMP }))
        .on('close', resolve)
        .on('error', reject);
});

// EMU units: 1 inch = 914400 EMU, 1 cm = 360000 EMU
// Slide is 12192000 x 6858000 EMU (33.867cm x 19.05cm) = 16:9

const slide2Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Slide 2">
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="1D0638"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="TopLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="54864"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="442EE0"/></a:solidFill>
          <a:ln/>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="LeftBar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="73152" cy="6858000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="442EE0"/></a:solidFill>
          <a:ln/>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="GreenDot"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="548640" y="411480"/><a:ext cx="228600" cy="228600"/></a:xfrm>
          <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
          <a:ln/>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="CARIAD"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="868680" y="384048"/><a:ext cx="2743200" cy="320040"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1100" dirty="0">
                <a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>C A R I A D</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="6" name="MainTitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="700000"/><a:ext cx="10972800" cy="700000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="3600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>C-NCAP 2027 </a:t>
            </a:r>
            <a:r>
              <a:rPr lang="en-US" sz="3600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>AEB</a:t>
            </a:r>
            <a:r>
              <a:rPr lang="en-US" sz="3600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t> 测试场景总览</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="7" name="Underline"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="1400000"/><a:ext cx="2743200" cy="45720"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
          <a:ln/>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="8" name="Subtitle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="1500000"/><a:ext cx="10972800" cy="320040"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="zh-CN" sz="1600" dirty="0">
                <a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>AEB / FCW / AES | 车对弱势道路使用者 + 车对车 | 基础 + 拓展场景</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="9" name="VLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="6131451" y="2050000"/><a:ext cx="45719" cy="3600000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="442EE0"/></a:solidFill>
          <a:ln/>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="10" name="LeftHeader"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="2000000"/><a:ext cx="5100000" cy="320040"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1300" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>附录P — VRU / 弱势道路使用者保护</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="11" name="LeftList"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="2350000"/><a:ext cx="5100000" cy="3800000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          ${buildVRULines()}
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="20" name="RightHeader"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="6636028" y="2000000"/><a:ext cx="5100000" cy="320040"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1300" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>附录Q — C2C / 车对车 ADAS 测试</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="21" name="RightList"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="6636028" y="2350000"/><a:ext cx="5100000" cy="3800000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          ${buildC2CLines()}
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="90" name="BtmLine"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="5875616"/><a:ext cx="11064240" cy="10973"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="373741"/></a:solidFill>
          <a:ln/>
        </p:spPr>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="91" name="Stat1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="5967056"/><a:ext cx="2743200" cy="292608"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>6+10</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="92" name="Stat1Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="6213944"/><a:ext cx="2743200" cy="182880"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="850" dirty="0">
                <a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>车对车测试场景 (必测+抽测)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="93" name="Stat2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="3292834" y="5967056"/><a:ext cx="2743200" cy="292608"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>31</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="94" name="Stat2Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="3292834" y="6213944"/><a:ext cx="2743200" cy="182880"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="850" dirty="0">
                <a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>VRU 测试场景 (行/二轮/三轮)</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="95" name="Stat3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="6357863" y="5967056"/><a:ext cx="2743200" cy="292608"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>0–120 km/h</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="96" name="Stat3Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="6357863" y="6191680"/><a:ext cx="2743200" cy="182880"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="850" dirty="0">
                <a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>测试速度范围</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="97" name="Stat4"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="9123923" y="5967056"/><a:ext cx="2743200" cy="292608"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1600" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="1EEF97"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>2027</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="98" name="Stat4Label"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="9123923" y="6191680"/><a:ext cx="2743200" cy="182880"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="850" dirty="0">
                <a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>C-NCAP 版本</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="99" name="Footer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="526774" y="6450000"/><a:ext cx="4572000" cy="137160"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="700" dirty="0">
                <a:solidFill><a:srgbClr val="888888"/></a:solidFill>
                <a:latin typeface="Arial" pitchFamily="34" charset="0"/>
                <a:ea typeface="Arial" pitchFamily="34" charset="-122"/>
              </a:rPr>
              <a:t>INTERNAL / CONFIDENTIAL  |  CSD class: 2.5  |  Source: C-NCAP 2027 Management Rules</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;

function buildVRULines() {
    const items = [
        {title: '车对行人 (6项基础 · 权重70%)', detail: 'CPLA 纵向冲突 / CPFA 远端横穿 / CPNYC 幼儿近端横穿 / CPTA-LF 左转横穿 / CPTA-RF 右转横穿 / CPYCL 幼儿低速', color: '1EEF97'},
        {title: '车对行人·拓展 (6项 · 权重30%)', detail: 'CPFAr 雨天远端横穿 / CPFC-MO 动态遮挡儿童 / CPNC-AI 行人干扰儿童 / CPFC-BI 二轮车干扰儿童 / CPFAh 高速远端横穿 / CPLAh 高速纵向'},
        {title: '车对二轮车 (6项基础 · 权重70%)', detail: 'CBLAb 电动自行车纵向(夜) / CBOA 电动自行车斜向 / CBNA 自行车近端横穿 / CSFA 踏板摩托远端横穿 / CSTA-LN 左转踏板摩托 / CSTA-RN 右转踏板摩托', color: '1EEF97'},
        {title: '车对二轮车·拓展 (10项 · 权重30%)', detail: 'CBLAt 转弯切入 / CBLAc 变道切入 / CBOFA 斜穿遮挡 / CBNA-AI 行人干扰 / CSFA-BI 二轮车干扰 / CBFA-MO 动态遮挡 / CSTAO-RN 右转遮挡 / C2S RCPf 右转远端 / C2B RCPn 右转近端 / CSFhol 直行对向'},
        {title: '车对三轮车 (3项基础)', detail: 'CTRc 切入追尾 / CTFA 横穿 / CTFT 转弯对撞', color: '1EEF97'},
        {title: '功能覆盖', detail: 'AEB + FCW + AES (自动紧急转向/避让)'},
    ];

    let lines = '';
    for (const item of items) {
        const color = item.color || '442EE0';
        lines += `<a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
        <a:r><a:rPr lang="zh-CN" sz="900" b="1" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${item.title}</a:t></a:r></a:p>
        <a:p><a:pPr marL="171450" indent="0" algn="l"><a:buNone/></a:pPr>
        <a:r><a:rPr lang="zh-CN" sz="750" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${item.detail}</a:t></a:r></a:p>
        <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr><a:endParaRPr lang="zh-CN" sz="300" dirty="0"/></a:p>`;
    }
    return lines;
}

function buildC2CLines() {
    const items = [
        {title: '基础场景 (6项 · 必测)', detail: 'CCRc 切入追尾 / CCRh 高速追尾 / CCRb 制动追尾 / C2C SCPf 远端横穿 / CCFT 转弯对撞 / CCRbc 弯道制动', color: '1EEF97'},
        {title: '拓展场景 (10项 · 抽测验证)', detail: 'CCRsr 雨天追尾 / CCRsf 雾天追尾 / CCFhos 直行对撞 / C2C SCPn 近端横穿 / SCPmo 动态遮挡横穿 / CCFTf 跟车转弯对撞 / LCPf/LCPn 左转横穿 / RCPf 右转横穿 / SCPso 静态遮挡横穿'},
        {title: '测试速度范围', detail: '基础: 40–120 km/h | 拓展: 30–120 km/h'},
        {title: '功能覆盖', detail: 'AEB + FCW + AES + LKA + ELK + DMS + BSD + DOW + RCTB + ISLS + AMAP'},
    ];

    let lines = '';
    for (const item of items) {
        const color = item.color || '442EE0';
        lines += `<a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr>
        <a:r><a:rPr lang="zh-CN" sz="900" b="1" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${item.title}</a:t></a:r></a:p>
        <a:p><a:pPr marL="171450" indent="0" algn="l"><a:buNone/></a:pPr>
        <a:r><a:rPr lang="zh-CN" sz="750" dirty="0"><a:solidFill><a:srgbClr val="CDCDD2"/></a:solidFill><a:latin typeface="Arial" pitchFamily="34" charset="0"/><a:ea typeface="Microsoft YaHei" pitchFamily="34" charset="-122"/></a:rPr><a:t>${item.detail}</a:t></a:r></a:p>
        <a:p><a:pPr marL="0" indent="0" algn="l"><a:buNone/></a:pPr><a:endParaRPr lang="zh-CN" sz="300" dirty="0"/></a:p>`;
    }
    return lines;
}

// Write slide2.xml
const slidesDir = path.join(TEMP, 'ppt', 'slides');
fs.mkdirSync(path.join(slidesDir, '_rels'), { recursive: true });
fs.writeFileSync(path.join(slidesDir, 'slide2.xml'), slide2Xml, 'utf-8');

// Write slide2.xml.rels
const slide2Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide2.xml"/>
</Relationships>`;
fs.writeFileSync(path.join(slidesDir, '_rels', 'slide2.xml.rels'), slide2Rels, 'utf-8');

// Write notesSlide2.xml
fs.mkdirSync(path.join(TEMP, 'ppt', 'notesSlides', '_rels'), { recursive: true });
const notesSlide2 = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notesSlide xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6858000" cy="12192000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld>
</p:notesSlide>`;
fs.writeFileSync(path.join(TEMP, 'ppt', 'notesSlides', 'notesSlide2.xml'), notesSlide2, 'utf-8');

const notesSlide2Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/>
</Relationships>`;
fs.writeFileSync(path.join(TEMP, 'ppt', 'notesSlides', '_rels', 'notesSlide2.xml.rels'), notesSlide2Rels, 'utf-8');

// Update presentation.xml
let presXml = fs.readFileSync(path.join(TEMP, 'ppt', 'presentation.xml'), 'utf-8');
presXml = presXml.replace('</p:sldIdLst>', '<p:sldId id="257" r:id="rId8"/></p:sldIdLst>');
fs.writeFileSync(path.join(TEMP, 'ppt', 'presentation.xml'), presXml, 'utf-8');

// Update presentation.xml.rels
let presRels = fs.readFileSync(path.join(TEMP, 'ppt', '_rels', 'presentation.xml.rels'), 'utf-8');
presRels = presRels.replace('</Relationships>', '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>');
fs.writeFileSync(path.join(TEMP, 'ppt', '_rels', 'presentation.xml.rels'), presRels, 'utf-8');

// Update notesMaster.xml.rels
const nmRelsPath = path.join(TEMP, 'ppt', 'notesMasters', '_rels', 'notesMaster1.xml.rels');
if (fs.existsSync(nmRelsPath)) {
    let nmRels = fs.readFileSync(nmRelsPath, 'utf-8');
    nmRels = nmRels.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide2.xml"/></Relationships>');
    fs.writeFileSync(nmRelsPath, nmRels, 'utf-8');
}

// Update [Content_Types].xml
let ctXml = fs.readFileSync(path.join(TEMP, '[Content_Types].xml'), 'utf-8');
ctXml = ctXml.replace('</Types>',
    '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
    '<Override PartName="/ppt/notesSlides/notesSlide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>' +
    '</Types>');
fs.writeFileSync(path.join(TEMP, '[Content_Types].xml'), ctXml, 'utf-8');

// Re-zip using archiver
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

// Timestamped backup of original
fs.copyFileSync(PPTX, BACKUP);

// Replace original
fs.unlinkSync(PPTX);
fs.copyFileSync(OUT, PPTX);
fs.unlinkSync(OUT);

// Clean up temp
fs.rmSync(TEMP, { recursive: true, force: true });

console.log(`Done! Backup saved as: AEB_Introduction_${ts}.pptx`);
