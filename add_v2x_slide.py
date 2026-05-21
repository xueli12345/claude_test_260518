"""Clone slide 2 as slide 3 with V2X content, preserving layout and styling."""
import copy, io, os, re, zipfile
from lxml import etree

PPX = r'C:\Users\86132\Desktop\test-2\AEB_Introduction_base.pptx'
OUT = r'C:\Users\86132\Desktop\test-2\AEB_Introduction_fixed.pptx'

# ── V2X slide text (each entry: (text_content, font_size_EMU, bold, green_color)) ──
LEFT_DATA = [
    # (text, sz/100 EMU, bold, green, indent)
    ("V2P 穿透遮挡 — 行人/二轮车感知", 900, True, True, 0),
    ("CPFC-MO 动态遮挡儿童 — 被挡儿童可被V2P提前感知", 700, False, False, 171450),
    ("CPNC-AI / CPFC-BI 干扰横穿 — 被遮挡弱势道路使用者", 700, False, False, 171450),
    ("CBOFA / CBFA-MO 遮挡斜穿 — 穿透车辆障碍物", 700, False, False, 171450),
    ("CBNA-AI / CSFA-BI 干扰近端 — V2P弥补传感器盲区", 700, False, False, 171450),
    ("CSTAO-RN 右转遮挡 — 路侧V2I+车端V2P协同", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("V2P 转弯/横穿预警", 900, True, True, 0),
    ("CPTA-LF / CPTA-RF 转弯横穿行人 — 提前感知意图", 700, False, False, 171450),
    ("C2S RCPf / C2B RCPn 转弯横穿二轮 — 路口盲区", 700, False, False, 171450),
    ("CBLAt 转弯切入 / CBLAc 变道切入 — 意图共享", 700, False, False, 171450),
    ("CSFhol 直行对向二轮 — 远距离V2P预警", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("恶劣环境 — 天气/光照不敏感", 900, True, False, 0),
    ("CPFAr 雨天远端横穿 — 摄像头受限，V2X不受影响", 700, False, False, 171450),
    ("CBLAb 电动自行车夜间 — 弱光下V2P优于视觉", 700, False, False, 171450),
    ("CPFAh / CPLAh 高速场景 — 远距离V2P提前预警", 700, False, False, 171450),
]

RIGHT_DATA = [
    ("V2V 遮挡/非视距 — 穿透障碍感知", 900, True, True, 0),
    ("C2C SCPmo 动态遮挡横穿 — 被挡车辆V2V预警", 700, False, False, 171450),
    ("C2C SCPso 静态遮挡横穿 — 越过建筑物感知", 700, False, False, 171450),
    ("CCFT / CCFTf 转弯对撞 — 路口建筑遮挡，V2V直连", 700, False, False, 171450),
    ("LCPf / LCPn 左转横穿 — 对向遮挡车辆提前通信", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("V2V 恶劣天气 — 传感器降级补偿", 900, True, True, 0),
    ("CCRsr 雨天追尾 — 雷达/摄像头衰减，V2V不受影响", 700, False, False, 171450),
    ("CCRsf 雾天追尾 — 能见度极低，V2V稳定通信", 700, False, False, 171450),
    ("CCFhos 直行对撞 — 雨雾天远距离V2V预警", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("V2V 意图共享 — 切入/制动预警", 900, True, False, 0),
    ("CCRc 切入追尾 — 邻车变道意图实时共享", 700, False, False, 171450),
    ("CCRb 制动追尾 — 前车制动信号低延迟广播", 700, False, False, 171450),
    ("CCRh 高速追尾 / CCRbc 弯道制动 — 超视距补盲", 700, False, False, 171450),
    ("RCPf 右转横穿 — 转向意图+位置共享", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("V2I 路侧协同", 900, True, False, 0),
    ("全部路口转弯/横穿场景可通过V2I路侧单元(RSU)", 700, False, False, 171450),
    ("获取路口全局目标列表，弥补单车感知盲区", 700, False, False, 171450),
    ("典型: CPTA系列 / CCFT系列 / C2C LCP&SCP系列", 700, False, False, 171450),
]

# XML namespaces
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}

def qn(ns, tag):
    """Create qualified name."""
    return '{%s}%s' % (NSMAP[ns], tag)

def clone_slide(zf, slide2_path, slide2_rels_path):
    """Clone slide2.xml as slide3.xml."""
    slide2_xml = zf.read(slide2_path)
    slide2_rels_xml = zf.read(slide2_rels_path) if slide2_rels_path in zf.namelist() else None

    # Parse
    slide_tree = etree.fromstring(slide2_xml)
    slide_rels_tree = etree.fromstring(slide2_rels_xml) if slide2_rels_xml else None

    return slide_tree, slide_rels_tree

def build_text_lines(data, green_color, grey_color, latin_font, ea_font):
    """Build <a:p> elements for text lines."""
    lines = []
    for text, sz, bold, is_green, indent in data:
        if not text:
            # Empty spacer paragraph
            p = etree.Element(qn('a', 'p'))
            etree.SubElement(p, qn('a', 'pPr'), {
                'marL': '0', 'indent': '0', 'algn': 'l'
            }).append(etree.Element(qn('a', 'buNone')))
            er = etree.SubElement(p, qn('a', 'endParaRPr'), {
                'lang': 'zh-CN', 'sz': str(sz), 'dirty': '0'
            })
            lines.append(p)
            continue

        color = green_color if is_green else grey_color
        p = etree.Element(qn('a', 'p'))
        pPr = etree.SubElement(p, qn('a', 'pPr'), {
            'marL': str(indent), 'indent': '0', 'algn': 'l'
        })
        pPr.append(etree.Element(qn('a', 'buNone')))

        r = etree.SubElement(p, qn('a', 'r'))
        rPr = etree.SubElement(r, qn('a', 'rPr'), {
            'lang': 'zh-CN', 'sz': str(sz),
        })
        if bold:
            rPr.set('b', '1')
        rPr.set('dirty', '0')

        sf = etree.SubElement(rPr, qn('a', 'solidFill'))
        etree.SubElement(sf, qn('a', 'srgbClr'), {'val': color})

        la = etree.SubElement(rPr, qn('a', 'latin'), {
            'typeface': latin_font, 'pitchFamily': '34', 'charset': '0'
        })
        ea = etree.SubElement(rPr, qn('a', 'ea'), {
            'typeface': ea_font, 'pitchFamily': '34', 'charset': '-122'
        })

        t = etree.SubElement(r, qn('a', 't'))
        t.text = text

        lines.append(p)
    return lines

def update_slide_content(slide_tree, side):
    """Update the content of a cloned slide with V2X data."""
    sld_ns = '{%s}' % NSMAP['p']
    a_ns = '{%s}' % NSMAP['a']

    green_color = '1EEF97'
    grey_color = 'CDCDD2'
    latin_font = 'Arial'
    ea_font = 'Microsoft YaHei'

    # Find all <p:sp> elements (shapes with text)
    sp_tree = slide_tree.find('.//' + sld_ns + 'spTree')

    # Collect all text shapes (txBody is in p: namespace!)
    text_shapes = []
    for sp in sp_tree:
        nvSpPr = sp.find(sld_ns + 'nvSpPr')
        cNvPr = nvSpPr.find(sld_ns + 'cNvPr') if nvSpPr is not None else None
        name = cNvPr.get('name', '') if cNvPr is not None else ''
        txBody = sp.find(sld_ns + 'txBody')
        if txBody is not None and name not in ('CARIAD',):
            text_shapes.append((name, sp, txBody))

    # Identify left and right list containers by name pattern
    left_txBody = None
    right_txBody = None
    main_title_txBody = None
    subtitle_txBody = None
    left_header_txBody = None
    right_header_txBody = None
    footer_txBody = None

    # Stats containers
    stat_bodies = []

    for name, sp, txBody in text_shapes:
        if name == 'MainTitle':
            main_title_txBody = txBody
        elif name == 'Subtitle':
            subtitle_txBody = txBody
        elif name == 'LeftHeader':
            left_header_txBody = txBody
        elif name == 'RightHeader':
            right_header_txBody = txBody
        elif name == 'LeftList':
            left_txBody = txBody
        elif name == 'RightList':
            right_txBody = txBody
        elif name == 'Footer':
            footer_txBody = txBody
        elif name.startswith('Stat'):
            stat_bodies.append((name, txBody))

    # ── Update main title ──
    if main_title_txBody is not None:
        # Clear existing paragraphs, add new ones
        for p in list(main_title_txBody):
            main_title_txBody.remove(p)

        # Title: "V2X 赋能 C-NCAP 2027 AEB 测试场景"
        p = etree.SubElement(main_title_txBody, a_ns + 'p')
        pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
        pPr.append(etree.Element(a_ns + 'buNone'))

        # "V2X " in white
        r1 = etree.SubElement(p, a_ns + 'r')
        rPr1 = etree.SubElement(r1, a_ns + 'rPr', {'lang': 'en-US', 'sz': '3200', 'b': '1', 'dirty': '0'})
        sf1 = etree.SubElement(rPr1, a_ns + 'solidFill')
        etree.SubElement(sf1, a_ns + 'srgbClr', {'val': 'FFFFFF'})
        etree.SubElement(rPr1, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr1, a_ns + 'ea', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '-122'})
        t1 = etree.SubElement(r1, a_ns + 't')
        t1.text = 'V2X '

        # "赋能" in green
        r2 = etree.SubElement(p, a_ns + 'r')
        rPr2 = etree.SubElement(r2, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': '3200', 'b': '1', 'dirty': '0'})
        sf2 = etree.SubElement(rPr2, a_ns + 'solidFill')
        etree.SubElement(sf2, a_ns + 'srgbClr', {'val': '1EEF97'})
        etree.SubElement(rPr2, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr2, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
        t2 = etree.SubElement(r2, a_ns + 't')
        t2.text = '赋能'

        # " C-NCAP 2027 AEB 测试场景" in white
        r3 = etree.SubElement(p, a_ns + 'r')
        rPr3 = etree.SubElement(r3, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': '3200', 'b': '1', 'dirty': '0'})
        sf3 = etree.SubElement(rPr3, a_ns + 'solidFill')
        etree.SubElement(sf3, a_ns + 'srgbClr', {'val': 'FFFFFF'})
        etree.SubElement(rPr3, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr3, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
        t3 = etree.SubElement(r3, a_ns + 't')
        t3.text = ' C-NCAP 2027 AEB 测试场景'

    # ── Update subtitle ──
    if subtitle_txBody is not None:
        for p in list(subtitle_txBody):
            subtitle_txBody.remove(p)
        p = etree.SubElement(subtitle_txBody, a_ns + 'p')
        pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
        pPr.append(etree.Element(a_ns + 'buNone'))
        r = etree.SubElement(p, a_ns + 'r')
        rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': '1400', 'dirty': '0'})
        sf = etree.SubElement(rPr, a_ns + 'solidFill')
        etree.SubElement(sf, a_ns + 'srgbClr', {'val': 'CDCDD2'})
        etree.SubElement(rPr, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
        t = etree.SubElement(r, a_ns + 't')
        t.text = 'V2V 车车直连 | V2P 车与弱势道路使用者 | V2I 车路协同 | 非视距穿透 · 恶劣天气抗扰 · 意图共享'

    # ── Update left header ──
    if left_header_txBody is not None:
        for p in list(left_header_txBody):
            left_header_txBody.remove(p)
        p = etree.SubElement(left_header_txBody, a_ns + 'p')
        pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
        pPr.append(etree.Element(a_ns + 'buNone'))
        r = etree.SubElement(p, a_ns + 'r')
        rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': '1100', 'b': '1', 'dirty': '0'})
        sf = etree.SubElement(rPr, a_ns + 'solidFill')
        etree.SubElement(sf, a_ns + 'srgbClr', {'val': '1EEF97'})
        etree.SubElement(rPr, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
        t = etree.SubElement(r, a_ns + 't')
        t.text = '附录P — VRU场景 | V2P / 车与弱势道路使用者'

    # ── Update right header ──
    if right_header_txBody is not None:
        for p in list(right_header_txBody):
            right_header_txBody.remove(p)
        p = etree.SubElement(right_header_txBody, a_ns + 'p')
        pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
        pPr.append(etree.Element(a_ns + 'buNone'))
        r = etree.SubElement(p, a_ns + 'r')
        rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': '1100', 'b': '1', 'dirty': '0'})
        sf = etree.SubElement(rPr, a_ns + 'solidFill')
        etree.SubElement(sf, a_ns + 'srgbClr', {'val': '1EEF97'})
        etree.SubElement(rPr, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
        t = etree.SubElement(r, a_ns + 't')
        t.text = '附录Q — C2C场景 | V2V / V2I 车车与车路协同'

    # ── Update left list ──
    if left_txBody is not None:
        for p in list(left_txBody):
            left_txBody.remove(p)
        new_ps = build_text_lines(LEFT_DATA, green_color, grey_color, latin_font, ea_font)
        for np in new_ps:
            left_txBody.append(np)

    # ── Update right list ──
    if right_txBody is not None:
        for p in list(right_txBody):
            right_txBody.remove(p)
        new_ps = build_text_lines(RIGHT_DATA, green_color, grey_color, latin_font, ea_font)
        for np in new_ps:
            right_txBody.append(np)

    # ── Update stats ──
    stat_texts = [
        ('~70%', '场景可由V2X增强'),
        ('V2V+P+I', 'V2X通信模式协同'),
        ('NLOS', '非视距穿透是核心优势'),
        ('全天候', '雨雾夜不受影响'),
    ]
    for i, (name, txBody) in enumerate(sorted(stat_bodies)):
        if i < len(stat_texts) * 2:
            idx = i // 2
            is_value = (i % 2 == 0)
            for p in list(txBody):
                txBody.remove(p)
            p = etree.SubElement(txBody, a_ns + 'p')
            pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
            pPr.append(etree.Element(a_ns + 'buNone'))
            r = etree.SubElement(p, a_ns + 'r')
            if is_value:
                rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': 'en-US', 'sz': '1600', 'b': '1', 'dirty': '0'})
                color = '1EEF97'
            else:
                rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': '800', 'dirty': '0'})
                color = 'CDCDD2'
            sf = etree.SubElement(rPr, a_ns + 'solidFill')
            etree.SubElement(sf, a_ns + 'srgbClr', {'val': color})
            etree.SubElement(rPr, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
            etree.SubElement(rPr, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
            t = etree.SubElement(r, a_ns + 't')
            t.text = stat_texts[idx][0 if is_value else 1]

    # ── Update footer ──
    if footer_txBody is not None:
        for p in list(footer_txBody):
            footer_txBody.remove(p)
        p = etree.SubElement(footer_txBody, a_ns + 'p')
        pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
        pPr.append(etree.Element(a_ns + 'buNone'))
        r = etree.SubElement(p, a_ns + 'r')
        rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': '700', 'dirty': '0'})
        sf = etree.SubElement(rPr, a_ns + 'solidFill')
        etree.SubElement(sf, a_ns + 'srgbClr', {'val': '888888'})
        etree.SubElement(rPr, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
        t = etree.SubElement(r, a_ns + 't')
        t.text = 'INTERNAL / CONFIDENTIAL  |  CSD class: 2.5  |  基于 C-NCAP 2027 场景的 V2X 赋能分析'

    # ── Update cSld name ──
    cSld = slide_tree.find(sld_ns + 'cSld')
    if cSld is not None:
        cSld.set('name', 'Slide 3 - V2X')

    return slide_tree


# ── Main ──

# Read original PPTX
with open(PPX, 'rb') as f:
    pptx_bytes = f.read()

zin = zipfile.ZipFile(io.BytesIO(pptx_bytes), 'r')
zout = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)

# Clone slide2 -> slide3
slide2_path = 'ppt/slides/slide2.xml'
slide2_rels_path = 'ppt/slides/_rels/slide2.xml.rels'
notes2_path = 'ppt/notesSlides/notesSlide2.xml'
notes2_rels_path = 'ppt/notesSlides/_rels/notesSlide2.xml.rels'

slide3_xml = zin.read(slide2_path)
slide3_rels_xml = zin.read(slide2_rels_path)
notes3_xml = zin.read(notes2_path)
notes3_rels_xml = zin.read(notes2_rels_path)

# Update slide content
slide3_tree = etree.fromstring(slide3_xml)
slide3_tree = update_slide_content(slide3_tree, 'v2x')

# Write all original files + new slide3 files
for item in zin.infolist():
    # Read and write each file from original
    data = zin.read(item.filename)

    if item.filename == 'ppt/presentation.xml':
        # Add new sldId before </p:sldIdLst>
        xml = data.decode('utf-8')
        xml = xml.replace('</p:sldIdLst>',
            '<p:sldId id="261" r:id="rId10"/></p:sldIdLst>')
        data = xml.encode('utf-8')
    elif item.filename == 'ppt/_rels/presentation.xml.rels':
        # Add new relationship
        xml = data.decode('utf-8')
        xml = xml.replace('</Relationships>',
            '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/></Relationships>')
        data = xml.encode('utf-8')
    elif item.filename == 'ppt/notesMasters/_rels/notesMaster1.xml.rels':
        xml = data.decode('utf-8')
        xml = xml.replace('</Relationships>',
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide3.xml"/></Relationships>')
        data = xml.encode('utf-8')
    elif item.filename == '[Content_Types].xml':
        xml = data.decode('utf-8')
        xml = xml.replace('</Types>',
            '<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            '<Override PartName="/ppt/notesSlides/notesSlide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'
            '</Types>')
        data = xml.encode('utf-8')

    zout.writestr(item, data)

# Add slide3.xml
zout.writestr('ppt/slides/slide3.xml', etree.tostring(slide3_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

# Add slide3.xml.rels
slide3_rels_tree = etree.fromstring(slide3_rels_xml)
# Update relationship targets from slide2 to slide3
for rel in slide3_rels_tree:
    target = rel.get('Target', '')
    target = target.replace('slide2', 'slide3').replace('notesSlide2', 'notesSlide3')
    rel.set('Target', target)
zout.writestr('ppt/slides/_rels/slide3.xml.rels', etree.tostring(slide3_rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

# Add notesSlide3.xml
notes3_tree = etree.fromstring(notes3_xml)
zout.writestr('ppt/notesSlides/notesSlide3.xml', etree.tostring(notes3_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

# Add notesSlide3.xml.rels
notes3_rels_tree = etree.fromstring(notes3_rels_xml)
for rel in notes3_rels_tree:
    target = rel.get('Target', '')
    target = target.replace('slide2', 'slide3')
    rel.set('Target', target)
zout.writestr('ppt/notesSlides/_rels/notesSlide3.xml.rels', etree.tostring(notes3_rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

zin.close()
zout.close()

# Verify
vfy = zipfile.ZipFile(OUT, 'r')
names = sorted(vfy.namelist())
slide_files = [n for n in names if 'slide' in n.lower()]
print(f'Total files in PPTX: {len(names)}')
print(f'Slide files: {slide_files}')
print(f'Slides: slide1, slide2, slide3 — all present: {"slide3.xml" in names}')
vfy.close()

print(f'\nFixed PPTX saved to: {OUT}')
