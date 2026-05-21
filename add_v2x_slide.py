"""Add V2X slide based on C-NCAP 2027 document. Only scenarios that explicitly mention V2X/C-V2X in the original document are included."""
import copy, io, os, re, zipfile
from lxml import etree

PPX = r'C:\Users\86132\Desktop\test-2\AEB_Introduction_base.pptx'
OUT = r'C:\Users\86132\Desktop\test-2\AEB_Introduction_fixed.pptx'

LEFT_DATA = [
    # (text, sz/100 EMU, bold, green_color, indent)
    ("附录P — VRU 弱势道路使用者保护", 950, True, True, 0),
    ("", 300, False, False, 0),
    ("结论：附录P中未明确提及 V2X / C-V2X 技术", 800, False, False, 171450),
    ("", 300, False, False, 0),
    ("附录P的31个VRU场景均基于车载传感器（摄像头、雷达、", 750, False, False, 171450),
    ("激光雷达）进行感知。V2P（车与行人/二轮车通信）尚未", 750, False, False, 171450),
    ("纳入C-NCAP 2027的VRU测试评价规程。", 750, False, False, 171450),
    ("", 300, False, False, 0),
    ("这意味着：当前VRU保护测试仅评估单车智能的感知能力，", 750, False, False, 171450),
    ("V2X在弱势道路使用者保护领域的标准化测试尚属空白。", 750, False, False, 171450),
]

RIGHT_DATA = [
    # Appendix Q scenes that explicitly reference V2X
    ("附录Q — C2C 车对车 ADAS · 涉及 V2X 的场景", 950, True, True, 0),
    ("", 300, False, False, 0),
    ("GVT 目标车 · 可具备 C-V2X 直连通信能力", 850, True, True, 171450),
    ("附录Q规定GVT（全局车辆目标物）可集成C-V2X直连通信", 700, False, False, 171450),
    ("模块，并具备第三方电子认证机构认可的互信能力。（见", 700, False, False, 171450),
    ("Q.3.6.1.2.4章节，第13/20/25页）", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("当车辆具备 C-V2X 时，以下为必测场景：", 850, True, True, 171450),
    ("", 300, False, False, 0),
    ("① CCRh — 高速追尾", 800, True, False, 171450),
    ("远距离前车减速预警，V2V可增加感知距离，弥补传感器", 700, False, False, 171450),
    ("探测范围限制，实现超视距协同制动", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("② CCFhos — 直行对撞", 800, True, False, 171450),
    ("有遮挡的交叉路口对向碰撞场景，V2V穿透建筑物/车辆", 700, False, False, 171450),
    ("遮挡，提前获知对向来车位置与速度", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("③ C2C SCPso — 静态遮挡横穿", 800, True, False, 171450),
    ("目标车被路边停放车辆遮挡后横穿，V2V可越过静态障碍", 700, False, False, 171450),
    ("感知被遮挡移动车辆，解决非视距（NLOS）问题", 700, False, False, 171450),
    ("", 300, False, False, 0),
    ("上述场景不参与排序抽取，直接列为必测（Q.3.6.1.7.11.5）", 700, False, False, 171450),
]

NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

def qn(ns, tag):
    return '{%s}%s' % (NSMAP[ns], tag)

def build_text_lines(data, green_color, grey_color, latin_font, ea_font):
    lines = []
    for text, sz, bold, is_green, indent in data:
        if not text:
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


def update_slide_content(slide_tree):
    sld_ns = '{%s}' % NSMAP['p']
    a_ns = '{%s}' % NSMAP['a']
    green_color = '1EEF97'
    grey_color = 'CDCDD2'
    latin_font = 'Arial'
    ea_font = 'Microsoft YaHei'

    sp_tree = slide_tree.find('.//' + sld_ns + 'spTree')

    # Collect text shapes (txBody is in p: namespace)
    name_to_txBody = {}
    for sp in sp_tree:
        nvSpPr = sp.find(sld_ns + 'nvSpPr')
        if nvSpPr is None:
            continue
        cNvPr = nvSpPr.find(sld_ns + 'cNvPr')
        if cNvPr is None:
            continue
        name = cNvPr.get('name', '')
        if not name or name == 'CARIAD':
            continue
        txBody = sp.find(sld_ns + 'txBody')
        if txBody is not None:
            name_to_txBody[name] = txBody

    def clear_and_fill(txBody, new_ps):
        for p in list(txBody):
            txBody.remove(p)
        for np in new_ps:
            txBody.append(np)

    def set_single_line(txBody, text, sz='3200', color='FFFFFF', lang='zh-CN', is_bold='1'):
        for p in list(txBody):
            txBody.remove(p)
        p = etree.SubElement(txBody, a_ns + 'p')
        pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
        pPr.append(etree.Element(a_ns + 'buNone'))
        r = etree.SubElement(p, a_ns + 'r')
        rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': lang, 'sz': sz, 'b': is_bold, 'dirty': '0'})
        sf = etree.SubElement(rPr, a_ns + 'solidFill')
        etree.SubElement(sf, a_ns + 'srgbClr', {'val': color})
        etree.SubElement(rPr, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
        etree.SubElement(rPr, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
        t = etree.SubElement(r, a_ns + 't')
        t.text = text

    # ── Title: "C-NCAP 2027 AEB 场景中的 V2X 技术" ──
    if 'MainTitle' in name_to_txBody:
        txBody = name_to_txBody['MainTitle']
        for p in list(txBody):
            txBody.remove(p)
        p = etree.SubElement(txBody, a_ns + 'p')
        pPr = etree.SubElement(p, a_ns + 'pPr', {'marL': '0', 'indent': '0', 'algn': 'l'})
        pPr.append(etree.Element(a_ns + 'buNone'))

        def add_run(parent, text, color, sz='3200', b='1'):
            r = etree.SubElement(parent, a_ns + 'r')
            rPr = etree.SubElement(r, a_ns + 'rPr', {'lang': 'zh-CN', 'sz': sz, 'b': b, 'dirty': '0'})
            sf = etree.SubElement(rPr, a_ns + 'solidFill')
            etree.SubElement(sf, a_ns + 'srgbClr', {'val': color})
            etree.SubElement(rPr, a_ns + 'latin', {'typeface': 'Arial', 'pitchFamily': '34', 'charset': '0'})
            etree.SubElement(rPr, a_ns + 'ea', {'typeface': 'Microsoft YaHei', 'pitchFamily': '34', 'charset': '-122'})
            t = etree.SubElement(r, a_ns + 't')
            t.text = text

        add_run(p, 'C-NCAP 2027 AEB 场景中的 ', 'FFFFFF')
        add_run(p, 'V2X / C-V2X', '1EEF97')
        add_run(p, ' 技术', 'FFFFFF')

    # ── Subtitle ──
    if 'Subtitle' in name_to_txBody:
        set_single_line(name_to_txBody['Subtitle'],
            '基于C-NCAP 2027版管理规则原文 · 仅列出文档中明确提及V2X/C-V2X的测试场景',
            sz='1400', color='CDCDD2', lang='zh-CN', is_bold='0')

    # ── Left header ──
    if 'LeftHeader' in name_to_txBody:
        set_single_line(name_to_txBody['LeftHeader'],
            '附录P — VRU保护 | V2X提及情况',
            sz='1100', color='1EEF97', lang='zh-CN', is_bold='1')

    # ── Right header ──
    if 'RightHeader' in name_to_txBody:
        set_single_line(name_to_txBody['RightHeader'],
            '附录Q — C2C ADAS | 涉及 C-V2X 的场景（文档原文）',
            sz='1100', color='1EEF97', lang='zh-CN', is_bold='1')

    # ── Left list ──
    if 'LeftList' in name_to_txBody:
        clear_and_fill(name_to_txBody['LeftList'], build_text_lines(LEFT_DATA, green_color, grey_color, latin_font, ea_font))

    # ── Right list ──
    if 'RightList' in name_to_txBody:
        clear_and_fill(name_to_txBody['RightList'], build_text_lines(RIGHT_DATA, green_color, grey_color, latin_font, ea_font))

    # ── Stats row ──
    stat_specs = [
        ('Stat1', '3', '1EEF97', '1600', 'bold'),    ('Stat1Label', 'V2X必测场景数', 'CDCDD2', '800', 'normal'),
        ('Stat2', 'GVT', '1EEF97', '1600', 'bold'),   ('Stat2Label', '目标物支持C-V2X', 'CDCDD2', '800', 'normal'),
        ('Stat3', 'NLOS', '1EEF97', '1600', 'bold'),   ('Stat3Label', '核心解决场景', 'CDCDD2', '800', 'normal'),
        ('Stat4', '附录Q', '1EEF97', '1600', 'bold'),   ('Stat4Label', 'V2X仅出现在附录Q', 'CDCDD2', '800', 'normal'),
    ]
    for name, text, color, sz, style in stat_specs:
        if name in name_to_txBody:
            set_single_line(name_to_txBody[name], text, sz=sz, color=color, is_bold='1' if style == 'bold' else '0')

    # ── Footer ──
    if 'Footer' in name_to_txBody:
        set_single_line(name_to_txBody['Footer'],
            'INTERNAL / CONFIDENTIAL  |  CSD class: 2.5  |  Source: C-NCAP 2027 附录Q (Q.3.6.1.2.4 / Q.3.6.1.7.11.5)',
            sz='700', color='888888', is_bold='0')

    # ── Update cSld name ──
    cSld = slide_tree.find(sld_ns + 'cSld')
    if cSld is not None:
        cSld.set('name', 'Slide 3 - V2X in C-NCAP 2027')

    return slide_tree


# ── Main ──
with open(PPX, 'rb') as f:
    pptx_bytes = f.read()

zin = zipfile.ZipFile(io.BytesIO(pptx_bytes), 'r')
zout = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)

slide2_path = 'ppt/slides/slide2.xml'
slide2_rels_path = 'ppt/slides/_rels/slide2.xml.rels'
notes2_path = 'ppt/notesSlides/notesSlide2.xml'
notes2_rels_path = 'ppt/notesSlides/_rels/notesSlide2.xml.rels'

slide3_xml = zin.read(slide2_path)
slide3_rels_xml = zin.read(slide2_rels_path)
notes3_xml = zin.read(notes2_path)
notes3_rels_xml = zin.read(notes2_rels_path)

slide3_tree = etree.fromstring(slide3_xml)
slide3_tree = update_slide_content(slide3_tree)

for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'ppt/presentation.xml':
        xml = data.decode('utf-8')
        xml = xml.replace('</p:sldIdLst>',
            '<p:sldId id="261" r:id="rId10"/></p:sldIdLst>')
        data = xml.encode('utf-8')
    elif item.filename == 'ppt/_rels/presentation.xml.rels':
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

zout.writestr('ppt/slides/slide3.xml', etree.tostring(slide3_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

# slide3 rels — clone from slide2, update targets
slide3_rels_tree = etree.fromstring(slide3_rels_xml)
for rel in slide3_rels_tree:
    target = rel.get('Target', '')
    target = target.replace('slide2', 'slide3').replace('notesSlide2', 'notesSlide3')
    rel.set('Target', target)
zout.writestr('ppt/slides/_rels/slide3.xml.rels', etree.tostring(slide3_rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

# notesSlide3
notes3_tree = etree.fromstring(notes3_xml)
zout.writestr('ppt/notesSlides/notesSlide3.xml', etree.tostring(notes3_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

notes3_rels_tree = etree.fromstring(notes3_rels_xml)
for rel in notes3_rels_tree:
    target = rel.get('Target', '')
    target = target.replace('slide2', 'slide3')
    rel.set('Target', target)
zout.writestr('ppt/notesSlides/_rels/notesSlide3.xml.rels', etree.tostring(notes3_rels_tree, xml_declaration=True, encoding='UTF-8', standalone=True))

zin.close()
zout.close()

print(f'Fixed PPTX saved to: {OUT}')
print('Verify with: python-pptx Presentation(OUT)')
