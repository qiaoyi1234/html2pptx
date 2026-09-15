#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
html2ppt.py — 把你的「结构化 HTML 幻灯片」转成 PPTX

用法:
    python html2ppt.py 你的文件.html                  # 输出 你的文件.pptx
    python html2ppt.py 你的文件.html -o 自定义名.pptx

依赖:
    pip install python-pptx

支持的 HTML 约定（与 AI教育-思维外包.html 同款）:
    每个幻灯片是一个 <section class="slide">...</section>
    页内元素按 class 自动识别：
      .kicker     页眉小标题（青色小字）
      h1 / h2     页面大标题（h1 优先，<br> 自动换行）
      .lead       引语段   .sub  备注小字
      .badge .foot 封面徽标 / 封面页脚
      .cards .card 卡片组（自动 2 或 3 列）: .ic .h3 .p
      .bullets .b 要点列表（<b> 加粗，<small> 变灰色小字）
      .stat .box  数据块（.num 大数字 + .lbl 说明，class 含 up/down 控制颜色）
      .vs  .side  左右对比栏（h3 + li）
      .chain .link 多米诺链条（带 ▼ 箭头）
      .num3 .row  编号方案（.n 序号 + h4 + p）
      .quote      金句（深色引用条）
      .mirror     全宽深色面板
    封面: <section class="slide cover">   结尾: <section class="slide final">

    不认识的元素会打印警告并跳过；复杂 CSS（图表/动画/绝对定位）无法转换。
"""
import argparse
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------- 配色 ----------------
NAVY   = RGBColor(0x10, 0x2A, 0x4A)
INK    = RGBColor(0x1E, 0x29, 0x3B)
DIM    = RGBColor(0x5B, 0x6B, 0x8C)
FAINT  = RGBColor(0x8A, 0x97, 0xB0)
CYAN   = RGBColor(0x0E, 0x74, 0x90)
VIOLET = RGBColor(0x7C, 0x3A, 0xED)
PINK   = RGBColor(0xDB, 0x27, 0x77)
GREEN  = RGBColor(0x0F, 0x9D, 0x58)
PANEL  = RGBColor(0xF1, 0xF5, 0xFB)
PANEL2 = RGBColor(0xE7, 0xEF, 0xFA)
PANELP = RGBColor(0xFD, 0xEF, 0xF6)
BORDER = RGBColor(0xD5, 0xE0, 0xF0)
BORDERP= RGBColor(0xF3, 0xC6, 0xDC)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
NAVY_BG= RGBColor(0x0B, 0x1A, 0x33)
SKY    = RGBColor(0x7D, 0xD3, 0xF0)
LIGHT  = RGBColor(0xC7, 0xD4, 0xE8)
FAINTB = RGBColor(0x6E, 0x80, 0xA0)

FONT = "Microsoft YaHei"
SW, SH = 13.333, 7.5

# ---------------- HTML 轻度解析 ----------------
class Node:
    __slots__ = ("tag", "attrs", "children")
    def __init__(self, tag=None, attrs=None):
        self.tag = tag
        self.attrs = dict(attrs or [])
        self.children = []

class DeckParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("root")
        self.stack = [self.root]
    def handle_starttag(self, tag, attrs):
        n = Node(tag, attrs)
        self.stack[-1].children.append(n)
        if tag != "br":
            self.stack.append(n)
    def handle_startendtag(self, tag, attrs):
        self.stack[-1].children.append(Node(tag, attrs))
    def handle_endtag(self, tag):
        if tag == "br":
            return
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                break
    def handle_data(self, data):
        if data:
            self.stack[-1].children.append(data)

def has_class(node, cls):
    return cls in node.attrs.get("class", "").split()

def text_of(node):
    if isinstance(node, str):
        return node
    return "".join(text_of(c) for c in node.children)

def find_first(node, cls=None, tag=None):
    if not isinstance(node, Node):
        return None
    if (cls is None or has_class(node, cls)) and (tag is None or node.tag == tag):
        return node
    for c in node.children:
        r = find_first(c, cls, tag)
        if r is not None:
            return r
    return None

def all_li(node):
    """按文档顺序收集节点下所有 <li>（跨任意层级）"""
    out = []
    if isinstance(node, Node):
        if node.tag == "li":
            out.append(node)
        for c in node.children:
            out.extend(all_li(c))
    return out

def inline(node, parent_bold=False):
    """把节点转成段落列表；每段 = [(文本, 是否加粗), ...]。<br> 分段，<b>/<em>/<span.grad> 加粗。"""
    paras = []
    cur = []

    def flush():
        if cur:
            paras.append(cur[:])
            cur.clear()

    def walk(n, bold):
        if isinstance(n, str):
            if n:
                cur.append((n, bold))
            return
        if n.tag == "br":
            flush()
            return
        b = bold or n.tag in ("b", "strong", "em", "i") or has_class(n, "grad")
        for c in n.children:
            walk(c, b)

    walk(node, parent_bold)
    flush()
    return paras

def join_segs(segs):
    return "".join(t for t, _ in segs)

def flat(paras):
    """把段落列表压平成纯文本"""
    return "".join(t for p in paras for t, _ in p)

# ---------------- 内容提取 ----------------
def extract_slide(sec, idx):
    cls = sec.attrs.get("class", "")
    kind = "cover" if "cover" in cls else ("final" if "final" in cls else "content")
    d = {"kind": kind, "kicker": "", "title": [], "badge": "", "foot": "",
         "lead": [], "items": []}
    # 标题 / kicker / 徽标 / 页脚
    t = find_first(sec, tag="h1") or find_first(sec, tag="h2")
    if t is not None:
        d["title"] = inline(t)
    k = find_first(sec, cls="kicker")
    if k is not None:
        d["kicker"] = text_of(k).strip()
    b = find_first(sec, cls="badge")
    if b is not None:
        d["badge"] = text_of(b).strip()
    f = find_first(sec, cls="foot")
    if f is not None:
        d["foot"] = text_of(f).strip()
    l = find_first(sec, cls="lead")
    if l is not None:
        d["lead"] = inline(l)

    # 按出现顺序提取正文块
    for c in sec.children:
        if not isinstance(c, Node):
            continue
        if c.tag in ("h1", "h2") or has_class(c, "kicker") or has_class(c, "badge") \
           or has_class(c, "foot") or has_class(c, "lead"):
            continue
        if has_class(c, "sub"):
            d["items"].append(("sub", inline(c)))
        elif has_class(c, "quote"):
            who = find_first(c, cls="who")
            who_txt = text_of(who).strip() if who else ""
            q = []
            for p in inline(c):
                p2 = [seg for seg in p if seg[0].strip() != who_txt]
                if p2 and join_segs(p2).strip() and join_segs(p2).strip() != who_txt:
                    q.append(p2)
            d["items"].append(("quote", q, who_txt))
        elif has_class(c, "mirror"):
            d["items"].append(("mirror", inline(c)))
        elif has_class(c, "cards"):
            ncol = 3 if has_class(c, "c3") else 2
            cards = []
            for cc in c.children:
                if isinstance(cc, Node) and has_class(cc, "card"):
                    ic = text_of(find_first(cc, cls="ic")).strip() if find_first(cc, cls="ic") else ""
                    h = find_first(cc, tag="h3")
                    p = find_first(cc, tag="p")
                    cards.append((ic, text_of(h).strip() if h else "", inline(p) if p else []))
            d["items"].append(("cards", ncol, cards))
        elif has_class(c, "bullets"):
            bs = []
            for bb in c.children:
                if isinstance(bb, Node) and has_class(bb, "b"):
                    sm = find_first(bb, tag="small")
                    sub = text_of(sm).strip() if sm else None
                    segs = [p for p in inline(bb) if join_segs(p).strip()]
                    bs.append((segs, sub))
            d["items"].append(("bullets", bs))
        elif has_class(c, "stat"):
            boxes = []
            for bx in c.children:
                if isinstance(bx, Node) and has_class(bx, "box"):
                    num = find_first(bx, cls="num")
                    lbl = find_first(bx, cls="lbl")
                    ncls = num.attrs.get("class", "") if num else ""
                    boxes.append((flat(inline(num)).strip() if num else "",
                                  "up" in ncls, "down" in ncls,
                                  inline(lbl) if lbl else []))
            d["items"].append(("stat", boxes))
        elif has_class(c, "vs"):
            sides = []
            for sd in c.children:
                if isinstance(sd, Node) and has_class(sd, "side"):
                    h = find_first(sd, tag="h3")
                    lis = [inline(li) for li in all_li(sd)]
                    sides.append((text_of(h).strip() if h else "",
                                  "a" in sd.attrs.get("class", ""), lis))
            d["items"].append(("vs", sides))
        elif has_class(c, "chain"):
            links = []
            for lk in c.children:
                if isinstance(lk, Node) and has_class(lk, "link"):
                    tg = find_first(lk, cls="tag")
                    tt = find_first(lk, cls="tt")
                    title, desc = "", ""
                    if tt is not None:
                        paras = inline(tt)
                        # 标题 = 加粗片段；描述 = 非加粗片段
                        title = "".join(t for p in paras for t, b in p if b)
                        desc = "".join(t for p in paras for t, b in p if not b)
                    if tg is not None:
                        title = text_of(tg).strip() + "  " + title
                    links.append((title.strip(), re.sub(r"\s+", " ", desc).strip()))
            d["items"].append(("chain", links))
        elif has_class(c, "num3"):
            rows = []
            for rw in c.children:
                if isinstance(rw, Node) and has_class(rw, "row"):
                    n = find_first(rw, cls="n")
                    h4 = find_first(rw, tag="h4")
                    p = find_first(rw, tag="p")
                    rows.append((text_of(n).strip() if n else "", text_of(h4).strip() if h4 else "",
                                 inline(p) if p else []))
            d["items"].append(("num3", rows))
        else:
            sys.stderr.write(f"[warn] slide {idx + 1}: 未识别的块 <{c.tag} class=\"{c.attrs.get('class','')}\"> 已跳过\n")
    return d

# ---------------- 绘制 ----------------
def set_font(run, size=16, bold=False, color=INK, name=FONT):
    from pptx.oxml.ns import qn
    f = run.font
    f.size = Pt(size); f.bold = bold; f.color.rgb = color; f.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", name)

def add_rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp

def add_text(slide, x, y, w, h, paras, size=16, color=INK, bold=False,
             align=PP_ALIGN.LEFT, line=1.18, space_after=6, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(paras, list) and paras and isinstance(paras[0], tuple):
        paras = [paras]
    first = True
    for segs in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line
        p.space_after = Pt(space_after)
        for txt, boldflag in segs:
            r = p.add_run(); r.text = txt
            set_font(r, size=size, bold=bold or boldflag, color=color)
    return tb

def header(slide, kicker, title_paras, page):
    add_text(slide, 0.62, 0.38, 12.0, 0.4, [[(kicker, False)]], size=13, bold=True, color=CYAN, space_after=0)
    segs = [(t, True) for p in title_paras for t, _ in p]
    add_text(slide, 0.62, 0.72, 12.1, 0.95, [segs], size=30, bold=True, color=NAVY, space_after=0)
    add_rect(slide, 0.64, 1.66, 1.35, 0.075, fill=CYAN)
    add_rect(slide, 1.99, 1.66, 0.55, 0.075, fill=VIOLET)
    add_text(slide, 0.62, 7.08, 8.0, 0.3, [[("AI × 教育 · 思维外包", False)]], size=9, color=FAINT, space_after=0)
    add_text(slide, 12.0, 7.08, 0.72, 0.3, [[(f"{page:02d} / {TOTAL}", False)]], size=9,
             color=FAINT, align=PP_ALIGN.RIGHT, space_after=0)

def draw_item(slide, item, y, warn_idx):
    t = item[0]
    if t == "sub":
        add_text(slide, 0.72, y, 12.0, 0.45, [item[1][0] if item[1] else [( "", False)]],
                 size=15, color=DIM, space_after=0)
        return y + 0.45
    if t == "quote":
        _, paras, who = item
        add_rect(slide, 0.72, y + 0.08, 11.9, 0.02, fill=BORDER)
        yy = y + 0.22
        add_text(slide, 0.72, yy, 11.9, 0.5, paras[:1] or [[("", False)]],
                 size=16, color=INK, line=1.3, space_after=0)
        yy += 0.5
        if who:
            add_text(slide, 0.72, yy, 11.9, 0.35, [[(who, False)]], size=12, color=FAINT, space_after=0)
            yy += 0.4
        return yy
    if t == "mirror":
        _, paras = item
        add_rect(slide, 0.72, y, 11.9, 3.4, fill=PANEL, line=BORDER, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        add_text(slide, 1.2, y + 0.45, 11.0, 2.6, paras, size=19, color=INK, line=1.65, space_after=10)
        return y + 3.55
    if t == "cards":
        _, ncol, cards = item
        cw = (11.9 - (ncol - 1) * 0.2) / ncol
        ch = 2.25 if len(cards) > ncol else 1.9
        for i, (ic, h, p) in enumerate(cards):
            x = 0.72 + (i % ncol) * (cw + 0.2)
            yy = y + (i // ncol) * (ch + 0.28)
            add_rect(slide, x, yy, cw, ch, fill=PANEL, line=BORDER, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(slide, x + 0.3, yy + 0.25, cw - 0.6, 0.5, [[(ic, False)]], size=22, space_after=0)
            add_text(slide, x + 0.3, yy + 0.8, cw - 0.6, 0.45, [[(h, True)]], size=17, color=NAVY, space_after=0)
            add_text(slide, x + 0.3, yy + 1.3, cw - 0.6, 0.55, p, size=12.5, color=DIM, line=1.2, space_after=0)
        return y + ((len(cards) + ncol - 1) // ncol) * (ch + 0.28) - 0.28
    if t == "bullets":
        _, bs = item
        yy = y
        for segs, sub in bs:
            mark = [[("▸ ", True), *segs[0]]] if segs else []
            add_text(slide, 0.72, yy, 11.9, 0.6, mark, size=16, line=1.2, space_after=0)
            yy += 0.52
            if sub:
                add_text(slide, 1.06, yy, 11.6, 0.45, [[(sub, False)]], size=13, color=DIM, space_after=0)
                yy += 0.42
            yy += 0.14
        return yy
    if t == "stat":
        _, boxes = item
        n = len(boxes) or 1
        bw = (11.9 - (n - 1) * 0.28) / n
        for i, (num, up, down, lbl) in enumerate(boxes):
            x = 0.72 + i * (bw + 0.28)
            add_rect(slide, x, y, bw, 2.05, fill=PANEL, line=BORDER, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            col = GREEN if up else (PINK if down else NAVY)
            add_text(slide, x, y + 0.3, bw, 1.0, [[(num, True)]], size=40, color=col,
                     align=PP_ALIGN.CENTER, space_after=0)
            add_text(slide, x + 0.3, y + 1.35, bw - 0.6, 0.6, lbl, size=12.5, color=DIM,
                     align=PP_ALIGN.CENTER, line=1.2, space_after=0)
        return y + 2.2
    if t == "vs":
        _, sides = item
        left, right = sides[0] if len(sides) > 0 else ("", True, []), sides[1] if len(sides) > 1 else ("", False, [])
        for i, (h, is_a, lis) in enumerate([left, right]):
            x = 0.72 + i * 6.02
            if is_a:
                add_rect(slide, x, y, 5.8, 3.6, fill=PANEL, line=BORDER, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
                hcol, mark, scol = CYAN, "✔", CYAN
            else:
                add_rect(slide, x, y, 5.8, 3.6, fill=PANELP, line=BORDERP, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
                hcol, mark, scol = PINK, "✖", PINK
            add_text(slide, x + 0.3, y + 0.3, 5.2, 0.45, [[(h, True)]], size=19, color=hcol, space_after=0)
            yy = y + 1.0
            for li in lis:
                segs = [(mark + " ", True), *li[0]] if li else []
                add_text(slide, x + 0.3, yy, 5.2, 0.55, [segs] if segs else [], size=14, color=INK, line=1.2, space_after=0)
                yy += 0.56
        return y + 3.75
    if t == "chain":
        _, links = item
        yy = y
        for i, (tt, dd) in enumerate(links):
            add_rect(slide, 0.72, yy, 11.9, 0.62, fill=PANEL, line=BORDER, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(slide, 1.55, yy + 0.09, 3.6, 0.45, [[(tt, True)]], size=15, color=NAVY, space_after=0)
            add_text(slide, 5.4, yy + 0.14, 7.0, 0.45, [[(dd, False)]], size=12.5, color=DIM, space_after=0)
            yy += 0.62
            if i < len(links) - 1:
                add_text(slide, 6.4, yy - 0.02, 0.5, 0.3, [[("▼", False)]], size=11, color=VIOLET,
                         align=PP_ALIGN.CENTER, space_after=0)
                yy += 0.3
        return yy
    if t == "num3":
        _, rows = item
        yy = y
        for n, h, p in rows:
            add_rect(slide, 0.72, yy, 11.9, 1.35, fill=PANEL, line=BORDER, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_rect(slide, 1.0, yy + 0.36, 0.62, 0.62, fill=CYAN, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            add_text(slide, 1.0, yy + 0.44, 0.62, 0.5, [[(n, True)]], size=22, color=WHITE,
                     align=PP_ALIGN.CENTER, space_after=0)
            add_text(slide, 1.9, yy + 0.2, 10.5, 0.5, [[(h, True)]], size=17.5, color=NAVY, space_after=0)
            if p:
                add_text(slide, 1.9, yy + 0.72, 10.45, 0.55, p, size=13, color=DIM, line=1.25, space_after=0)
            yy += 1.5
        return yy
    sys.stderr.write(f"[warn] slide {warn_idx + 1}: 未知内容类型 {t} 已跳过\n")
    return y

# ---------------- 组装 ----------------
def build(html, out_path):
    global TOTAL
    parser = DeckParser()
    parser.feed(html)
    sections = []

    def walk(node):
        if isinstance(node, Node):
            if node.tag == "section" and "slide" in node.attrs.get("class", ""):
                sections.append(node)
            for c in node.children:
                walk(c)

    walk(parser.root)
    if not sections:
        sys.exit("错误：没有找到 <section class=\"slide\">…… 请确认 HTML 符合约定结构。")
    TOTAL = len(sections)
    prs = Presentation()
    prs.slide_width = Inches(SW); prs.slide_height = Inches(SH)
    BLANK = prs.slide_layouts[6]

    for idx, sec in enumerate(sections):
        d = extract_slide(sec, idx)
        s = prs.slides.add_slide(BLANK)
        if d["kind"] in ("cover", "final"):
            add_rect(s, 0, 0, SW, SH, fill=NAVY_BG)
            add_rect(s, 0, 0, SW, 0.09, fill=CYAN)
            add_rect(s, 0, 0.09, SW * 0.5, 0.09, fill=VIOLET)
            if d["badge"]:
                add_text(s, 0.9, 1.35, 11.5, 0.5, [[(d["badge"], True)]], size=15, color=CYAN, space_after=0)
            if d["kicker"] and not d["badge"]:
                add_text(s, 0.9, 0.5, 11.5, 0.4, [[(d["kicker"], True)]], size=13, color=CYAN, space_after=0)
            title = d["title"]
            title_segs = [[(t, True) for t, _ in p] for p in title] or [[("", False)]]
            add_text(s, 0.9, 2.1, 11.6, 2.1, title_segs, size=40, color=WHITE, line=1.15, space_after=10)
            yy = 4.4
            for p in d["lead"]:
                add_text(s, 0.9, yy, 11.0, 0.7, [p], size=17, color=LIGHT, line=1.4, space_after=0)
                yy += 0.62
            if d["foot"]:
                add_text(s, 0.9, 6.6, 11.5, 0.4, [[(d["foot"], False)]], size=11, color=FAINTB, space_after=0)
        else:
            header(s, d["kicker"], d["title"], idx + 1)
            y = 2.05
            for item in d["items"]:
                y = draw_item(s, item, y, idx)
        notes = f"由 {out_path.name} 的 slide {idx + 1} 自动生成"
        s.notes_slide.notes_text_frame.text = notes

    prs.save(str(out_path))
    print(f"完成：{len(sections)} 页 -> {out_path}")

def main():
    ap = argparse.ArgumentParser(description="结构化 HTML 幻灯片 -> PPTX")
    ap.add_argument("input", help="输入 HTML 文件")
    ap.add_argument("-o", "--output", default=None, help="输出 PPTX 路径（默认与输入同名）")
    args = ap.parse_args()
    src = Path(args.input)
    if not src.exists():
        sys.exit(f"错误：找不到文件 {src}")
    out = Path(args.output) if args.output else src.with_suffix(".pptx")
    raw = None
    for enc in ("utf-8", "gb18030"):
        try:
            raw = src.read_text(encoding=enc)
            if enc != "utf-8":
                sys.stderr.write(f"[warn] 文件不是 UTF-8，已按 {enc} 读取\n")
            break
        except UnicodeDecodeError:
            continue
    if raw is None:
        sys.exit("错误：无法以 UTF-8 或 GB18030 读取文件")
    build(raw, out)

if __name__ == "__main__":
    main()