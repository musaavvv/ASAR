#!/usr/bin/env python3
"""Qissa matnini (bob*.txt, oldin*.txt) A5 formatdagi Word (.docx) fayliga yig'adi.

Faqat standart kutubxona ishlatiladi. Ishga tushirish:
    python3 build_docx.py            # faylni yig'adi va bet sonini taxmin qiladi
    python3 build_docx.py --estimate # faqat taxmin
"""
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).parent
TITLE = "Ishqsiz qolishdan emas, sensiz qolishdan qo‘rqaman"
SUBTITLE = "Qissa"
OUT = HERE / "Ishqsiz_qolishdan_emas_sensiz_qolishdan_qorqaman.docx"

# Tartib: bobning fayl nomi. "oldin" parchalari tegishli boblardan keyin keladi.
ORDER = [
    "bob01", "bob02", "oldin1", "bob03", "bob04", "bob05", "bob06", "oldin2",
    "bob07", "bob08", "bob09", "oldin3", "bob10", "bob11", "bob12", "bob13",
]

# Sahifa (twips; 1 sm = 567 twips). A5: 14.8 x 21 sm.
PAGE_W, PAGE_H = 8391, 11906
MARGIN_TOP, MARGIN_BOTTOM = 1247, 1247
MARGIN_LEFT, MARGIN_RIGHT = 1134, 1134
FONT = "Times New Roman"
FONT_PT = 12
LINE_PITCH_PT = 16       # qator oralig'i (aniq, pt): betdagi qator soni barqaror bo'lsin
INDENT_TW = 425          # abzats burchagi (0.75 sm)


def fix_apostrophes(text: str) -> str:
    """ASCII ' belgisini o‘/g‘ (burilgan) yoki tutuq belgisi (’) ga almashtiradi."""
    def repl(m):
        prev = m.group(1)
        return prev + ("‘" if prev in "oOgG" else "’")
    return re.sub(r"([A-Za-zА-Яа-я])'", repl, text)


def load_chapter(name: str):
    raw = (HERE / f"{name}.txt").read_text(encoding="utf-8")
    raw = fix_apostrophes(raw)
    lines = raw.split("\n")
    head = lines[0].lstrip("# ").strip()
    body = [l for l in lines[1:]]
    paras, cur = [], []
    for l in body:
        if l.strip() == "":
            if cur:
                paras.append(" ".join(cur).strip())
                cur = []
        else:
            cur.append(l.strip())
    if cur:
        paras.append(" ".join(cur).strip())
    return head, paras


def split_heading(head: str):
    m = re.match(r"(\d+)\.\s+(.*)", head)
    if m:
        return f"{m.group(1)}-bob", m.group(2)
    m = re.match(r"Oldin\.\s+(.*)", head)
    if m:
        return "Oldin", m.group(1)
    return "", head


# ---------------------------------------------------------------- XML qismlari
def run(text, bold=False, italic=False, size=None, caps=False, spacing=None):
    rpr = ""
    if bold:
        rpr += "<w:b/>"
    if italic:
        rpr += "<w:i/>"
    if caps:
        rpr += "<w:caps/>"
    if spacing:
        rpr += f'<w:spacing w:val="{spacing}"/>'
    if size:
        rpr += f'<w:sz w:val="{size*2}"/><w:szCs w:val="{size*2}"/>'
    rpr = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(inner, style=None, jc=None, ind_first=None, before=None, after=None,
         keep_next=False, page_break_before=False, outline=None):
    ppr = ""
    if style:
        ppr += f'<w:pStyle w:val="{style}"/>'
    if keep_next:
        ppr += "<w:keepNext/>"
    if page_break_before:
        ppr += "<w:pageBreakBefore/>"
    if before is not None or after is not None:
        b = f' w:before="{before}"' if before is not None else ""
        a = f' w:after="{after}"' if after is not None else ""
        ppr += f"<w:spacing{b}{a}/>"
    if ind_first is not None:
        ppr += f'<w:ind w:firstLine="{ind_first}"/>'
    if jc:
        ppr += f'<w:jc w:val="{jc}"/>'
    if outline is not None:
        ppr += f'<w:outlineLvl w:val="{outline}"/>'
    ppr = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr}{inner}</w:p>"


def chapter_xml(name: str):
    head, paras = load_chapter(name)
    label, title = split_heading(head)
    out = []
    if label:
        out.append(para(run(label, size=11, caps=True, spacing=40), style="BobBelgisi",
                        keep_next=True, page_break_before=True))
        out.append(para(run(title), style="Heading1"))
    else:
        out.append(para(run(title), style="Heading1", page_break_before=True))
    first = True
    for p in paras:
        if re.fullmatch(r"(\\\*\s*){3}", p) or p.replace(" ", "") in ("***", "\\*\\*\\*"):
            out.append(para(run("*   *   *"), style="Ajratgich"))
            first = True  # sahna ajratgichdan keyin abzats burchaksiz
            continue
        if first:
            out.append(para(run(p), style="BirinchiAbzats"))
            first = False
        else:
            out.append(para(run(p)))
    return "".join(out), label, title


def build_document():
    body = []
    # Titul varag'i
    body.append(para("", before=2600, after=0))
    body.append(para(run(TITLE, bold=True, size=22), jc="center", ind_first=0, after=300,
                     before=0))
    body.append(para(run(SUBTITLE, italic=True, size=14), jc="center", ind_first=0, after=0))
    # Eslatma
    body.append(para(run("Eslatma", bold=True, size=12), jc="center", ind_first=0,
                     page_break_before=True, before=2600, after=240))
    note = ("Ushbu asar badiiy to‘qima. Qahramonlar, ularning taqdiri va voqealar muallif "
            "xayoli mahsulidir. Mo‘ynoq, Nukus, Toshkent kabi joy nomlari va Orol bo‘yida "
            "saksovul ekish ishlari real manzaralarga tayansa-da, kitobdagi shaxslar va "
            "hodisalar to‘qima.")
    body.append(para(run(note), jc="both", ind_first=0))

    chapters = [chapter_xml(n) for n in ORDER]

    # Mundarija
    body.append(para(run("Mundarija", bold=True, size=14), jc="center", ind_first=0,
                     page_break_before=True, after=360))
    for _, label, title in chapters:
        text = f"{label}. {title}" if label and label != "Oldin" else (
            f"Oldin. {title}" if label == "Oldin" else title)
        body.append(para(run(text), ind_first=0, after=120, jc="left"))

    for xml, _, _ in chapters:
        body.append(xml)

    sect = (
        '<w:sectPr>'
        '<w:footerReference w:type="default" r:id="rId3"/>'
        f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}"/>'
        f'<w:pgMar w:top="{MARGIN_TOP}" w:right="{MARGIN_RIGHT}" w:bottom="{MARGIN_BOTTOM}" '
        f'w:left="{MARGIN_LEFT}" w:header="567" w:footer="567" w:gutter="0"/>'
        '<w:titlePg/>'
        '</w:sectPr>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{"".join(body)}{sect}</w:body></w:document>'
    )


def build_styles():
    line = int(LINE_PITCH_PT * 20)
    half = FONT_PT * 2
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:cs="{FONT}" w:eastAsia="{FONT}"/>'
        f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>'
        '<w:lang w:val="uz-Latn-UZ" w:eastAsia="uz-Latn-UZ" w:bidi="ar-SA"/>'
        '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
        f'<w:pPr><w:widowControl/><w:spacing w:after="0" w:line="{line}" w:lineRule="exact"/>'
        f'<w:ind w:firstLine="{INDENT_TW}"/><w:jc w:val="both"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="BirinchiAbzats"><w:name w:val="Birinchi abzats"/>'
        '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:firstLine="0"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/><w:next w:val="BirinchiAbzats"/><w:qFormat/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="0" w:after="480" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:firstLine="0"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="34"/><w:szCs w:val="34"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="BobBelgisi"><w:name w:val="Bob belgisi"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="1400" w:after="120" '
        'w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Ajratgich"><w:name w:val="Sahna ajratgichi"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="240" '
        'w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>'
        '<w:rPr><w:sz w:val="20"/></w:rPr></w:style>'
        '</w:styles>'
    )


FOOTER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>'
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    '<w:r><w:t>1</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'
)

SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:defaultTabStop w:val="709"/><w:autoHyphenation w:val="false"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    '</w:settings>'
)

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    '</Types>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
)

DOC_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    '</Relationships>'
)


def core_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dc:title>{escape(TITLE)}</dc:title><dc:language>uz-Latn-UZ</dc:language>'
        '</cp:coreProperties>'
    )


APP_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
    '<Application>Microsoft Office Word</Application></Properties>'
)


# --------------------------------------------------------------- bet taxmini
AFM = {
    " ": 250, "a": 444, "b": 500, "c": 444, "d": 500, "e": 444, "f": 333, "g": 500, "h": 500,
    "i": 278, "j": 278, "k": 500, "l": 278, "m": 778, "n": 500, "o": 500, "p": 500, "q": 500,
    "r": 333, "s": 389, "t": 278, "u": 500, "v": 500, "w": 722, "x": 500, "y": 500, "z": 444,
    "A": 722, "B": 667, "C": 667, "D": 722, "E": 611, "F": 556, "G": 722, "H": 722, "I": 333,
    "J": 389, "K": 722, "L": 611, "M": 889, "N": 722, "O": 722, "P": 556, "Q": 722, "R": 667,
    "S": 556, "T": 611, "U": 722, "V": 722, "W": 944, "X": 722, "Y": 722, "Z": 611,
    ".": 250, ",": 250, "‘": 333, "’": 333, "«": 500, "»": 500, "—": 1000, "-": 333,
    ":": 278, ";": 278, "?": 444, "!": 333, "(": 333, ")": 333, "*": 500,
}


def text_width_pt(text: str) -> float:
    return sum(AFM.get(ch, 500 if ch.isdigit() else 500) for ch in text) * FONT_PT / 1000.0


def wrap_lines(text: str, width_pt: float, first_indent_pt: float) -> int:
    words = text.split(" ")
    lines, cur, empty = 1, first_indent_pt, True
    space = AFM[" "] * FONT_PT / 1000.0
    for w in words:
        ww = text_width_pt(w)
        need = ww if empty else ww + space
        if cur + need > width_pt and not empty:
            lines += 1
            cur = ww
        else:
            cur += need
        empty = False
    return lines


def estimate_pages():
    width_pt = (PAGE_W - MARGIN_LEFT - MARGIN_RIGHT) / 20.0
    height_pt = (PAGE_H - MARGIN_TOP - MARGIN_BOTTOM) / 20.0
    pitch = LINE_PITCH_PT
    lines_per_page = int(height_pt // pitch)
    indent = INDENT_TW / 20.0
    total_pages = 3 + 1  # titul, eslatma, mundarija (+ mundarija uzunligi hisobga olinadi)
    report = []
    body_words = 0
    for name in ORDER:
        head, paras = load_chapter(name)
        lines = 0
        first = True
        for p in paras:
            if p.replace(" ", "") in ("\\*\\*\\*", "***"):
                lines += 2.5
                first = True
                continue
            body_words += len(p.split())
            lines += wrap_lines(p, width_pt, 0 if first else indent)
            first = False
        heading_lines = 8.25  # bob belgisi (70 pt bo'shliq) + sarlavha + oraliqlar, 16 pt qatorda
        pages = -(-(lines + heading_lines) // lines_per_page)
        total_pages += pages
        report.append((name, int(lines), pages))
    return total_pages, lines_per_page, report, body_words


def main():
    total, lpp, report, words = estimate_pages()
    print(f"Qator/bet: {lpp}; so'zlar: {words}; taxminiy bet: {total}")
    for name, lines, pages in report:
        print(f"  {name}: {lines} qator, {pages} bet")
    if "--estimate" in sys.argv:
        return
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", build_document())
        z.writestr("word/styles.xml", build_styles())
        z.writestr("word/settings.xml", SETTINGS)
        z.writestr("word/footer1.xml", FOOTER)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", APP_XML)
    print("Yozildi:", OUT)


if __name__ == "__main__":
    main()
