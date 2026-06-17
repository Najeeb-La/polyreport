"""
PolyReport — Générateur de rapport technique
Polytechnique Montréal · Python pur (python-docx)
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import os

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────

POLY_RED  = RGBColor(0xCC, 0x00, 0x00)
POLY_DARK = RGBColor(0x1F, 0x1F, 0x1F)
POLY_GRAY = RGBColor(0x59, 0x59, 0x59)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)

FONTS_AVAILABLE = [
    "Cambria", "Times New Roman", "Arial", "Calibri",
    "Garamond", "Georgia", "Palatino Linotype", "Book Antiqua",
    "Helvetica", "Verdana", "Tahoma", "Century Schoolbook",
]
REF_STYLES = ["IEEE", "APA", "Chicago"]

MARGIN_CM      = 2.5
HEADER_DIST_CM = 1.25
FOOTER_DIST_CM = 1.25

# Chemin du logo — même dossier que ce script
_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(_DIR, "poly_logo.png")


# ─────────────────────────────────────────────
# HELPERS XML
# ─────────────────────────────────────────────

def _set_margins(section):
    section.top_margin      = Cm(MARGIN_CM)
    section.bottom_margin   = Cm(MARGIN_CM)
    section.left_margin     = Cm(MARGIN_CM)
    section.right_margin    = Cm(MARGIN_CM)
    section.header_distance = Cm(HEADER_DIST_CM)
    section.footer_distance = Cm(FOOTER_DIST_CM)


def _no_borders(table):
    tbl   = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    for existing in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(existing)
    tblBdr = OxmlElement('w:tblBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        tblBdr.append(el)
    tblPr.append(tblBdr)


def _shade_cell(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    for e in tcPr.findall(qn('w:shd')): tcPr.remove(e)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill_hex)
    tcPr.append(shd)


def _cell_borders(cell, color="AAAAAA", sz=4):
    tcPr  = cell._tc.get_or_add_tcPr()
    for e in tcPr.findall(qn('w:tcBorders')): tcPr.remove(e)
    tcBdr = OxmlElement('w:tcBorders')
    for side in ('top','left','bottom','right'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'),   'single')
        el.set(qn('w:sz'),    str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        tcBdr.append(el)
    tcPr.append(tcBdr)


def _nil_borders(cell):
    tcPr  = cell._tc.get_or_add_tcPr()
    for e in tcPr.findall(qn('w:tcBorders')): tcPr.remove(e)
    tcBdr = OxmlElement('w:tcBorders')
    for side in ('top','left','bottom','right','insideH','insideV'):
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'nil')
        tcBdr.append(el)
    tcPr.append(tcBdr)


def _run(para, text, font, size_pt, bold=False, italic=False, color=None):
    r = para.add_run(text)
    r.font.name  = font
    r.font.size  = Pt(size_pt)
    r.bold       = bold
    r.italic     = italic
    if color: r.font.color.rgb = color
    return r


def _para(doc, text="", font="Cambria", size=12,
          bold=False, italic=False, color=None,
          align=WD_ALIGN_PARAGRAPH.LEFT,
          sb=0, sa=6, ls=1.5):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before      = Pt(sb)
    p.paragraph_format.space_after       = Pt(sa)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE \
        if ls == 1.5 else WD_LINE_SPACING.MULTIPLE
    if ls != 1.5:
        p.paragraph_format.line_spacing = ls
    if text:
        _run(p, text, font, size, bold, italic, color)
    return p


def _page_break(doc):
    p  = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    r  = p.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r._r.append(br)
    return p


def _field(para, instr, font, size_pt):
    run = para.add_run()
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin')
    run._r.append(b)
    ins = OxmlElement('w:instrText'); ins.set(qn('xml:space'), 'preserve')
    ins.text = f' {instr} '; run._r.append(ins)
    s = OxmlElement('w:fldChar'); s.set(qn('w:fldCharType'), 'separate')
    run._r.append(s)
    ph = para.add_run("?"); ph.font.name = font; ph.font.size = Pt(size_pt)
    er = para.add_run()
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end')
    er._r.append(e)


def _toc_field(doc, instr, font, placeholder):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run()
    b   = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin')
    b.set(qn('w:dirty'), 'true'); run._r.append(b)
    ins = OxmlElement('w:instrText'); ins.set(qn('xml:space'), 'preserve')
    ins.text = f' {instr} '; run._r.append(ins)
    s = OxmlElement('w:fldChar'); s.set(qn('w:fldCharType'), 'separate')
    run._r.append(s)
    ph = p.add_run(placeholder)
    ph.font.name = font; ph.font.size = Pt(10); ph.italic = True
    ph.font.color.rgb = POLY_GRAY
    er = p.add_run()
    e  = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end')
    er._r.append(e)
    return p


def _add_section_break(doc, break_type='nextPage', pgnum_fmt='decimal', pgnum_start=1):
    p      = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    pPr    = p._p.get_or_add_pPr()
    sectPr = OxmlElement('w:sectPr')
    typeEl = OxmlElement('w:type'); typeEl.set(qn('w:val'), break_type)
    sectPr.append(typeEl)
    pgNum  = OxmlElement('w:pgNumType')
    pgNum.set(qn('w:fmt'),   pgnum_fmt)
    pgNum.set(qn('w:start'), str(pgnum_start))
    sectPr.append(pgNum)
    pPr.append(sectPr)
    return sectPr, p


# ─────────────────────────────────────────────
# STYLES POLY
# ─────────────────────────────────────────────

def _register_poly_styles(doc, font):
    from docx.enum.style import WD_STYLE_TYPE

    def _get(name):
        try: return doc.styles[name]
        except KeyError:
            return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)

    specs = {
        'CORPS_Titre1': dict(size=14, bold=True,         sb=18, sa=6),
        'CORPS_Titre2': dict(size=12, bold=True,         sb=12, sa=6),
        'CORPS_Titre3': dict(size=12, bold=True, it=True, sb=10, sa=6),
        'CORPS_Titre4': dict(size=12, it=True,           sb=10, sa=6, li=0.95),
        'Body Text':    dict(size=12,                    sb=0,  sa=6),
        'CORPS_REFERENCES': dict(size=12, single=True,   sb=0,  sa=4, li=0.75, fi=-0.75),
        'CORPS_ListePuceHérarchisée': dict(size=12,      sb=0,  sa=4, li=0.75),
        'LEGENDE_Figure':  dict(size=10, it=True,        sb=4,  sa=10),
        'LEGENDE_Tableau': dict(size=10, bold=True,      sb=10, sa=4),
    }

    for name, sp in specs.items():
        s  = _get(name)
        s.font.name   = font
        s.font.size   = Pt(sp['size'])
        s.font.bold   = sp.get('bold', False)
        s.font.italic = sp.get('it', False)
        pf = s.paragraph_format
        pf.space_before = Pt(sp.get('sb', 0))
        pf.space_after  = Pt(sp.get('sa', 6))
        if sp.get('single'):
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        else:
            pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        if sp.get('li'):  pf.left_indent        = Cm(sp['li'])
        if sp.get('fi'):  pf.first_line_indent  = Cm(sp['fi'])


# ─────────────────────────────────────────────
# EN-TÊTE — simple et professionnel
# ─────────────────────────────────────────────

def _build_header(section, config, font="Cambria"):
    """
    En-tête épuré : titre du rapport à gauche, numéro de page à droite,
    fin filet gris discret en dessous. Pas de bloc "Document/Révision"
    de type contrôle qualité — look rapport académique, pas gabarit ISO.
    """
    report_title = config.get("report_title", "").strip()

    header = section.header
    header.is_linked_to_previous = False
    for p in header.paragraphs: p.clear()

    tbl = header.add_table(1, 2, width=Inches(6.3))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _no_borders(tbl)
    tbl.cell(0,0).width = Inches(4.8)
    tbl.cell(0,1).width = Inches(1.5)
    _nil_borders(tbl.cell(0,0)); _nil_borders(tbl.cell(0,1))

    fs = 9

    p_left = tbl.cell(0,0).paragraphs[0]
    r_left = p_left.add_run(report_title[:70] if report_title else "")
    r_left.font.name = font; r_left.font.size = Pt(fs)
    r_left.font.color.rgb = POLY_GRAY

    p_right = tbl.cell(0,1).paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_pre = p_right.add_run("Page ")
    r_pre.font.name = font; r_pre.font.size = Pt(fs); r_pre.font.color.rgb = POLY_GRAY
    _field(p_right, 'PAGE \\* MERGEFORMAT', font, fs)
    r_mid = p_right.add_run(" / ")
    r_mid.font.name = font; r_mid.font.size = Pt(fs); r_mid.font.color.rgb = POLY_GRAY
    _field(p_right, 'SECTIONPAGES \\* MERGEFORMAT', font, fs)

    # Filet fin gris discret (pas rouge, pas épais)
    p_rule = header.add_paragraph()
    p_rule.paragraph_format.space_before = Pt(3)
    p_rule.paragraph_format.space_after  = Pt(0)
    pPr  = p_rule._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot  = OxmlElement('w:bottom')
    bot.set(qn('w:val'),   'single'); bot.set(qn('w:sz'), '4')
    bot.set(qn('w:space'), '1');      bot.set(qn('w:color'), 'CCCCCC')
    pBdr.append(bot); pPr.append(pBdr)


# ─────────────────────────────────────────────
# PAGE TITRE — style ton rapport MEC3520
# Logo centré en haut, texte centré, tableau seulement si équipe
# ─────────────────────────────────────────────

def build_title_page(doc, config):
    font         = config.get("font", "Cambria")
    course_code  = config.get("course_code", "").strip()
    course_name  = config.get("course_name", "").strip()
    group_number = config.get("group_number", "").strip()
    report_title = config.get("report_title", "Titre du rapport").strip()
    professors   = [p.strip() for p in config.get("professors", []) if str(p).strip()]
    members      = [m for m in config.get("members", [])
                    if (m.get("name","") if isinstance(m,dict) else str(m)).strip()]
    doc_date     = config.get("doc_date", "").strip()
    is_individual= config.get("is_individual", False)

    section = doc.sections[0]
    _set_margins(section)
    _build_header(section, config, font)

    sectPr = section._sectPr
    pgNum  = OxmlElement('w:pgNumType')
    pgNum.set(qn('w:fmt'), 'decimal'); pgNum.set(qn('w:start'), '1')
    sectPr.append(pgNum)

    # ── Logo Poly centré ─────────────────────
    if os.path.exists(LOGO_PATH):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(24)
        p_logo.paragraph_format.space_after  = Pt(24)
        run_logo = p_logo.add_run()
        run_logo.add_picture(LOGO_PATH, width=Inches(2.8))
    else:
        # Fallback texte si logo absent
        p_logo = _para(doc, "POLYTECHNIQUE MONTRÉAL", font=font, size=16,
                       bold=True, color=POLY_RED,
                       align=WD_ALIGN_PARAGRAPH.CENTER, sb=36, sa=24)

    # ── Cours + groupe ───────────────────────
    if course_code or course_name:
        txt = " — ".join(filter(None, [course_code, course_name]))
        p_c = doc.add_paragraph()
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_c.paragraph_format.space_before = Pt(0)
        p_c.paragraph_format.space_after  = Pt(0)
        r = p_c.add_run(txt)
        r.font.name = font; r.font.size = Pt(12); r.bold = True

    if group_number:
        p_g = doc.add_paragraph()
        p_g.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_g.paragraph_format.space_before = Pt(0)
        p_g.paragraph_format.space_after  = Pt(18)
        _run(p_g, f"Groupe-cours no {group_number}", font, 12)
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(18)

    # ── Titre du rapport ─────────────────────
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(0)
    p_t.paragraph_format.space_after  = Pt(24)
    _run(p_t, report_title, font, 13)

    # ── Présenté à ───────────────────────────
    if professors:
        p_pr = doc.add_paragraph()
        p_pr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_pr.paragraph_format.space_before = Pt(0)
        p_pr.paragraph_format.space_after  = Pt(6)
        _run(p_pr, "Présenté à " + ", ".join(professors), font, 12)

    # ── Auteur(s) ────────────────────────────
    if is_individual:
        p_fb = doc.add_paragraph()
        p_fb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_fb.paragraph_format.space_before = Pt(24)
        p_fb.paragraph_format.space_after  = Pt(6)
        _run(p_fb, "Fait par", font, 12)
        for m in (members if members else []):
            nom = m.get("name","") if isinstance(m,dict) else str(m)
            mat = m.get("matricule","") if isinstance(m,dict) else ""
            line = f"{nom}  {mat}".strip()
            p_m = doc.add_paragraph()
            p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_m.paragraph_format.space_before = Pt(0)
            p_m.paragraph_format.space_after  = Pt(2)
            _run(p_m, line, font, 12)
    else:
        # Travail d'équipe : "Fait par" + tableau signatures
        p_fb = doc.add_paragraph()
        p_fb.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_fb.paragraph_format.space_before = Pt(24)
        p_fb.paragraph_format.space_after  = Pt(12)
        _run(p_fb, "Fait par", font, 12)

        n = max(len(members), 4)
        tbl = doc.add_table(rows=n+1, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr_w = [Inches(2.0), Inches(1.0), Inches(3.0)]
        hdrs  = ["Nom, Prénom", "Matricule", "_______Signatures requises_______"]
        for j, (h, w) in enumerate(zip(hdrs, hdr_w)):
            c = tbl.cell(0, j); c.width = w
            _shade_cell(c, "1F1F1F"); _cell_borders(c, "1F1F1F")
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h); r.font.name=font; r.font.size=Pt(10)
            r.bold=True; r.font.color.rgb=WHITE

        for i in range(n):
            row = tbl.rows[i+1]
            m   = members[i] if i < len(members) else None
            nom = (m.get("name","") if isinstance(m,dict) else str(m)) if m else ""
            mat = (m.get("matricule","") if isinstance(m,dict) else "") if m else ""
            sig = "_" * 42
            for j, (val, w) in enumerate(zip([nom, mat, sig], hdr_w)):
                c = row.cells[j]; c.width=w; _cell_borders(c, "AAAAAA")
                p = c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(val); r.font.name=font; r.font.size=Pt(10)

    # ── Date ─────────────────────────────────
    if doc_date:
        p_d = doc.add_paragraph()
        p_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_d.paragraph_format.space_before = Pt(24)
        p_d.paragraph_format.space_after  = Pt(0)
        _run(p_d, f"Le {doc_date}", font, 12)
        p_m2 = doc.add_paragraph()
        p_m2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_m2.paragraph_format.space_before = Pt(0)
        p_m2.paragraph_format.space_after  = Pt(0)
        _run(p_m2, "à Montréal", font, 12)

    _page_break(doc)


# ─────────────────────────────────────────────
# RÉSUMÉ + ABSTRACT
# ─────────────────────────────────────────────

def build_resume_abstract(doc, config):
    font        = config.get("font", "Cambria")
    size        = config.get("font_size", 12)
    resume_fr   = config.get("resume_fr", "").strip()
    abstract_en = config.get("abstract_en", "").strip()
    keywords_fr = config.get("keywords_fr", "").strip()
    keywords_en = config.get("keywords_en", "").strip()

    def sec(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after  = Pt(6)
        r = p.add_run(title); r.font.name=font; r.font.size=Pt(size); r.bold=True

    sec("Résumé")
    try: p = doc.add_paragraph(style='Body Text')
    except: p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p, resume_fr if resume_fr else
         "Insérer le résumé ici (max. 1 page). Mise en contexte, objectif, résultats, prochaines actions.",
         font, size)
    if keywords_fr:
        try: pk = doc.add_paragraph(style='Body Text')
        except: pk = doc.add_paragraph()
        pk.paragraph_format.space_after = Pt(12)
        rb = pk.add_run("Mots clés : "); rb.font.name=font; rb.font.size=Pt(size); rb.bold=True
        pk.add_run(keywords_fr).font.name = font

    sec("Abstract")
    try: p = doc.add_paragraph(style='Body Text')
    except: p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _run(p, abstract_en if abstract_en else
         "Insert abstract here (max. 1 page). Context, objective, main results, next steps.",
         font, size)
    if keywords_en:
        try: pk = doc.add_paragraph(style='Body Text')
        except: pk = doc.add_paragraph()
        pk.paragraph_format.space_after = Pt(0)
        rb = pk.add_run("Keywords: "); rb.font.name=font; rb.font.size=Pt(size); rb.bold=True
        pk.add_run(keywords_en).font.name = font

    _page_break(doc)


# ─────────────────────────────────────────────
# TABLE DES MATIÈRES + LISTE FIGURES + LISTE TABLEAUX
# ─────────────────────────────────────────────

def build_toc_section(doc, config):
    font = config.get("font", "Cambria")

    def sec(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(8)
        r = p.add_run(title); r.font.name=font; r.font.size=Pt(14); r.bold=True

    sec("Table des matières")
    _toc_field(doc, 'TOC \\o "1-4" \\h \\z \\u', font,
               "[ Ctrl+A → F9 dans Word pour mettre à jour ]")
    _page_break(doc)

    sec("Liste des tableaux")
    _toc_field(doc, 'TOC \\h \\z \\c "Tableau"', font,
               "[ Liste des tableaux — mise à jour auto (F9) ]")
    _page_break(doc)

    sec("Liste des figures")
    _toc_field(doc, 'TOC \\h \\z \\c "Figure"', font,
               "[ Liste des figures — mise à jour auto (F9) ]")
    _page_break(doc)


# ─────────────────────────────────────────────
# SECTIONS — placeholders structurés
# L'utilisateur écrit son texte dans Word directement
# ─────────────────────────────────────────────

def _heading(doc, text, level, font, number=""):
    style_map = {1:'CORPS_Titre1', 2:'CORPS_Titre2', 3:'CORPS_Titre3', 4:'CORPS_Titre4'}
    try: p = doc.add_paragraph(style=style_map.get(level,'CORPS_Titre1'))
    except: p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(18 if level==1 else 12)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    full = f"{number}\t{text}" if number else text
    r = p.add_run(full); r.font.name=font
    r.font.size = Pt(14 if level==1 else 12)
    r.bold   = level <= 2
    r.italic = level >= 3
    return p


def _placeholder_text(doc, font, size, text):
    """Paragraphe gris italique — zone à remplir par l'utilisateur."""
    try: p = doc.add_paragraph(style='Body Text')
    except: p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text); r.font.name=font; r.font.size=Pt(size)
    r.italic=True; r.font.color.rgb=POLY_GRAY
    return p


def _insert_figure_placeholder(doc, font, size, fig_num, caption=""):
    """Zone réservée figure + légende."""
    # Cadre gris pour l'espace de la figure
    tbl = doc.add_table(1, 1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.cell(0, 0)
    c.width = Inches(5.0)
    _shade_cell(c, "F2F2F2")
    _cell_borders(c, "CCCCCC", sz=4)
    cp = c.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(24)
    cp.paragraph_format.space_after  = Pt(24)
    r = cp.add_run(f"[ Insérer Figure {fig_num} ici ]")
    r.font.name=font; r.font.size=Pt(10); r.italic=True; r.font.color.rgb=POLY_GRAY

    # Légende avec style LEGENDE_Figure (capturée par la liste des figures)
    cap_text = caption if caption else f"Description de la figure {fig_num}."
    try: p_cap = doc.add_paragraph(style='LEGENDE_Figure')
    except: p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(4)
    p_cap.paragraph_format.space_after  = Pt(12)

    # Champ SEQ Figure pour numérotation automatique
    r_pre = p_cap.add_run("Figure "); r_pre.font.name=font; r_pre.font.size=Pt(10); r_pre.italic=True
    seq_run = p_cap.add_run()
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin'); seq_run._r.append(b)
    ins = OxmlElement('w:instrText'); ins.set(qn('xml:space'),'preserve')
    ins.text = ' SEQ Figure \\* ARABIC '; seq_run._r.append(ins)
    s = OxmlElement('w:fldChar'); s.set(qn('w:fldCharType'),'separate'); seq_run._r.append(s)
    ph = p_cap.add_run(str(fig_num)); ph.font.name=font; ph.font.size=Pt(10); ph.italic=True
    er = p_cap.add_run()
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'),'end'); er._r.append(e)
    r_cap = p_cap.add_run(f" — {cap_text}")
    r_cap.font.name=font; r_cap.font.size=Pt(10); r_cap.italic=True
    return p_cap


def _insert_table_placeholder(doc, font, size, tbl_num, caption="", rows=3, cols=3):
    """Tableau placeholder avec en-tête + légende."""
    # Légende AVANT le tableau (convention Poly)
    cap_text = caption if caption else f"Description du tableau {tbl_num}."
    try: p_cap = doc.add_paragraph(style='LEGENDE_Tableau')
    except: p_cap = doc.add_paragraph()
    p_cap.paragraph_format.space_before = Pt(10)
    p_cap.paragraph_format.space_after  = Pt(4)

    r_pre = p_cap.add_run("Tableau "); r_pre.font.name=font; r_pre.font.size=Pt(10); r_pre.bold=True
    seq_run = p_cap.add_run()
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'),'begin'); seq_run._r.append(b)
    ins = OxmlElement('w:instrText'); ins.set(qn('xml:space'),'preserve')
    ins.text = ' SEQ Tableau \\* ARABIC '; seq_run._r.append(ins)
    s = OxmlElement('w:fldChar'); s.set(qn('w:fldCharType'),'separate'); seq_run._r.append(s)
    ph = p_cap.add_run(str(tbl_num)); ph.font.name=font; ph.font.size=Pt(10); ph.bold=True
    er = p_cap.add_run()
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'),'end'); er._r.append(e)
    r_cap = p_cap.add_run(f" : {cap_text}")
    r_cap.font.name=font; r_cap.font.size=Pt(10); r_cap.bold=True

    # Tableau avec en-tête gris
    data_tbl = doc.add_table(rows=rows, cols=cols)
    data_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j in range(cols):
        c = data_tbl.cell(0, j)
        _shade_cell(c, "1F1F1F"); _cell_borders(c, "1F1F1F")
        p = c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"En-tête {j+1}"); r.font.name=font; r.font.size=Pt(10)
        r.bold=True; r.font.color.rgb=WHITE
    for i in range(1, rows):
        for j in range(cols):
            c = data_tbl.cell(i,j); _cell_borders(c,"AAAAAA")
            if i % 2 == 0: _shade_cell(c, "F7F7F7")
            p = c.paragraphs[0]
            r = p.add_run("—"); r.font.name=font; r.font.size=Pt(10)

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(10)
    return p_after


def build_body_sections(doc, config):
    font  = config.get("font", "Cambria")
    size  = config.get("font_size", 12)
    secs  = config.get("sections", [])

    if not secs:
        secs = [
            {"title":"Introduction","level":1,"figures":0,"tables":0},
            {"title":"Mise en contexte","level":2,"figures":0,"tables":0},
            {"title":"Objectifs spécifiques","level":2,"figures":0,"tables":0},
            {"title":"Développement","level":1,"figures":1,"tables":1},
            {"title":"Résultats","level":1,"figures":1,"tables":0},
            {"title":"Conclusion","level":1,"figures":0,"tables":0},
        ]

    counters  = [0,0,0,0]
    fig_count = 0
    tbl_count = 0

    for sec in secs:
        lvl   = sec.get("level",1) - 1
        title = sec.get("title","")
        n_fig = int(sec.get("figures", 0))
        n_tbl = int(sec.get("tables",  0))
        fig_captions = sec.get("fig_captions", [])
        tbl_captions = sec.get("tbl_captions", [])

        counters[lvl] += 1
        for i in range(lvl+1, 4): counters[i] = 0
        num = ".".join(str(counters[i]) for i in range(lvl+1))

        _heading(doc, title, lvl+1, font, number=num)

        # Placeholder texte
        _placeholder_text(doc, font, size,
            "[ Rédiger le contenu de cette section directement dans Word. "
            "Supprimer ce texte gris avant de remettre le rapport. ]")

        # Figures
        for fi in range(n_fig):
            fig_count += 1
            cap = fig_captions[fi] if fi < len(fig_captions) else ""
            _insert_figure_placeholder(doc, font, size, fig_count, cap)

        # Tableaux
        for ti in range(n_tbl):
            tbl_count += 1
            cap = tbl_captions[ti] if ti < len(tbl_captions) else ""
            _insert_table_placeholder(doc, font, size, tbl_count, cap)


# ─────────────────────────────────────────────
# RÉFÉRENCES
# ─────────────────────────────────────────────

def build_references(doc, config):
    font       = config.get("font", "Cambria")
    size       = config.get("font_size", 12)
    ref_style  = config.get("ref_style", "IEEE")
    references = config.get("references", [])  # list of dicts {formatted, url}

    _heading(doc, "Références", 1, font)

    if not references:
        _placeholder_text(doc, font, size,
            "[ Aucune référence. Ajouter les sources ici. ]")
        return

    for i, ref in enumerate(references, 1):
        formatted = ref.get("formatted","") if isinstance(ref,dict) else str(ref)
        url       = ref.get("url","")       if isinstance(ref,dict) else ""
        if not formatted.strip(): continue

        try: p = doc.add_paragraph(style='CORPS_REFERENCES')
        except:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent      = Cm(0.75)
            p.paragraph_format.first_line_indent = Cm(-0.75)

        prefix = f"[{i}]\t" if ref_style=="IEEE" else ""
        r = p.add_run(prefix + formatted)
        r.font.name=font; r.font.size=Pt(size)

        # URL en hyperlien si présente
        if url:
            p2 = doc.add_paragraph()
            p2.paragraph_format.left_indent = Cm(0.75)
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after  = Pt(4)
            ru = p2.add_run(f"    {url}")
            ru.font.name=font; ru.font.size=Pt(size-1)
            ru.font.color.rgb = RGBColor(0,70,180)


# ─────────────────────────────────────────────
# ANNEXES — même logique que les sections
# ─────────────────────────────────────────────

def build_annexes(doc, config):
    font    = config.get("font", "Cambria")
    size    = config.get("font_size", 12)
    annexes = config.get("annexes", [])
    if not annexes: return

    letters   = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    fig_count = config.get("_fig_count_start", 0)
    tbl_count = config.get("_tbl_count_start", 0)

    for i, annexe in enumerate(annexes):
        letter      = letters[i] if i < len(letters) else str(i+1)
        title       = annexe.get("title", f"Titre de l'annexe {letter}")
        n_fig       = int(annexe.get("figures", 0))
        n_tbl       = int(annexe.get("tables", 0))
        fig_caps    = annexe.get("fig_captions", [])
        tbl_caps    = annexe.get("tbl_captions", [])

        _add_section_break(doc, 'nextPage', 'decimal', 1)

        p_ann = doc.add_paragraph()
        p_ann.paragraph_format.space_before = Pt(0)
        p_ann.paragraph_format.space_after  = Pt(12)
        r = p_ann.add_run(f"Annexe {letter} : {title}")
        r.font.name=font; r.font.size=Pt(16); r.bold=True

        _placeholder_text(doc, font, size,
            "[ Rédiger le contenu de cette annexe directement dans Word. "
            "Supprimer ce texte gris avant de remettre. ]")

        for fi in range(n_fig):
            fig_count += 1
            cap = fig_caps[fi] if fi < len(fig_caps) else ""
            _insert_figure_placeholder(doc, font, size, fig_count, cap)

        for ti in range(n_tbl):
            tbl_count += 1
            cap = tbl_caps[ti] if ti < len(tbl_caps) else ""
            _insert_table_placeholder(doc, font, size, tbl_count, cap)


# ─────────────────────────────────────────────
# FONCTION PRINCIPALE
# ─────────────────────────────────────────────

def generate_poly_report(config: dict) -> BytesIO:
    doc = Document()
    font_name = config.get("font", "Cambria")
    doc.styles['Normal'].font.name = font_name
    doc.styles['Normal'].font.size = Pt(config.get("font_size", 12))
    _register_poly_styles(doc, font_name)

    build_title_page(doc, config)

    _add_section_break(doc, 'nextPage', 'lowerRoman', 1)
    _set_margins(doc.sections[-1])
    _build_header(doc.sections[-1], config, font_name)
    build_resume_abstract(doc, config)
    build_toc_section(doc, config)

    _add_section_break(doc, 'nextPage', 'decimal', 1)
    _set_margins(doc.sections[-1])
    _build_header(doc.sections[-1], config, font_name)
    build_body_sections(doc, config)
    build_references(doc, config)

    build_annexes(doc, config)

    for s in doc.sections: _set_margins(s)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


if __name__ == "__main__":
    cfg = {
        "font":"Cambria","font_size":12,
        "doc_number":"","doc_date":"9 avril 2026",
        "course_code":"MEC3520","course_name":"Industrialisation des produits",
        "group_number":"02","report_title":"Recyclage à l'état solide des copeaux métalliques",
        "semester":"Hiver 2026","professors":["Marwan Azzi"],
        "members":[{"name":"Najeeb Lababidi","matricule":"2369834"},
                   {"name":"Victor Béland","matricule":"2381714"}],
        "is_individual":False,
        "resume_fr":"Ce rapport présente une analyse du procédé SSR.",
        "abstract_en":"This report presents an analysis of the SSR process.",
        "keywords_fr":"SSR, recyclage, aluminium","keywords_en":"SSR, recycling, aluminium",
        "ref_style":"APA",
        "sections":[
            {"title":"Introduction","level":1,"figures":0,"tables":0},
            {"title":"Description du procédé","level":1,"figures":1,"tables":0,
             "fig_captions":["Schéma du procédé SSR (Adapté de Krolo et al., 2019)."]},
            {"title":"Principe général","level":2,"figures":0,"tables":0},
            {"title":"Résultats","level":1,"figures":0,"tables":1,
             "tbl_captions":["Comparaison des procédés concurrents."]},
            {"title":"Conclusion","level":1,"figures":0,"tables":0},
        ],
        "references":[
            {"formatted":"Altharan, Y. M. et al. (2024). A review on solid-state recycling of aluminum machining chips. Heliyon, 10(14), e34433.",
             "url":"https://doi.org/10.1016/j.heliyon.2024.e34433"},
            {"formatted":"Bocchi, S., D'Urso, G., & Giardini, C. (2025). Enhancing sustainability in aluminum recycling. Journal of Materials Engineering and Performance.",
             "url":"https://doi.org/10.1007/s11665-025-11434-9"},
        ],
        "annexes":[
            {"title":"Données thermodynamiques","figures":1,"tables":1,
             "fig_captions":["Profil de température mesuré."],
             "tbl_captions":["Propriétés mécaniques mesurées."]},
        ],
    }
    buf = generate_poly_report(cfg)
    with open("/home/claude/rapport_test.docx","wb") as f: f.write(buf.read())
    print("✅ Généré: rapport_test.docx")
