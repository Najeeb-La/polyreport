"""
PolyReport — Générateur de rapport technique
Polytechnique Montréal · Par Najeeb Lababidi · 100% gratuit, aucune clé API
"""

import streamlit as st
from generate_report import generate_poly_report, FONTS_AVAILABLE, REF_STYLES
from datetime import date
import requests, os

# ─────────────────────────────────────────────
# CONFIG PAGE
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="PolyReport — Générateur de rapports techniques",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────

st.markdown("""
<style>
  /* ── Hero header ── */
  .hero-wrap {
    background: linear-gradient(120deg,#A60000 0%,#7A0000 100%);
    border-radius: 14px;
    padding: 1.6rem 2rem;
    margin-bottom: 1.6rem;
    display: flex;
    align-items: center;
    gap: 1.6rem;
    box-shadow: 0 2px 14px rgba(166,0,0,0.18);
  }
  .hero-logo {
    background: white;
    border-radius: 10px;
    padding: 8px 14px;
    display: flex;
    align-items: center;
    flex-shrink: 0;
  }
  .hero-text h1 {
    color: white; margin: 0; font-size: 1.5rem; font-weight: 700;
    letter-spacing: -.01em;
  }
  .hero-text .sub {
    color: rgba(255,255,255,.85); font-size: .88rem; margin: .35rem 0 0;
    font-weight: 500;
  }
  .hero-text .credit {
    color: rgba(255,255,255,.62); font-size: .76rem; margin: .55rem 0 0;
    line-height: 1.5;
  }

  /* ── Badges ── */
  .badge {
    background:#A60000; color:white;
    padding:.16rem .58rem; border-radius:5px;
    font-size:.7rem; font-weight:700; letter-spacing:.02em;
    margin-right:.45rem; text-transform: uppercase;
  }
  .badge-gray {
    background:#444; color:white;
    padding:.16rem .58rem; border-radius:5px;
    font-size:.7rem; font-weight:700; letter-spacing:.02em;
    margin-right:.45rem; text-transform: uppercase;
  }

  /* ── Info boxes ── */
  .info-box {
    background:#F7F7F8; border-left:3px solid #A60000;
    padding:.6rem .9rem; border-radius:0 8px 8px 0;
    margin:.5rem 0; font-size:.83rem; color:#444; line-height:1.5;
  }
  .warn-box {
    background:#FFF7E6; border-left:3px solid #E8A33D;
    padding:.6rem .9rem; border-radius:0 8px 8px 0;
    margin:.5rem 0; font-size:.83rem; color:#5a4a2a; line-height:1.5;
  }
  .success-box {
    background:#EAF7ED; border-left:3px solid #2E9E50;
    padding:.6rem .9rem; border-radius:0 8px 8px 0;
    margin:.5rem 0; font-size:.83rem; color:#1d6b35; line-height:1.5;
  }

  /* ── Section divider label ── */
  .section-label {
    font-size: .95rem; font-weight: 700; color: #1a1a1a;
    margin: .2rem 0 .6rem;
  }

  /* ── Tabs ── */
  div[data-testid="stTabs"] button {
    font-size: .92rem;
  }
  div[data-testid="stTabs"] button[aria-selected="true"] {
    border-bottom: 3px solid #A60000 !important;
    color: #A60000 !important; font-weight: 700;
  }

  /* ── Sidebar ── */
  section[data-testid="stSidebar"] {
    background: #FAFAFA;
    border-right: 1px solid #ECECEC;
  }
  section[data-testid="stSidebar"] h3 {
    font-size: 1rem; color: #1a1a1a;
  }

  /* ── Expanders : plus de respiration et bordure subtile ── */
  div[data-testid="stExpander"] {
    border: 1px solid #E8E8E8 !important;
    border-radius: 8px !important;
    margin-bottom: .5rem;
  }

  /* ── Primary button ── */
  button[kind="primary"] {
    background: #A60000 !important;
    border: none !important;
    font-weight: 700 !important;
    border-radius: 8px !important;
  }
  button[kind="primary"]:hover {
    background: #8a0000 !important;
  }

  /* ── Footer credit ── */
  .footer-credit {
    font-size: .74rem; color: #999; text-align: center; line-height: 1.6;
  }
  .footer-credit a { color: #A60000; text-decoration: none; font-weight: 600; }
  .footer-credit a:hover { text-decoration: underline; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────

_DIR  = os.path.dirname(os.path.abspath(__file__))
_LOGO = os.path.join(_DIR, "poly_logo.png")

import base64
_logo_b64 = ""
if os.path.exists(_LOGO):
    with open(_LOGO, "rb") as _f:
        _logo_b64 = base64.b64encode(_f.read()).decode()

st.markdown(f"""
<div class="hero-wrap">
  {f'<div class="hero-logo"><img src="data:image/png;base64,{_logo_b64}" width="130"></div>' if _logo_b64 else ''}
  <div class="hero-text">
    <h1>📄 PolyReport — Générateur de rapports techniques</h1>
    <div class="sub">Polytechnique Montréal · Rapports techniques académiques · 100&nbsp;% gratuit</div>
    <div class="credit">
      Développé par <strong>Najeeb Lababidi</strong> · Étudiant en génie aérospatial, Polytechnique Montréal<br>
      Génère le squelette Word complet — tu écris ton texte directement dans le .docx
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ Mise en forme")
    st.markdown("---")
    font      = st.selectbox("Police", FONTS_AVAILABLE, index=0,
                              help="Cambria est recommandée pour les rapports académiques")
    font_size = st.slider("Taille du corps (pt)", 10, 14, 12)
    ref_style = st.selectbox("Style de références", REF_STYLES, index=0,
                              help="IEEE recommandé en génie")
    st.markdown("---")
    st.markdown("### 💡 Comment utiliser")
    st.markdown("""
    <div class="info-box">
    <b>1.</b> Remplis le formulaire<br>
    <b>2.</b> Clique « Générer »<br>
    <b>3.</b> Ouvre le .docx dans Word<br>
    <b>4.</b> Ctrl+A → F9 (TDM auto)<br>
    <b>5.</b> Écris dans les zones grises
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("""
    <div class="footer-credit">
    Outil gratuit pour les étudiants de Polytechnique Montréal<br>
    © Najeeb Lababidi · Génie aérospatial<br>
    <a href="mailto:najeeb.lababidi@etud.polymtl.ca">Signaler un problème</a>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ONGLETS
# ─────────────────────────────────────────────

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Identification",
    "📝 Résumé & Abstract",
    "📖 Sections",
    "📚 Références",
    "📎 Annexes",
])

# ══════════════════════════════════════════════
# TAB 1 — IDENTIFICATION
# ══════════════════════════════════════════════
with tab1:
    st.markdown('<span class="badge">Document</span><span class="section-label">Informations générales</span>',
                unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        doc_date_raw = st.date_input("Date du document", value=date.today())
        MONTHS_FR = {1:"janvier",2:"février",3:"mars",4:"avril",5:"mai",6:"juin",
                     7:"juillet",8:"août",9:"septembre",10:"octobre",
                     11:"novembre",12:"décembre"}
        doc_date = f"{doc_date_raw.day} {MONTHS_FR[doc_date_raw.month]} {doc_date_raw.year}"

        sc1, sc2 = st.columns(2)
        with sc1:
            season = st.selectbox("Session", ["Hiver","Été","Automne"], index=0,
                                   label_visibility="visible")
        with sc2:
            year_sess = st.number_input("Année", min_value=2020, max_value=2035,
                                         value=date.today().year, step=1,
                                         label_visibility="visible")
        semester = f"{season} {year_sess}"
    with c2:
        group_number = st.text_input("Groupe-cours n°", value="",
                                      placeholder="ex : 02",
                                      help="Laisser vide si non applicable")

    st.markdown("---")
    st.markdown('<span class="badge">Cours</span><span class="section-label">Cours associé</span>',
                unsafe_allow_html=True)
    cc1, cc2 = st.columns(2)
    with cc1:
        course_code = st.text_input("Sigle", value="", placeholder="ex : MEC 3520")
    with cc2:
        course_name = st.text_input("Titre du cours", value="",
                                     placeholder="ex : Industrialisation des produits")

    st.markdown("---")
    st.markdown('<span class="badge">Rapport</span><span class="section-label">Titre du rapport</span>',
                unsafe_allow_html=True)
    report_title = st.text_input("Titre", value="",
                                  placeholder="Titre de votre rapport technique")

    st.markdown("---")
    st.markdown('<span class="badge">Équipe</span><span class="section-label">Encadrement et auteur(s)</span>',
                unsafe_allow_html=True)

    professors_raw = st.text_input(
        "Présenté à (professeur(s), séparés par virgule)",
        value="", placeholder="ex : Marwan Azzi  ou  Prof. Morin, Dr. Ahmadi",
        help="Laisser vide si non applicable"
    )
    professors = [p.strip() for p in professors_raw.split(",") if p.strip()]

    is_individual = st.toggle("🎓 Travail individuel — pas de tableau de signatures",
                               value=False)

    if is_individual:
        st.markdown('<span class="badge-gray">Étudiant</span>'
                    '<span class="section-label">Vos informations</span>',
                    unsafe_allow_html=True)
        st.markdown("""<div class="info-box">
        Format : <code>Nom Prénom | Matricule</code>
        </div>""", unsafe_allow_html=True)
        m_raw = st.text_input("Nom et matricule", value="",
                               placeholder="Nom Prénom | 2XXXXXX",
                               label_visibility="collapsed")
        if "|" in m_raw:
            parts = m_raw.split("|", 1)
            members = [{"name": parts[0].strip(), "matricule": parts[1].strip()}]
        elif m_raw.strip():
            members = [{"name": m_raw.strip(), "matricule": ""}]
        else:
            members = []
    else:
        st.markdown('<span class="badge-gray">Équipe</span>'
                    '<span class="section-label">Membres</span>',
                    unsafe_allow_html=True)
        st.markdown("""<div class="info-box">
        Format : <code>Nom Prénom | Matricule</code> — un membre par ligne.
        </div>""", unsafe_allow_html=True)
        members_raw = st.text_area("Membres", label_visibility="collapsed",
                                    height=120, value="",
                                    placeholder="Nom Prénom | 2XXXXXX\nNom Prénom | 2XXXXXX")
        members = []
        for line in members_raw.strip().split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                p2 = line.split("|", 1)
                members.append({"name": p2[0].strip(), "matricule": p2[1].strip()})
            else:
                members.append({"name": line, "matricule": ""})

# ══════════════════════════════════════════════
# TAB 2 — RÉSUMÉ & ABSTRACT
# ══════════════════════════════════════════════
with tab2:
    st.markdown("""<div class="info-box">
    Le résumé <b>ne dépasse pas une page</b>. Contenu : mise en contexte, objectif,
    résultats principaux, prochaines actions.<br>
    L'abstract est une traduction fidèle du résumé.
    </div>""", unsafe_allow_html=True)

    st.markdown('<span class="badge">FR</span><span class="section-label">Résumé en français</span>',
                unsafe_allow_html=True)
    resume_fr   = st.text_area("Résumé", label_visibility="collapsed", height=170,
                                placeholder="Coller ou écrire votre résumé ici…")
    keywords_fr = st.text_input("Mots clés FR (séparés par virgules)",
                                 placeholder="ex : turbofan, propulsion, efficacité")

    st.markdown("---")
    st.markdown('<span class="badge">EN</span><span class="section-label">Abstract</span>',
                unsafe_allow_html=True)

    # ── Traduction gratuite (Google Translate, sans clé API) ──
    # Pattern : clé de widget versionnée. Streamlit donne priorité à la valeur
    # cachée du widget (liée à sa clé) sur le paramètre value= au rerun suivant.
    # En changeant la clé après chaque traduction, on force un widget "neuf"
    # qui prend bien la nouvelle valeur.
    if "abstract_value" not in st.session_state:
        st.session_state["abstract_value"] = ""
    if "abstract_ver" not in st.session_state:
        st.session_state["abstract_ver"] = 0

    col_abs, col_btn = st.columns([4, 1])
    with col_abs:
        abstract_en = st.text_area(
            "Abstract", label_visibility="collapsed", height=170,
            placeholder="Paste or write your abstract here…",
            value=st.session_state["abstract_value"],
            key=f"abstract_widget_{st.session_state['abstract_ver']}",
        )
        st.session_state["abstract_value"] = abstract_en
    with col_btn:
        st.markdown("<br><br>", unsafe_allow_html=True)
        translate_clicked = st.button("🌐 Traduire\nle résumé",
                                       help="Traduction gratuite FR → EN via Google Translate",
                                       use_container_width=True)

    if translate_clicked:
        if not resume_fr.strip():
            st.warning("Écris d'abord ton résumé en français.")
        else:
            with st.spinner("Traduction en cours…"):
                translated   = None
                error_detail = None
                try:
                    from deep_translator import GoogleTranslator
                    text = resume_fr.strip()
                    MAX  = 4500   # limite Google Translate par requête

                    if len(text) <= MAX:
                        translated = GoogleTranslator(
                            source='fr', target='en').translate(text)
                    else:
                        parts = [p for p in text.split("\n\n") if p.strip()]
                        translated = "\n\n".join(
                            GoogleTranslator(source='fr', target='en').translate(p)
                            for p in parts
                        )
                except Exception as e:
                    error_detail = str(e)

            if translated and translated.strip():
                st.session_state["abstract_value"] = translated.strip()
                st.session_state["abstract_ver"]   += 1   # force un widget neuf
                st.session_state["_translation_success"] = True
                st.rerun()
            else:
                st.markdown(f"""<div class="warn-box">
                ⚠️ Traduction automatique indisponible pour le moment
                {f"({error_detail})" if error_detail else ""}.<br>
                <b>Alternative gratuite en 30 secondes :</b>
                copie ton résumé → colle sur
                <a href="https://www.deepl.com/translator" target="_blank">
                <b>DeepL.com</b></a> (gratuit, excellente qualité technique)
                → colle le résultat dans le champ Abstract ci-dessus.
                </div>""", unsafe_allow_html=True)


    if st.session_state.pop("_translation_success", False):
        st.success("✅ Traduction générée — vérifie les termes techniques avant de remettre.")

    keywords_en = st.text_input("Keywords EN (comma-separated)",
                                 placeholder="ex: turbofan, propulsion, efficiency")

# ══════════════════════════════════════════════
# TAB 3 — SECTIONS
# ══════════════════════════════════════════════
with tab3:
    st.markdown("""<div class="info-box">
    Configure la <b>structure</b> du rapport (titres, niveaux, figures, tableaux).<br>
    Le .docx généré contiendra des <b>zones grises</b> avec les styles Word corrects
    — tu écris ton texte directement dans Word.<br>
    <b>Ctrl+A → F9</b> met à jour la TDM, liste des figures et liste des tableaux automatiquement.
    </div>""", unsafe_allow_html=True)

    DEFAULT_SECTIONS = [
        {"title":"Introduction",         "level":1,"figures":0,"tables":0,
         "fig_captions":[],"tbl_captions":[]},
        {"title":"Mise en contexte",     "level":2,"figures":0,"tables":0,
         "fig_captions":[],"tbl_captions":[]},
        {"title":"Objectifs",            "level":2,"figures":0,"tables":0,
         "fig_captions":[],"tbl_captions":[]},
        {"title":"Développement",        "level":1,"figures":0,"tables":0,
         "fig_captions":[],"tbl_captions":[]},
        {"title":"Résultats",            "level":1,"figures":0,"tables":0,
         "fig_captions":[],"tbl_captions":[]},
        {"title":"Analyse et discussion","level":1,"figures":0,"tables":0,
         "fig_captions":[],"tbl_captions":[]},
        {"title":"Conclusion",           "level":1,"figures":0,"tables":0,
         "fig_captions":[],"tbl_captions":[]},
    ]
    if "sections" not in st.session_state:
        st.session_state.sections = DEFAULT_SECTIONS.copy()

    ca, cb, _ = st.columns([1, 1, 4])
    with ca:
        if st.button("➕ Section", use_container_width=True):
            st.session_state.sections.append(
                {"title":"Nouvelle section","level":1,"figures":0,"tables":0,
                 "fig_captions":[],"tbl_captions":[]})
            st.rerun()
    with cb:
        if st.button("🔄 Réinitialiser", use_container_width=True):
            st.session_state.sections = DEFAULT_SECTIONS.copy()
            st.rerun()

    to_del = []
    LEVEL_LABELS = {1: "Titre principal", 2: "Sous-titre", 3: "Sous-sous-titre", 4: "Sous-section"}
    for i, sec in enumerate(st.session_state.sections):
        indent = "　" * (sec["level"] - 1)   # espace pleine largeur, indentation visuelle propre
        level_tag = LEVEL_LABELS.get(sec["level"], "")
        with st.expander(f"{indent}**{sec['title']}**  ·  {level_tag}", expanded=False):
            r1c1, r1c2, r1c3 = st.columns([3, 1, 0.6])
            with r1c1:
                new_t = st.text_input("Titre", value=sec["title"], key=f"st_{i}")
            with r1c2:
                new_l = st.selectbox("Niveau", [1,2,3,4], index=sec["level"]-1,
                                      key=f"sl_{i}",
                                      format_func=lambda x: f"H{x}")
            with r1c3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"sd_{i}"):
                    to_del.append(i); st.rerun()

            r2c1, r2c2 = st.columns(2)
            with r2c1:
                new_nf = st.number_input("Figures", min_value=0, max_value=10,
                                          value=sec.get("figures",0), key=f"snf_{i}")
            with r2c2:
                new_nt = st.number_input("Tableaux", min_value=0, max_value=10,
                                          value=sec.get("tables",0), key=f"snt_{i}")

            new_fcaps, new_tcaps = [], []
            if new_nf > 0:
                st.markdown("**Légendes des figures** *(modifiables dans Word)*")
                for fi in range(int(new_nf)):
                    ex  = (sec.get("fig_captions",[]) or [])[fi] \
                          if fi < len(sec.get("fig_captions",[])) else ""
                    cap = st.text_input(f"  Figure {fi+1}", value=ex,
                                         key=f"sfc_{i}_{fi}",
                                         placeholder="Description de la figure…")
                    new_fcaps.append(cap)

            if new_nt > 0:
                st.markdown("**Légendes des tableaux** *(modifiables dans Word)*")
                for ti in range(int(new_nt)):
                    ex  = (sec.get("tbl_captions",[]) or [])[ti] \
                          if ti < len(sec.get("tbl_captions",[])) else ""
                    cap = st.text_input(f"  Tableau {ti+1}", value=ex,
                                         key=f"stc_{i}_{ti}",
                                         placeholder="Description du tableau…")
                    new_tcaps.append(cap)

            st.session_state.sections[i] = {
                "title": new_t, "level": new_l,
                "figures": int(new_nf), "tables": int(new_nt),
                "fig_captions": new_fcaps, "tbl_captions": new_tcaps,
            }

    for idx in sorted(to_del, reverse=True):
        st.session_state.sections.pop(idx)

# ══════════════════════════════════════════════
# TAB 4 — RÉFÉRENCES
# ══════════════════════════════════════════════
with tab4:
    st.markdown("""<div class="info-box">
    Colle le <b>DOI</b> ou l'<b>URL</b> de l'article → la référence est générée automatiquement
    via <b>CrossRef</b> (base de données académique gratuite).<br>
    Fonctionne pour ~95 % des articles scientifiques. Pour les autres, entre le texte directement.
    </div>""", unsafe_allow_html=True)

    # ── Fonctions CrossRef (100 % gratuites) ──

    def _format_ref(data, style):
        authors = data.get("author", [])

        def fmt_author(a):
            family   = a.get("family","")
            given    = a.get("given","")
            if not family: return a.get("name","")
            initials = " ".join(f"{n[0]}." for n in given.split()) if given else ""
            if style == "APA":    return f"{family}, {initials}".strip(", ")
            elif style == "IEEE": return f"{initials} {family}".strip()
            else:                 return f"{family}, {given}".strip(", ")

        if style == "APA":
            if not authors:         auth_str = "Auteur inconnu"
            elif len(authors) == 1: auth_str = fmt_author(authors[0])
            elif len(authors) <= 7:
                auth_str  = ", ".join(fmt_author(a) for a in authors[:-1])
                auth_str += f", & {fmt_author(authors[-1])}"
            else:
                auth_str  = ", ".join(fmt_author(a) for a in authors[:6])
                auth_str += f", … {fmt_author(authors[-1])}"
        elif style == "IEEE":
            if not authors:          auth_str = "Auteur inconnu"
            elif len(authors) <= 3:  auth_str = ", ".join(fmt_author(a) for a in authors)
            else:                    auth_str = fmt_author(authors[0]) + " et al."
        else:
            if not authors:          auth_str = "Auteur inconnu"
            elif len(authors) == 1:  auth_str = fmt_author(authors[0])
            else:                    auth_str = fmt_author(authors[0]) + " et al."

        year = ""
        for f2 in ["published","published-print","published-online","issued"]:
            dp = data.get(f2,{}).get("date-parts",[[]])
            if dp and dp[0]: year = str(dp[0][0]); break

        titles    = data.get("title",[""]); title = titles[0] if titles else ""
        ct        = data.get("container-title",[]); container = ct[0] if ct else ""
        volume    = data.get("volume","")
        issue     = data.get("issue","")
        pages     = data.get("page","")
        doi_clean = data.get("DOI","")

        if style == "APA":
            fmt = f"{auth_str} ({year}). {title}."
            if container: fmt += f" {container}"
            if volume:    fmt += f", {volume}"
            if issue:     fmt += f"({issue})"
            if pages:     fmt += f", {pages}"
            fmt += "."
            if doi_clean: fmt += f" https://doi.org/{doi_clean}"
        elif style == "IEEE":
            fmt = f"{auth_str}, \"{title},\""
            if container: fmt += f" {container},"
            if volume:    fmt += f" vol. {volume},"
            if issue:     fmt += f" no. {issue},"
            if pages:     fmt += f" pp. {pages},"
            if year:      fmt += f" {year}."
            if doi_clean: fmt += f" doi: {doi_clean}."
        else:
            fmt = f"{auth_str}. \"{title}.\""
            if container: fmt += f" {container}"
            if volume:    fmt += f" {volume}"
            if issue:     fmt += f", no. {issue}"
            if year:      fmt += f" ({year})"
            if pages:     fmt += f": {pages}"
            fmt += "."
            if doi_clean: fmt += f" https://doi.org/{doi_clean}."

        return fmt

    def fetch_reference(doi_or_url):
        """CrossRef uniquement — 100 % gratuit, aucune clé API."""
        raw = doi_or_url.strip()

        # Extraire le DOI brut
        doi = raw
        for prefix in ["https://doi.org/","http://doi.org/","doi.org/","DOI:","doi:"]:
            if doi.lower().startswith(prefix.lower()):
                doi = doi[len(prefix):].strip()
                break

        # Vérifier si ça ressemble à un DOI
        looks_like_doi = doi.startswith("10.") and "/" in doi

        if not looks_like_doi:
            return None, (
                "L'URL fournie ne contient pas de DOI reconnaissable (format 10.xxxx/xxx).\n"
                "Entre la référence manuellement dans le champ ci-dessous."
            )

        try:
            r = requests.get(
                f"https://api.crossref.org/works/{doi}",
                headers={
                    "User-Agent":
                    "PolyReport/1.0 (mailto:najeeb.lababidi@etud.polymtl.ca)"
                },
                timeout=12
            )
            if r.status_code == 200:
                return _format_ref(r.json()["message"], ref_style), "ok"
            elif r.status_code == 404:
                return None, "DOI introuvable dans CrossRef — entre la référence manuellement."
            else:
                return None, f"Erreur CrossRef ({r.status_code}) — entre la référence manuellement."
        except requests.exceptions.Timeout:
            return None, "Délai dépassé — vérifie ta connexion et réessaie."
        except Exception as e:
            return None, f"Erreur réseau : {e}"

    # ── Interface références ────────────────────

    if "references" not in st.session_state:
        st.session_state.references = [{"url":"","formatted":"","status":""}]

    col_add_ref, _ = st.columns([1, 5])
    with col_add_ref:
        if st.button("➕ Référence"):
            st.session_state.references.append({"url":"","formatted":"","status":""})
            st.rerun()

    ref_del = []
    for i, ref in enumerate(st.session_state.references):
        preview = ref.get("formatted","")[:55] + "…" if ref.get("formatted") else "Nouvelle référence"
        # Version du widget : change chaque fois qu'on régénère, force Streamlit
        # à redessiner le text_area avec la nouvelle valeur au lieu de garder
        # l'ancienne valeur cachée liée à l'ancienne clé.
        widget_ver = ref.get("_v", 0)

        with st.expander(f"[{i+1}] {preview}", expanded=(not ref.get("formatted"))):
            rc1, rc2, rc3 = st.columns([3, 1, 0.6])
            with rc1:
                new_url = st.text_input(
                    "DOI ou URL de l'article",
                    value=ref.get("url",""),
                    key=f"rurl_{i}",
                    placeholder="https://doi.org/10.xxxx/xxxxx"
                )
            with rc2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🔍 Générer", key=f"rfetch_{i}", use_container_width=True):
                    if new_url.strip():
                        with st.spinner("Recherche CrossRef…"):
                            fmt, status = fetch_reference(new_url.strip())
                        if fmt:
                            st.session_state.references[i]["formatted"] = fmt
                            st.session_state.references[i]["status"]    = "ok"
                            st.session_state.references[i]["url"]       = new_url.strip()
                            st.session_state.references[i]["_v"]        = widget_ver + 1
                        else:
                            st.session_state.references[i]["status"] = "err:" + status
                            st.session_state.references[i]["url"]    = new_url.strip()
                        st.rerun()
                    else:
                        st.warning("Entre un DOI ou URL d'abord.")
            with rc3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"rdel_{i}"):
                    ref_del.append(i); st.rerun()

            new_fmt = st.text_area(
                "Référence formatée (modifiable)",
                value=ref.get("formatted",""),
                key=f"rfmt_{i}_{widget_ver}", height=75,
                placeholder="La référence apparaîtra ici, ou entre-la manuellement…"
            )
            st.session_state.references[i]["formatted"] = new_fmt
            st.session_state.references[i]["url"]       = new_url

            status = ref.get("status","")
            if status == "ok":
                st.markdown('<div class="success-box">✅ Référence générée via CrossRef</div>',
                            unsafe_allow_html=True)
            elif status.startswith("err:"):
                st.markdown(f'<div class="warn-box">⚠️ {status[4:]}</div>',
                            unsafe_allow_html=True)

    for idx in sorted(ref_del, reverse=True):
        st.session_state.references.pop(idx)

# ══════════════════════════════════════════════
# TAB 5 — ANNEXES
# ══════════════════════════════════════════════
with tab5:
    st.markdown("""<div class="info-box">
    Même logique que les sections — configure titres, figures et tableaux.<br>
    Chaque annexe a sa propre numérotation de pages (repart à 1).<br>
    Cite chaque annexe dans le corps du rapport (ex : <i>voir Annexe A</i>).
    </div>""", unsafe_allow_html=True)

    if "annexes" not in st.session_state:
        st.session_state.annexes = [
            {"title":"Données supplémentaires","figures":0,"tables":1,
             "fig_captions":[],"tbl_captions":["Données brutes mesurées."]}
        ]

    ann_letters = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    ca2, _ = st.columns([1, 5])
    with ca2:
        if st.button("➕ Annexe", use_container_width=True):
            st.session_state.annexes.append(
                {"title":"Nouvelle annexe","figures":0,"tables":0,
                 "fig_captions":[],"tbl_captions":[]})
            st.rerun()

    ann_del = []
    for i, ann in enumerate(st.session_state.annexes):
        letter = ann_letters[i] if i < len(ann_letters) else str(i+1)
        with st.expander(f"Annexe {letter} — {ann['title']}", expanded=(i==0)):
            ac1, ac2 = st.columns([3, 0.6])
            with ac1:
                new_at = st.text_input("Titre", value=ann["title"], key=f"at_{i}")
            with ac2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"ad_{i}"):
                    ann_del.append(i); st.rerun()

            ac3, ac4 = st.columns(2)
            with ac3:
                new_anf = st.number_input("Figures", min_value=0, max_value=10,
                                           value=ann.get("figures",0), key=f"anf_{i}")
            with ac4:
                new_ant = st.number_input("Tableaux", min_value=0, max_value=10,
                                           value=ann.get("tables",0), key=f"ant_{i}")

            new_afcaps, new_atcaps = [], []
            if new_anf > 0:
                st.markdown("**Légendes figures**")
                for fi in range(int(new_anf)):
                    ex  = (ann.get("fig_captions",[]) or [])[fi] \
                          if fi < len(ann.get("fig_captions",[])) else ""
                    cap = st.text_input(f"  Figure {fi+1}", value=ex,
                                         key=f"afc_{i}_{fi}",
                                         placeholder="Description…")
                    new_afcaps.append(cap)
            if new_ant > 0:
                st.markdown("**Légendes tableaux**")
                for ti in range(int(new_ant)):
                    ex  = (ann.get("tbl_captions",[]) or [])[ti] \
                          if ti < len(ann.get("tbl_captions",[])) else ""
                    cap = st.text_input(f"  Tableau {ti+1}", value=ex,
                                         key=f"atc_{i}_{ti}",
                                         placeholder="Description…")
                    new_atcaps.append(cap)

            st.session_state.annexes[i] = {
                "title": new_at, "figures": int(new_anf), "tables": int(new_ant),
                "fig_captions": new_afcaps, "tbl_captions": new_atcaps,
            }

    for idx in sorted(ann_del, reverse=True):
        st.session_state.annexes.pop(idx)

# ─────────────────────────────────────────────
# GÉNÉRATION
# ─────────────────────────────────────────────

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown('<div class="section-label" style="font-size:1.1rem;margin-top:.3rem;">'
            '🚀 Générer mon rapport</div>', unsafe_allow_html=True)

cb1, cb2 = st.columns([2, 3])
with cb1:
    gen_clicked = st.button("⚡ Générer le rapport .docx",
                             type="primary", use_container_width=True)
with cb2:
    st.markdown("""<div class="warn-box">
    ⚠️ <b>Après ouverture dans Word :</b> Ctrl+A → F9 pour mettre à jour la TDM,
    liste des figures et liste des tableaux.<br>
    Les zones grises sont des <b>placeholders</b> — les remplacer par ta rédaction.
    </div>""", unsafe_allow_html=True)

if gen_clicked:
    refs_clean = [
        {"formatted": r.get("formatted",""), "url": r.get("url","")}
        for r in st.session_state.get("references",[])
        if r.get("formatted","").strip()
    ]

    config = {
        "font":          font,
        "font_size":     font_size,
        "ref_style":     ref_style,
        "doc_date":      doc_date,
        "course_code":   course_code,
        "course_name":   course_name,
        "group_number":  group_number,
        "report_title":  report_title,
        "semester":      semester,
        "professors":    professors,
        "members":       members,
        "is_individual": is_individual,
        "resume_fr":     resume_fr,
        "abstract_en":   st.session_state.get("abstract_value",""),
        "keywords_fr":   keywords_fr,
        "keywords_en":   keywords_en,
        "sections":      st.session_state.get("sections",[]),
        "references":    refs_clean,
        "annexes":       st.session_state.get("annexes",[]),
    }

    with st.spinner("Génération du rapport en cours…"):
        try:
            buf      = generate_poly_report(config)
            safe_t   = (report_title.strip() or "Rapport_technique")[:50].replace(" ","_").replace("/","-")
            filename = f"{safe_t}.docx"
            st.success("✅ Rapport généré avec succès !")
            st.download_button(
                label="📥 Télécharger le .docx",
                data=buf,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
            st.exception(e)
