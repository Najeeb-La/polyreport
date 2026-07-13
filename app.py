"""
PolyRapport — Générateur de rapports techniques
Polytechnique Montréal · Par Najeeb Lababidi · 100 % gratuit
"""

import streamlit as st
from generate_report import generate_poly_report, FONTS_AVAILABLE, REF_STYLES
from datetime import date
import requests, os, base64, json

st.set_page_config(
    page_title="PolyRapport — Générateur de rapports techniques",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────

st.markdown("""
<style>
  /* Hero */
  .hero-wrap {
    background: linear-gradient(120deg,#A60000 0%,#7A0000 100%);
    border-radius:14px; padding:1.5rem 2rem;
    margin-bottom:1.6rem; display:flex;
    align-items:center; gap:1.6rem;
    box-shadow:0 2px 14px rgba(166,0,0,.18);
  }
  .hero-logo { background:white; border-radius:10px; padding:8px 14px; flex-shrink:0; }
  .hero-text h1 { color:white; margin:0; font-size:1.5rem; font-weight:700; }
  .hero-text .sub    { color:rgba(255,255,255,.85); font-size:.88rem; margin:.35rem 0 0; }
  .hero-text .credit { color:rgba(255,255,255,.62); font-size:.76rem; margin:.5rem 0 0; line-height:1.5; }

  /* Badges */
  .badge {
    background:#A60000; color:white; padding:.16rem .58rem;
    border-radius:5px; font-size:.7rem; font-weight:700;
    letter-spacing:.04em; margin-right:.45rem; text-transform:uppercase;
  }
  .badge-gray {
    background:#444; color:white; padding:.16rem .58rem;
    border-radius:5px; font-size:.7rem; font-weight:700;
    letter-spacing:.04em; margin-right:.45rem; text-transform:uppercase;
  }
  .section-label { font-size:.95rem; font-weight:700; color:#1a1a1a; }

  /* Boxes */
  .info-box  { background:#F7F7F8; border-left:3px solid #A60000;
               padding:.55rem .9rem; border-radius:0 8px 8px 0;
               margin:.45rem 0; font-size:.82rem; color:#444; line-height:1.5; }
  .warn-box  { background:#FFF7E6; border-left:3px solid #E8A33D;
               padding:.55rem .9rem; border-radius:0 8px 8px 0;
               margin:.45rem 0; font-size:.82rem; color:#5a4a2a; line-height:1.5; }
  .success-box { background:#EAF7ED; border-left:3px solid #2E9E50;
                 padding:.55rem .9rem; border-radius:0 8px 8px 0;
                 margin:.45rem 0; font-size:.82rem; color:#1d6b35; line-height:1.5; }

  /* Tabs */
  div[data-testid="stTabs"] button { font-size:.92rem; }
  div[data-testid="stTabs"] button[aria-selected="true"] {
    border-bottom:3px solid #A60000 !important;
    color:#A60000 !important; font-weight:700;
  }

  /* Sidebar */
  section[data-testid="stSidebar"] { background:#FAFAFA; border-right:1px solid #ECECEC; }

  /* Expanders */
  div[data-testid="stExpander"] {
    border:1px solid #E8E8E8 !important;
    border-radius:8px !important; margin-bottom:.45rem;
  }

  /* Bouton primaire */
  button[kind="primary"] {
    background:#A60000 !important; border:none !important;
    font-weight:700 !important; border-radius:8px !important;
  }
  button[kind="primary"]:hover { background:#8a0000 !important; }

  /* Footer */
  .footer-credit {
    font-size:.73rem; color:#999; text-align:center; line-height:1.7;
  }
  .footer-credit a { color:#A60000; text-decoration:none; font-weight:600; }
  .footer-credit a:hover { text-decoration:underline; }

  /* Séparateur de section */
  hr.section-sep { border:none; border-top:1px solid #E8E8E8; margin:1rem 0; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────

_DIR  = os.path.dirname(os.path.abspath(__file__))
_LOGO = os.path.join(_DIR, "poly_logo.png")
_logo_b64 = ""
if os.path.exists(_LOGO):
    with open(_LOGO, "rb") as _f:
        _logo_b64 = base64.b64encode(_f.read()).decode()

st.markdown(f"""
<div class="hero-wrap">
  {f'<div class="hero-logo"><img src="data:image/png;base64,{_logo_b64}" width="120"></div>' if _logo_b64 else ''}
  <div class="hero-text">
    <h1>PolyRapport — Générateur de rapports techniques</h1>
    <div class="sub">Polytechnique Montréal &nbsp;·&nbsp; Rapports techniques académiques &nbsp;·&nbsp; 100 % gratuit</div>
    <div class="credit">
      Développé par <strong>Najeeb Lababidi</strong> · Étudiant en génie aérospatial, Polytechnique Montréal<br>
      Génère le squelette Word complet — rédigez votre texte directement dans le .docx
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
                              help="Cambria est recommandée pour les rapports académiques à Poly")
    font_size = st.slider("Taille du corps (pt)", 10, 14, 12)
    ref_style = st.selectbox("Style de références", REF_STYLES, index=0,
                              help="IEEE est recommandé en génie")
    st.markdown("---")
    st.markdown("### 💡 Comment utiliser")
    st.markdown("""<div class="info-box">
    <b>1.</b> Remplissez le formulaire<br>
    <b>2.</b> Cliquez « Générer »<br>
    <b>3.</b> Ouvrez le .docx dans Word<br>
    <b>4.</b> Ctrl+A → F9 (table des matières)<br>
    <b>5.</b> Rédigez dans les zones grises
    </div>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("""<div class="footer-credit">
    Outil gratuit pour les étudiants de Polytechnique Montréal<br>
    © Najeeb Lababidi · Génie aérospatial<br>
    <a href="mailto:najeeb.lababidi@etud.polymtl.ca">Signaler un problème</a>
    </div>""", unsafe_allow_html=True)

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

    st.markdown('<span class="badge">Document</span>'
                '<span class="section-label"> Informations générales</span>',
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
            season    = st.selectbox("Session", ["Hiver","Été","Automne"], index=0)
        with sc2:
            year_sess = st.number_input("Année", min_value=2020, max_value=2035,
                                         value=date.today().year, step=1)
        semester = f"{season} {year_sess}"
    with c2:
        group_number = st.text_input("Groupe-cours n°", value="",
                                      placeholder="ex : 02",
                                      help="Laisser vide si non applicable")

    st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
    st.markdown('<span class="badge">Cours</span>'
                '<span class="section-label"> Cours associé</span>',
                unsafe_allow_html=True)
    cc1, cc2 = st.columns(2)
    with cc1:
        course_code = st.text_input("Sigle", value="",
                                     placeholder="ex : MEC 3520")
    with cc2:
        course_name = st.text_input("Titre du cours", value="",
                                     placeholder="ex : Industrialisation des produits")

    st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
    st.markdown('<span class="badge">Rapport</span>'
                '<span class="section-label"> Titre du rapport</span>',
                unsafe_allow_html=True)
    report_title = st.text_input("Titre", value="",
                                  placeholder="Titre complet de votre rapport technique")

    st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
    st.markdown('<span class="badge">Équipe</span>'
                '<span class="section-label"> Encadrement et auteur(s)</span>',
                unsafe_allow_html=True)

    professors_raw = st.text_input(
        "Présenté à (séparés par virgule si plusieurs)",
        value="", placeholder="ex : Prof. Marwan Azzi, Dr. Leila Ahmadi",
        help="Chaque professeur apparaîtra sur une ligne distincte dans le rapport"
    )
    professors = [p.strip() for p in professors_raw.split(",") if p.strip()]

    is_individual = st.toggle(
        "🎓 Travail individuel — pas de tableau de signatures", value=False
    )

    if is_individual:
        st.markdown('<span class="badge-gray">Étudiant</span>'
                    '<span class="section-label"> Vos informations</span>',
                    unsafe_allow_html=True)
        st.markdown("""<div class="info-box">
        Format : <code>Nom Prénom | Matricule</code>
        </div>""", unsafe_allow_html=True)
        m_raw = st.text_input("", value="",
                               placeholder="Nom Prénom | 2XXXXXX",
                               label_visibility="collapsed")
        if "|" in m_raw:
            pts = m_raw.split("|", 1)
            members = [{"name": pts[0].strip(), "matricule": pts[1].strip()}]
        elif m_raw.strip():
            members = [{"name": m_raw.strip(), "matricule": ""}]
        else:
            members = []
    else:
        st.markdown('<span class="badge-gray">Équipe</span>'
                    '<span class="section-label"> Membres</span>',
                    unsafe_allow_html=True)
        st.markdown("""<div class="info-box">
        Un membre par ligne · Format : <code>Nom Prénom | Matricule</code>
        </div>""", unsafe_allow_html=True)
        members_raw = st.text_area("", label_visibility="collapsed",
                                    height=110, value="",
                                    placeholder="Nom Prénom | 2XXXXXX\nNom Prénom | 2XXXXXX")
        members = []
        for line in members_raw.strip().split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                pts = line.split("|", 1)
                members.append({"name": pts[0].strip(), "matricule": pts[1].strip()})
            else:
                members.append({"name": line, "matricule": ""})

# ══════════════════════════════════════════════
# TAB 2 — RÉSUMÉ & ABSTRACT
# ══════════════════════════════════════════════
with tab2:
    st.markdown("""<div class="info-box">
    ℹ️ <b>Le résumé se rédige après avoir complété le rapport.</b>
    Si vous n'avez pas encore votre texte, laissez ces champs vides —
    un placeholder sera inséré dans le .docx et vous pourrez le compléter plus tard directement dans Word.
    </div>""", unsafe_allow_html=True)

    st.markdown('<span class="badge">FR</span>'
                '<span class="section-label"> Résumé en français</span>',
                unsafe_allow_html=True)
    resume_fr   = st.text_area("", label_visibility="collapsed", height=160,
                                placeholder="(Optionnel) Coller votre résumé ici une fois le rapport rédigé…")
    keywords_fr = st.text_input("Mots clés FR (séparés par virgule)",
                                 placeholder="ex : turbofan, propulsion, efficacité isentropique")

    st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
    st.markdown('<span class="badge">EN</span>'
                '<span class="section-label"> Abstract</span>',
                unsafe_allow_html=True)

    # Initialisation état traduction
    if "abstract_value" not in st.session_state:
        st.session_state["abstract_value"] = ""
    if "abstract_ver" not in st.session_state:
        st.session_state["abstract_ver"] = 0

    col_abs, col_btn = st.columns([4, 1])
    with col_abs:
        abstract_en = st.text_area(
            "", label_visibility="collapsed", height=160,
            placeholder="(Optionnel) Coller ou écrire votre abstract ici…",
            value=st.session_state["abstract_value"],
            key=f"abstract_widget_{st.session_state['abstract_ver']}",
        )
        st.session_state["abstract_value"] = abstract_en

    with col_btn:
        st.markdown("<br><br>", unsafe_allow_html=True)
        translate_clicked = st.button("🌐 Traduire\nle résumé",
                                       help="Traduction automatique FR → EN via Google Translate (gratuit)",
                                       use_container_width=True)

    if translate_clicked:
        if not resume_fr.strip():
            st.warning("Veuillez d'abord saisir votre résumé en français.")
        else:
            with st.spinner("Traduction en cours…"):
                translated   = None
                error_detail = None
                try:
                    from deep_translator import GoogleTranslator
                    text = resume_fr.strip()
                    MAX  = 4500
                    if len(text) <= MAX:
                        translated = GoogleTranslator(source='fr', target='en').translate(text)
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
                st.session_state["abstract_ver"]  += 1
                st.session_state["_trans_ok"]      = True
                st.rerun()
            else:
                st.markdown(f"""<div class="warn-box">
                ⚠️ Traduction automatique indisponible pour le moment.<br>
                <b>Alternative gratuite :</b> copiez votre résumé →
                <a href="https://www.deepl.com/translator" target="_blank"><b>DeepL.com</b></a>
                → collez la traduction dans le champ ci-dessus.
                </div>""", unsafe_allow_html=True)

    if st.session_state.pop("_trans_ok", False):
        st.success("✅ Traduction générée — vérifiez les termes techniques avant de remettre.")

    keywords_en = st.text_input("Keywords EN (comma-separated)",
                                 placeholder="ex: turbofan, propulsion, isentropic efficiency")

# ══════════════════════════════════════════════
# TAB 3 — SECTIONS
# ══════════════════════════════════════════════
with tab3:
    st.markdown("""<div class="info-box">
    Configurez la <b>structure</b> de votre rapport : titres, hiérarchie, figures et tableaux.<br>
    Le .docx contiendra des zones grises (placeholders) avec les styles Word corrects —
    rédigez directement dans Word. <b>Ctrl+A → F9</b> met à jour la table des matières automatiquement.
    </div>""", unsafe_allow_html=True)

    LEVEL_LABELS = {1:"Titre 1", 2:"Titre 2", 3:"Titre 3", 4:"Titre 4"}

    DEFAULT_SECTIONS = [
        {"title":"Introduction",         "level":1,"figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]},
        {"title":"Mise en contexte",     "level":2,"figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]},
        {"title":"Objectifs",            "level":2,"figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]},
        {"title":"Développement",        "level":1,"figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]},
        {"title":"Résultats",            "level":1,"figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]},
        {"title":"Analyse et discussion","level":1,"figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]},
        {"title":"Conclusion",           "level":1,"figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]},
    ]

    if "sections" not in st.session_state:
        st.session_state.sections = [dict(s) for s in DEFAULT_SECTIONS]

    ca, cb, _ = st.columns([1, 1.3, 4])
    with ca:
        if st.button("➕ Ajouter une section", use_container_width=True):
            st.session_state.sections.append(
                {"title":"Nouvelle section","level":1,"figures":0,"tables":0,
                 "tbl_rows":3,"tbl_cols":3,"fig_captions":[],"tbl_captions":[]})
            st.rerun()
    with cb:
        if st.button("🔄 Réinitialiser la structure", use_container_width=True):
            st.session_state.sections = [dict(s) for s in DEFAULT_SECTIONS]
            st.rerun()

    to_del = []
    for i, sec in enumerate(st.session_state.sections):
        indent    = "　" * (sec["level"] - 1)
        lbl       = LEVEL_LABELS.get(sec["level"], "")
        exp_label = f"{indent}{sec['title']}  ·  {lbl}"

        with st.expander(exp_label, expanded=False):
            r1, r2, r3 = st.columns([3, 1.2, 0.6])
            with r1:
                new_t = st.text_input("Titre de la section", value=sec["title"], key=f"st_{i}")
            with r2:
                new_l = st.selectbox(
                    "Niveau",
                    options=[1,2,3,4],
                    index=sec["level"]-1,
                    key=f"sl_{i}",
                    format_func=lambda x: LEVEL_LABELS[x]
                )
            with r3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"sd_{i}", help="Supprimer cette section"):
                    to_del.append(i)

            r4, r5 = st.columns(2)
            with r4:
                new_nf = st.number_input("Nombre de figures", min_value=0, max_value=10,
                                          value=sec.get("figures",0), key=f"snf_{i}")
            with r5:
                new_nt = st.number_input("Nombre de tableaux", min_value=0, max_value=10,
                                          value=sec.get("tables",0), key=f"snt_{i}")

            # Dimensions des tableaux
            new_tr, new_tc = sec.get("tbl_rows",3), sec.get("tbl_cols",3)
            if new_nt > 0:
                rt, ct = st.columns(2)
                with rt:
                    new_tr = st.number_input("Lignes par tableau (incl. en-tête)",
                                              min_value=2, max_value=20,
                                              value=sec.get("tbl_rows",3), key=f"str_{i}")
                with ct:
                    new_tc = st.number_input("Colonnes par tableau",
                                              min_value=1, max_value=10,
                                              value=sec.get("tbl_cols",3), key=f"stc2_{i}")

            # Légendes figures
            new_fcaps = []
            if new_nf > 0:
                st.caption("Légendes des figures (optionnel — modifiables dans Word)")
                for fi in range(int(new_nf)):
                    ex  = (sec.get("fig_captions") or [])[fi] if fi < len(sec.get("fig_captions") or []) else ""
                    cap = st.text_input(f"  Figure {fi+1}", value=ex, key=f"sfc_{i}_{fi}",
                                         placeholder="Description de la figure…")
                    new_fcaps.append(cap)

            # Légendes tableaux
            new_tcaps = []
            if new_nt > 0:
                st.caption("Légendes des tableaux (optionnel — modifiables dans Word)")
                for ti in range(int(new_nt)):
                    ex  = (sec.get("tbl_captions") or [])[ti] if ti < len(sec.get("tbl_captions") or []) else ""
                    cap = st.text_input(f"  Tableau {ti+1}", value=ex, key=f"stca_{i}_{ti}",
                                         placeholder="Description du tableau…")
                    new_tcaps.append(cap)

            st.session_state.sections[i] = {
                "title": new_t, "level": new_l,
                "figures": int(new_nf), "tables": int(new_nt),
                "tbl_rows": int(new_tr), "tbl_cols": int(new_tc),
                "fig_captions": new_fcaps, "tbl_captions": new_tcaps,
            }

    if to_del:
        for idx in sorted(to_del, reverse=True):
            st.session_state.sections.pop(idx)
        st.rerun()

# ══════════════════════════════════════════════
# TAB 4 — RÉFÉRENCES
# ══════════════════════════════════════════════
with tab4:
    st.markdown("""<div class="info-box">
    Collez le <b>DOI</b>, une <b>URL</b> (Moodle, site web, PDF, etc.) ou un lien d'article →
    cliquez <b>🔍 Générer</b> pour obtenir la référence formatée.<br>
    Pour les DOI (ex : <code>https://doi.org/10.xxxx/xxx</code>), la génération est automatique et précise via CrossRef.
    Pour les autres URLs, entrez la référence manuellement dans le champ texte.
    </div>""", unsafe_allow_html=True)

    def _format_crossref(data, style):
        authors = data.get("author", [])
        def fmt(a):
            fam  = a.get("family",""); giv = a.get("given","")
            if not fam: return a.get("name","")
            ini  = " ".join(f"{n[0]}." for n in giv.split()) if giv else ""
            if style=="APA":    return f"{fam}, {ini}".strip(", ")
            elif style=="IEEE": return f"{ini} {fam}".strip()
            else:               return f"{fam}, {giv}".strip(", ")

        if not authors:       auth = "Auteur inconnu"
        elif style=="APA":
            if len(authors)==1:   auth = fmt(authors[0])
            elif len(authors)<=7:
                auth = ", ".join(fmt(a) for a in authors[:-1]) + f", & {fmt(authors[-1])}"
            else:
                auth = ", ".join(fmt(a) for a in authors[:6]) + f", … {fmt(authors[-1])}"
        elif style=="IEEE":
            auth = ", ".join(fmt(a) for a in authors) if len(authors)<=3 \
                   else fmt(authors[0])+" et al."
        else:
            auth = fmt(authors[0]) if len(authors)==1 else fmt(authors[0])+" et al."

        year = ""
        for f2 in ["published","published-print","published-online","issued"]:
            dp = data.get(f2,{}).get("date-parts",[[]])
            if dp and dp[0]: year=str(dp[0][0]); break

        title     = (data.get("title",[""])[0]) if data.get("title") else ""
        container = (data.get("container-title",[""])[0]) if data.get("container-title") else ""
        volume    = data.get("volume",""); issue=data.get("issue","")
        pages     = data.get("page","");  doi=data.get("DOI","")

        if style=="APA":
            out = f"{auth} ({year}). {title}."
            if container: out += f" {container}"
            if volume:    out += f", {volume}"
            if issue:     out += f"({issue})"
            if pages:     out += f", {pages}"
            out += "."
            if doi: out += f" https://doi.org/{doi}"
        elif style=="IEEE":
            out = f"{auth}, \"{title},\""
            if container: out += f" {container},"
            if volume:    out += f" vol. {volume},"
            if issue:     out += f" no. {issue},"
            if pages:     out += f" pp. {pages},"
            if year:      out += f" {year}."
            if doi:       out += f" doi: {doi}."
        else:
            out = f"{auth}. \"{title}.\""
            if container: out += f" {container}"
            if volume:    out += f" {volume}"
            if issue:     out += f", no. {issue}"
            if year:      out += f" ({year})"
            if pages:     out += f": {pages}"
            out += "."
            if doi: out += f" https://doi.org/{doi}."
        return out

    def fetch_reference(raw_url):
        """
        Essaie CrossRef si DOI détecté.
        Pour toute autre URL, retourne un message clair invitant à saisir manuellement.
        """
        raw = raw_url.strip()
        doi = raw
        for pfx in ["https://doi.org/","http://doi.org/","doi.org/","DOI:","doi:"]:
            if doi.lower().startswith(pfx.lower()):
                doi = doi[len(pfx):].strip(); break

        if doi.startswith("10.") and "/" in doi:
            try:
                r = requests.get(
                    f"https://api.crossref.org/works/{doi}",
                    headers={"User-Agent":"PolyRapport/1.0 (mailto:najeeb.lababidi@etud.polymtl.ca)"},
                    timeout=12
                )
                if r.status_code == 200:
                    return _format_crossref(r.json()["message"], ref_style), "crossref"
                return None, f"DOI non trouvé dans CrossRef (code {r.status_code}). Saisissez la référence manuellement."
            except requests.exceptions.Timeout:
                return None, "Délai dépassé. Vérifiez votre connexion et réessayez."
            except Exception as e:
                return None, f"Erreur : {e}"
        else:
            return None, (
                "Cette URL n'est pas un DOI reconnaissable. "
                "Saisissez la référence formatée manuellement dans le champ ci-dessous, "
                "et collez l'URL dans le champ URL — elle apparaîtra comme lien cliquable dans le rapport."
            )

    if "references" not in st.session_state:
        st.session_state.references = [{"url":"","formatted":"","status":"","_v":0}]

    col_add_ref, _ = st.columns([1, 5])
    with col_add_ref:
        if st.button("➕ Ajouter une référence"):
            st.session_state.references.append({"url":"","formatted":"","status":"","_v":0})
            st.rerun()

    ref_del = []
    for i, ref in enumerate(st.session_state.references):
        preview = ref.get("formatted","")[:60]+"…" if ref.get("formatted") else "Nouvelle référence"
        wv      = ref.get("_v", 0)

        with st.expander(f"[{i+1}]  {preview}", expanded=(not ref.get("formatted"))):
            rc1, rc2, rc3 = st.columns([3, 1, 0.6])
            with rc1:
                new_url = st.text_input(
                    "DOI ou URL de la source",
                    value=ref.get("url",""), key=f"rurl_{i}",
                    placeholder="https://doi.org/10.xxxx/xxx  ou  https://moodle.polymtl.ca/…"
                )
            with rc2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🔍 Générer", key=f"rfetch_{i}", use_container_width=True):
                    if new_url.strip():
                        with st.spinner("Recherche en cours…"):
                            fmt, status = fetch_reference(new_url.strip())
                        if fmt:
                            st.session_state.references[i].update(
                                {"formatted":fmt,"status":"crossref",
                                 "url":new_url.strip(),"_v":wv+1})
                        else:
                            st.session_state.references[i].update(
                                {"status":"info:"+status,"url":new_url.strip()})
                        st.rerun()
                    else:
                        st.warning("Veuillez entrer une URL ou un DOI.")
            with rc3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"rdel_{i}"):
                    ref_del.append(i); st.rerun()

            new_fmt = st.text_area(
                "Référence formatée",
                value=ref.get("formatted",""),
                key=f"rfmt_{i}_{wv}", height=75,
                placeholder="La référence apparaîtra ici après génération, ou saisissez-la manuellement…"
            )
            st.session_state.references[i]["formatted"] = new_fmt
            st.session_state.references[i]["url"]       = new_url

            status = ref.get("status","")
            if status == "crossref":
                st.markdown('<div class="success-box">✅ Référence générée via CrossRef</div>',
                            unsafe_allow_html=True)
            elif status.startswith("info:"):
                st.markdown(f'<div class="info-box">ℹ️ {status[5:]}</div>',
                            unsafe_allow_html=True)

    for idx in sorted(ref_del, reverse=True):
        st.session_state.references.pop(idx)

# ══════════════════════════════════════════════
# TAB 5 — ANNEXES
# ══════════════════════════════════════════════
with tab5:
    st.markdown("""<div class="info-box">
    Même logique que les sections — configurez les titres, figures et tableaux.<br>
    Chaque annexe démarre sur une nouvelle page avec sa propre numérotation (page 1, 2…).<br>
    Citez chaque annexe dans le corps du rapport (ex : <i>voir Annexe A</i>).
    </div>""", unsafe_allow_html=True)

    if "annexes" not in st.session_state:
        st.session_state.annexes = []

    ann_letters = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
    ca2, _ = st.columns([1.4, 5])
    with ca2:
        if st.button("➕ Ajouter une annexe", use_container_width=True):
            st.session_state.annexes.append(
                {"title":"","figures":0,"tables":0,"tbl_rows":3,"tbl_cols":3,
                 "fig_captions":[],"tbl_captions":[]})
            st.rerun()

    ann_del = []
    for i, ann in enumerate(st.session_state.annexes):
        letter = ann_letters[i] if i < len(ann_letters) else str(i+1)
        title  = ann.get("title","") or f"Annexe {letter}"
        with st.expander(f"Annexe {letter} — {title}", expanded=(i==0)):
            ac1, ac2 = st.columns([3, 0.6])
            with ac1:
                new_at = st.text_input("Titre de l'annexe", value=ann.get("title",""),
                                        key=f"at_{i}", placeholder="ex : Données brutes, Code Python…")
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

            new_atr, new_atc = ann.get("tbl_rows",3), ann.get("tbl_cols",3)
            if new_ant > 0:
                atr_c, atc_c = st.columns(2)
                with atr_c:
                    new_atr = st.number_input("Lignes", min_value=2, max_value=20,
                                               value=ann.get("tbl_rows",3), key=f"atr_{i}")
                with atc_c:
                    new_atc = st.number_input("Colonnes", min_value=1, max_value=10,
                                               value=ann.get("tbl_cols",3), key=f"atc2_{i}")

            new_afcaps, new_atcaps = [], []
            if new_anf > 0:
                st.caption("Légendes des figures")
                for fi in range(int(new_anf)):
                    ex  = (ann.get("fig_captions") or [])[fi] if fi < len(ann.get("fig_captions") or []) else ""
                    cap = st.text_input(f"  Figure {fi+1}", value=ex, key=f"afc_{i}_{fi}",
                                         placeholder="Description…")
                    new_afcaps.append(cap)
            if new_ant > 0:
                st.caption("Légendes des tableaux")
                for ti in range(int(new_ant)):
                    ex  = (ann.get("tbl_captions") or [])[ti] if ti < len(ann.get("tbl_captions") or []) else ""
                    cap = st.text_input(f"  Tableau {ti+1}", value=ex, key=f"atca_{i}_{ti}",
                                         placeholder="Description…")
                    new_atcaps.append(cap)

            st.session_state.annexes[i] = {
                "title": new_at, "figures": int(new_anf), "tables": int(new_ant),
                "tbl_rows": int(new_atr), "tbl_cols": int(new_atc),
                "fig_captions": new_afcaps, "tbl_captions": new_atcaps,
            }

    for idx in sorted(ann_del, reverse=True):
        st.session_state.annexes.pop(idx)

# ─────────────────────────────────────────────
# GÉNÉRATION
# ─────────────────────────────────────────────

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("---")
st.markdown('<div class="section-label" style="font-size:1.05rem;margin-bottom:.6rem;">'
            '🚀 Générer le rapport</div>', unsafe_allow_html=True)

cb1, cb2 = st.columns([2, 3])
with cb1:
    gen_clicked = st.button("⚡ Générer le rapport .docx",
                             type="primary", use_container_width=True)
with cb2:
    st.markdown("""<div class="warn-box">
    ⚠️ <b>Après ouverture dans Word :</b> Ctrl+A → F9 pour mettre à jour la table des matières,
    la liste des figures et la liste des tableaux.<br>
    Remplacez les zones de texte gris par votre rédaction.
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
            buf    = generate_poly_report(config)
            safe_t = (report_title.strip() or "Rapport_technique")[:50]\
                     .replace(" ","_").replace("/","-")
            st.success("✅ Rapport généré avec succès !")
            st.download_button(
                label="📥 Télécharger le .docx",
                data=buf,
                file_name=f"{safe_t}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"❌ Erreur lors de la génération : {e}")
            st.exception(e)
