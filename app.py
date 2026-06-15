"""
Générateur de rapport technique — Polytechnique Montréal
Interface Streamlit · Par Najeeb Lababidi
"""

import streamlit as st
from generate_report import generate_poly_report, FONTS_AVAILABLE, REF_STYLES
from datetime import date
import requests, json, base64, os

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
  :root { --red:#CC0000; --dark:#1a1a1a; --gray:#595959; }

  /* Header principal */
  .hero {
    background: linear-gradient(135deg,#CC0000 0%,#8B0000 100%);
    padding:1.6rem 2rem 1.2rem; border-radius:12px;
    margin-bottom:1.4rem; color:white;
  }
  .hero h1 { color:white; margin:0; font-size:1.55rem; font-weight:700; }
  .hero .sub { color:rgba(255,255,255,.82); font-size:.88rem; margin:.3rem 0 0; }
  .hero .credit { color:rgba(255,255,255,.65); font-size:.78rem; margin:.6rem 0 0; }

  /* Badges */
  .badge {
    background:#CC0000; color:white;
    padding:.14rem .55rem; border-radius:10px;
    font-size:.72rem; font-weight:700; margin-right:.4rem;
  }
  .badge-gray {
    background:#555; color:white;
    padding:.14rem .55rem; border-radius:10px;
    font-size:.72rem; font-weight:700; margin-right:.4rem;
  }

  /* Boxes */
  .info-box {
    background:#f8f8f8; border-left:3px solid #CC0000;
    padding:.55rem .9rem; border-radius:0 6px 6px 0;
    margin:.4rem 0; font-size:.82rem; color:#555;
  }
  .warn-box {
    background:#fff8e1; border-left:3px solid #f9a825;
    padding:.55rem .9rem; border-radius:0 6px 6px 0;
    margin:.4rem 0; font-size:.82rem; color:#666;
  }
  .success-box {
    background:#e8f5e9; border-left:3px solid #2e7d32;
    padding:.55rem .9rem; border-radius:0 6px 6px 0;
    margin:.4rem 0; font-size:.82rem; color:#2e7d32;
  }

  /* Tabs */
  div[data-testid="stTabs"] button[aria-selected="true"] {
    border-bottom:3px solid #CC0000 !important;
    color:#CC0000 !important; font-weight:600;
  }

  /* Sidebar */
  section[data-testid="stSidebar"] { background:#fafafa; }

  /* Bouton primaire */
  div.stButton > button[kind="primary"] {
    background:#CC0000 !important; color:white !important;
    border:none !important; border-radius:6px !important;
    font-weight:600 !important;
  }
  div.stButton > button[kind="primary"]:hover {
    background:#aa0000 !important;
  }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HEADER PERSONNALISÉ
# ─────────────────────────────────────────────

_DIR = os.path.dirname(os.path.abspath(__file__))
_LOGO = os.path.join(_DIR, "poly_logo.png")

col_logo, col_txt = st.columns([1, 4])
with col_logo:
    if os.path.exists(_LOGO):
        st.image(_LOGO, width=160)
with col_txt:
    st.markdown("""
    <div class="hero">
      <h1>📄 PolyReport — Générateur de rapports techniques</h1>
      <div class="sub">Polytechnique Montréal &nbsp;·&nbsp; Gabarit S02 Rev.F &nbsp;·&nbsp; Zéro installation requise</div>
      <div class="credit">
        Développé par <strong>Najeeb Lababidi</strong> · Étudiant en génie aérospatial, Polytechnique Montréal<br>
        Outil gratuit pour les étudiants · Génère le squelette Word — tu écris ton texte directement dans le .docx
      </div>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR — MISE EN FORME
# ─────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ Mise en forme")
    st.markdown("---")

    font = st.selectbox("Police", FONTS_AVAILABLE, index=0,
                        help="Gabarit officiel Poly : Cambria")
    font_size = st.slider("Taille du corps (pt)", 10, 14, 12)
    ref_style = st.selectbox("Style de références", REF_STYLES, index=0,
                             help="IEEE recommandé en génie")

    st.markdown("---")
    st.markdown("### 💡 Comment utiliser")
    st.markdown("""
    <div class="info-box">
    <b>1.</b> Remplis le formulaire<br>
    <b>2.</b> Génère le .docx<br>
    <b>3.</b> Ouvre dans Word<br>
    <b>4.</b> Ctrl+A → F9 (TDM)<br>
    <b>5.</b> Écris ton texte dans les zones grises
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style="font-size:.73rem;color:#aaa;text-align:center">
    Outil gratuit pour les étudiants de Poly<br>
    © Najeeb Lababidi · Génie aérospatial<br>
    <a href="mailto:najeeb.lababidi@etud.polymtl.ca" style="color:#CC0000">Signaler un problème</a>
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
    st.markdown('<span class="badge">DOC</span> **Informations du document**',
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        doc_number = st.text_input("Numéro de document", value="MEC3520 DEV4",
                                    help="Laisser vide si non applicable")
        revision   = st.text_input("Révision", value="--")
    with c2:
        doc_date_raw = st.date_input("Date du document", value=date.today())
        # Format "9 avril 2026" comme ton rapport
        MONTHS_FR = {1:"janvier",2:"février",3:"mars",4:"avril",5:"mai",6:"juin",
                     7:"juillet",8:"août",9:"septembre",10:"octobre",11:"novembre",12:"décembre"}
        doc_date = f"{doc_date_raw.day} {MONTHS_FR[doc_date_raw.month]} {doc_date_raw.year}"
        semester = st.text_input("Session", value="Hiver 2026")
    with c3:
        group_number = st.text_input("Groupe-cours n°", value="02",
                                      help="Laisser vide si non applicable")

    st.markdown("---")
    st.markdown('<span class="badge">COURS</span> **Cours**', unsafe_allow_html=True)
    cc1, cc2 = st.columns(2)
    with cc1: course_code = st.text_input("Sigle", value="MEC 3520")
    with cc2: course_name = st.text_input("Titre du cours", value="Industrialisation des produits")

    st.markdown("---")
    st.markdown('<span class="badge">RAPPORT</span> **Titre du rapport**',
                unsafe_allow_html=True)
    report_title = st.text_input("Titre", value="Titre de votre rapport technique")

    st.markdown("---")
    st.markdown('<span class="badge">ÉQUIPE</span> **Encadrement et auteur(s)**',
                unsafe_allow_html=True)

    professors_raw = st.text_input("Présenté à (professeur(s), séparés par virgule)",
                                    value="Nom du professeur",
                                    help="Ex: Marwan Azzi  ou  Prof. Morin, Dr. Ahmadi")
    professors = [p.strip() for p in professors_raw.split(",") if p.strip()]

    is_individual = st.toggle("🎓 Travail individuel (pas de tableau signatures)",
                               value=False)

    if is_individual:
        st.markdown('<span class="badge-gray">1 étudiant</span> **Vos informations**',
                    unsafe_allow_html=True)
        st.markdown("""<div class="info-box">Format : <code>Nom Prénom | Matricule</code></div>""",
                    unsafe_allow_html=True)
        m_raw = st.text_input("Nom et matricule", value="Nom Prénom | 2XXXXXX",
                               label_visibility="collapsed")
        if "|" in m_raw:
            parts = m_raw.split("|",1)
            members = [{"name": parts[0].strip(), "matricule": parts[1].strip()}]
        else:
            members = [{"name": m_raw.strip(), "matricule": ""}]
        team_number = ""
    else:
        st.markdown('<span class="badge-gray">Équipe</span> **Membres** '
                    '(format : `Nom Prénom | Matricule`, un par ligne)',
                    unsafe_allow_html=True)
        members_raw = st.text_area("Membres", label_visibility="collapsed", height=120,
                                    value="Najeeb Lababidi | 2369834\nVictor Béland | 2381714")
        members = []
        for line in members_raw.strip().split("\n"):
            line = line.strip()
            if not line: continue
            if "|" in line:
                p2 = line.split("|",1)
                members.append({"name":p2[0].strip(),"matricule":p2[1].strip()})
            else:
                members.append({"name":line,"matricule":""})

# ══════════════════════════════════════════════
# TAB 2 — RÉSUMÉ & ABSTRACT
# ══════════════════════════════════════════════
with tab2:
    st.markdown("""<div class="info-box">
    Le résumé <b>ne dépasse pas une page</b>. Il contient : mise en contexte, objectif,
    principaux résultats, prochaines actions.<br>
    L'abstract est une traduction fidèle du résumé.
    </div>""", unsafe_allow_html=True)

    st.markdown('<span class="badge">FR</span> **Résumé en français**',
                unsafe_allow_html=True)
    resume_fr = st.text_area("Résumé", label_visibility="collapsed", height=170,
                              placeholder="Coller ou écrire votre résumé ici…")
    keywords_fr = st.text_input("Mots clés FR (séparés par virgules)",
                                 placeholder="ex: turbofan, propulsion, efficacité")

    st.markdown("---")
    st.markdown('<span class="badge">EN</span> **Abstract**', unsafe_allow_html=True)

    col_abs, col_btn = st.columns([4, 1])
    with col_abs:
        abstract_en = st.text_area("Abstract", label_visibility="collapsed", height=170,
                                    placeholder="Paste or write your abstract here…",
                                    key="abstract_field")
    with col_btn:
        st.markdown("<br><br>", unsafe_allow_html=True)
        translate_clicked = st.button("🤖 Traduire\nle résumé",
                                       help="Traduit automatiquement votre résumé FR → EN avec Claude AI",
                                       use_container_width=True)

    if translate_clicked:
        if not resume_fr.strip():
            st.warning("Écris d'abord ton résumé en français.")
        else:
            # Récupérer la clé API depuis les secrets Streamlit
            api_key = None
            try:
                api_key = st.secrets.get("ANTHROPIC_API_KEY", None)
            except Exception:
                pass

            if not api_key:
                st.markdown("""<div class="warn-box">
                <b>Clé API non configurée.</b><br>
                Pour activer la traduction automatique sur Streamlit Cloud :<br>
                <b>Settings → Secrets</b> → ajouter <code>ANTHROPIC_API_KEY = "sk-ant-..."</code><br><br>
                En attendant, voici une traduction mot à mot du résumé que tu peux raffiner :
                </div>""", unsafe_allow_html=True)
                # Afficher le résumé pour copier-coller manuel
                st.text_area("Ton résumé FR à traduire manuellement :",
                             value=resume_fr, height=120, key="resume_copy")
            else:
                with st.spinner("Traduction en cours avec Claude AI…"):
                    try:
                        resp = requests.post(
                            "https://api.anthropic.com/v1/messages",
                            headers={
                                "Content-Type": "application/json",
                                "x-api-key": api_key,
                                "anthropic-version": "2023-06-01",
                            },
                            json={
                                "model": "claude-sonnet-4-20250514",
                                "max_tokens": 1000,
                                "messages": [{
                                    "role": "user",
                                    "content": (
                                        "Tu es un traducteur technique spécialisé en génie. "
                                        "Traduis fidèlement ce résumé du français vers l'anglais académique. "
                                        "Conserve exactement la même structure, les mêmes données chiffrées, "
                                        "et le même registre technique. Réponds UNIQUEMENT avec la traduction, "
                                        "sans introduction ni commentaire.\n\n"
                                        f"RÉSUMÉ FR :\n{resume_fr}"
                                    )
                                }]
                            },
                            timeout=30
                        )
                        if resp.status_code == 200:
                            translated = resp.json()["content"][0]["text"].strip()
                            st.session_state["abstract_field"] = translated
                            st.rerun()
                        elif resp.status_code == 401:
                            st.error("Clé API invalide. Vérifier dans Settings → Secrets.")
                        else:
                            st.error(f"Erreur API ({resp.status_code}) — traduis manuellement.")
                    except requests.exceptions.Timeout:
                        st.error("Délai dépassé — réessaie dans quelques secondes.")
                    except Exception as e:
                        st.error(f"Erreur de connexion : {e}")

    keywords_en = st.text_input("Keywords EN (comma-separated)",
                                 placeholder="ex: turbofan, propulsion, efficiency")

# ══════════════════════════════════════════════
# TAB 3 — SECTIONS
# ══════════════════════════════════════════════
with tab3:
    st.markdown("""<div class="info-box">
    Configure la <b>structure</b> du rapport (titres, niveaux, nombre de figures/tableaux par section).<br>
    Le .docx généré contiendra des <b>zones grises</b> avec les styles Word corrects — tu écris ton texte directement dans Word.<br>
    La table des matières, liste des figures et liste des tableaux se mettent à jour automatiquement (<b>Ctrl+A → F9</b>).
    </div>""", unsafe_allow_html=True)

    DEFAULT_SECTIONS = [
        {"title":"Introduction",        "level":1,"figures":0,"tables":0,"fig_captions":[],"tbl_captions":[]},
        {"title":"Mise en contexte",    "level":2,"figures":0,"tables":0,"fig_captions":[],"tbl_captions":[]},
        {"title":"Objectifs",           "level":2,"figures":0,"tables":0,"fig_captions":[],"tbl_captions":[]},
        {"title":"Développement",       "level":1,"figures":0,"tables":0,"fig_captions":[],"tbl_captions":[]},
        {"title":"Résultats",           "level":1,"figures":0,"tables":0,"fig_captions":[],"tbl_captions":[]},
        {"title":"Analyse et discussion","level":1,"figures":0,"tables":0,"fig_captions":[],"tbl_captions":[]},
        {"title":"Conclusion",          "level":1,"figures":0,"tables":0,"fig_captions":[],"tbl_captions":[]},
    ]
    if "sections" not in st.session_state:
        st.session_state.sections = DEFAULT_SECTIONS.copy()

    ca, cb, cc = st.columns([1,1,4])
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
    for i, sec in enumerate(st.session_state.sections):
        prefix = "§" * sec["level"]
        with st.expander(f"{prefix} {sec['title']}", expanded=False):
            r1c1, r1c2, r1c3 = st.columns([3, 1, 0.6])
            with r1c1:
                new_t = st.text_input("Titre", value=sec["title"], key=f"st_{i}")
            with r1c2:
                new_l = st.selectbox("Niveau", [1,2,3,4], index=sec["level"]-1,
                                      key=f"sl_{i}", format_func=lambda x: f"H{x}")
            with r1c3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"sd_{i}"):
                    to_del.append(i); st.rerun()

            r2c1, r2c2 = st.columns(2)
            with r2c1:
                new_nf = st.number_input("Figures dans cette section",
                                          min_value=0, max_value=10,
                                          value=sec.get("figures",0), key=f"snf_{i}")
            with r2c2:
                new_nt = st.number_input("Tableaux dans cette section",
                                          min_value=0, max_value=10,
                                          value=sec.get("tables",0), key=f"snt_{i}")

            # Légendes figures
            new_fcaps = []
            if new_nf > 0:
                st.markdown("**Légendes des figures** *(optionnel — à modifier dans Word)*")
                for fi in range(int(new_nf)):
                    existing = sec.get("fig_captions",[])[fi] \
                               if fi < len(sec.get("fig_captions",[])) else ""
                    cap = st.text_input(f"  Figure {fi+1}", value=existing,
                                         key=f"sfc_{i}_{fi}",
                                         placeholder="Description de la figure…")
                    new_fcaps.append(cap)

            # Légendes tableaux
            new_tcaps = []
            if new_nt > 0:
                st.markdown("**Légendes des tableaux** *(optionnel — à modifier dans Word)*")
                for ti in range(int(new_nt)):
                    existing = sec.get("tbl_captions",[])[ti] \
                               if ti < len(sec.get("tbl_captions",[])) else ""
                    cap = st.text_input(f"  Tableau {ti+1}", value=existing,
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
    <b>Coller le DOI ou l'URL</b> de chaque article → le moteur génère automatiquement
    la référence formatée en <b>IEEE / APA / Chicago</b> via l'API CrossRef.<br>
    Pour les URLs sans DOI (sites web, rapports), entre le texte directement.
    </div>""", unsafe_allow_html=True)

    if "references" not in st.session_state:
        st.session_state.references = [
            {"url":"https://doi.org/10.1016/j.heliyon.2024.e34433","formatted":"","status":""},
        ]

    def _format_ref(data, style):
        """Formate une référence CrossRef selon le style choisi."""
        authors = data.get("author", [])

        def fmt_author(a):
            family   = a.get("family", "")
            given    = a.get("given", "")
            if not family: return a.get("name", "")
            initials = " ".join(f"{n[0]}." for n in given.split()) if given else ""
            if style == "APA":   return f"{family}, {initials}".strip(", ")
            elif style == "IEEE": return f"{initials} {family}".strip()
            else:                 return f"{family}, {given}".strip(", ")

        if style == "APA":
            if not authors:           auth_str = "Auteur inconnu"
            elif len(authors) == 1:   auth_str = fmt_author(authors[0])
            elif len(authors) <= 7:
                auth_str = ", ".join(fmt_author(a) for a in authors[:-1])
                auth_str += f", & {fmt_author(authors[-1])}"
            else:
                auth_str = ", ".join(fmt_author(a) for a in authors[:6])
                auth_str += f", … {fmt_author(authors[-1])}"
        elif style == "IEEE":
            if not authors:           auth_str = "Auteur inconnu"
            elif len(authors) <= 3:   auth_str = ", ".join(fmt_author(a) for a in authors)
            else:                     auth_str = fmt_author(authors[0]) + " et al."
        else:  # Chicago
            if not authors:           auth_str = "Auteur inconnu"
            elif len(authors) == 1:   auth_str = fmt_author(authors[0])
            else:                     auth_str = fmt_author(authors[0]) + " et al."

        year = ""
        for f2 in ["published","published-print","published-online","issued"]:
            dp = data.get(f2, {}).get("date-parts", [[]])
            if dp and dp[0]: year = str(dp[0][0]); break

        titles    = data.get("title", [""]); title = titles[0] if titles else ""
        ct        = data.get("container-title", []); container = ct[0] if ct else ""
        volume    = data.get("volume", "")
        issue     = data.get("issue",  "")
        pages     = data.get("page",   "")
        doi_clean = data.get("DOI",    "")

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
        else:  # Chicago
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
        """
        Essaie CrossRef en premier (DOI).
        Si CrossRef échoue, utilise l'API Claude pour parser l'URL.
        """
        raw = doi_or_url.strip()

        # ── Étape 1 : extraire le DOI ──────────────
        doi = raw
        for prefix in ["https://doi.org/","http://doi.org/","doi.org/","DOI:","doi:"]:
            if doi.lower().startswith(prefix.lower()):
                doi = doi[len(prefix):].strip()
                break

        # ── Étape 2 : essayer CrossRef ─────────────
        # Seulement si ça ressemble à un DOI (contient un slash après "10.")
        looks_like_doi = doi.startswith("10.") and "/" in doi
        if looks_like_doi:
            try:
                r = requests.get(
                    f"https://api.crossref.org/works/{doi}",
                    headers={"User-Agent":
                             "PolyReport/1.0 (mailto:najeeb.lababidi@etud.polymtl.ca)"},
                    timeout=12
                )
                if r.status_code == 200:
                    formatted = _format_ref(r.json()["message"], ref_style)
                    return formatted, "crossref"
            except Exception:
                pass  # CrossRef indisponible → fallback Claude

        # ── Étape 3 : fallback Claude AI ──────────
        api_key = None
        try:
            api_key = st.secrets.get("ANTHROPIC_API_KEY", None)
        except Exception:
            pass

        if not api_key:
            return None, (
                "CrossRef inaccessible et clé API non configurée. "
                "Entre la référence manuellement, ou configure ANTHROPIC_API_KEY dans "
                "Settings → Secrets de Streamlit Cloud."
            )

        style_examples = {
            "APA":     "Altharan, Y. M., Shamsudin, S. (2024). Titre. Heliyon, 10(14), e34433. https://doi.org/xxx",
            "IEEE":    'Y. M. Altharan, S. Shamsudin, "Titre," Heliyon, vol. 10, no. 14, 2024. doi: xxx.',
            "Chicago": 'Altharan, Y. M. "Titre." Heliyon 10, no. 14 (2024): e34433. https://doi.org/xxx.',
        }
        prompt = (
            f"Tu es un assistant de bibliographie académique. "
            f"À partir de l'URL ou DOI suivant, génère la référence complète au format {ref_style}.\n"
            f"Exemple {ref_style} : {style_examples[ref_style]}\n"
            f"URL/DOI : {raw}\n\n"
            f"Réponds UNIQUEMENT avec la référence formatée, sans commentaire ni introduction. "
            f"Si tu ne peux pas accéder à l'URL, génère une référence plausible basée sur les "
            f"informations visibles dans l'URL elle-même."
        )
        try:
            resp = requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={
                    "Content-Type": "application/json",
                    "x-api-key": api_key,
                    "anthropic-version": "2023-06-01",
                },
                json={
                    "model": "claude-sonnet-4-20250514",
                    "max_tokens": 300,
                    "messages": [{"role":"user","content": prompt}]
                },
                timeout=20
            )
            if resp.status_code == 200:
                formatted = resp.json()["content"][0]["text"].strip()
                return formatted, "ai"
            else:
                return None, f"Erreur API Claude ({resp.status_code}) — entre la référence manuellement."
        except Exception as e:
            return None, f"Erreur réseau : {e}"

    col_add_ref, _ = st.columns([1,5])
    with col_add_ref:
        if st.button("➕ Référence"):
            st.session_state.references.append({"url":"","formatted":"","status":""})
            st.rerun()

    ref_del = []
    for i, ref in enumerate(st.session_state.references):
        with st.expander(f"Référence {i+1}" +
                         (f" — {ref['formatted'][:60]}…" if ref.get('formatted') else ""),
                         expanded=(not ref.get('formatted'))):
            rc1, rc2, rc3 = st.columns([3,1,0.6])
            with rc1:
                new_url = st.text_input("DOI / URL de l'article",
                                         value=ref.get("url",""), key=f"rurl_{i}",
                                         placeholder="https://doi.org/10.xxxx/xxxxx")
            with rc2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🔍 Générer", key=f"rfetch_{i}", use_container_width=True):
                    if new_url.strip():
                        with st.spinner("Recherche en cours…"):
                            fmt, status = fetch_reference(new_url.strip())
                            if fmt:
                                st.session_state.references[i]["formatted"] = fmt
                                st.session_state.references[i]["status"]    = status
                                st.session_state.references[i]["url"]       = new_url.strip()
                            else:
                                st.session_state.references[i]["status"] = "error:" + status
                        st.rerun()
            with rc3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"rdel_{i}"):
                    ref_del.append(i); st.rerun()

            new_fmt = st.text_area("Référence formatée (modifiable)",
                                    value=ref.get("formatted",""),
                                    key=f"rfmt_{i}", height=80,
                                    placeholder="La référence apparaîtra ici après génération, "
                                                "ou entrez-la manuellement…")
            st.session_state.references[i]["formatted"] = new_fmt
            st.session_state.references[i]["url"]       = new_url

            status = ref.get("status","")
            if status == "crossref":
                st.markdown('<div class="success-box">✅ Référence générée via CrossRef (base de données officielle)</div>',
                            unsafe_allow_html=True)
            elif status == "ai":
                st.markdown('<div class="success-box">🤖 Référence générée par IA — vérifier les détails avant de remettre</div>',
                            unsafe_allow_html=True)
            elif status.startswith("error:"):
                st.markdown(f'<div class="warn-box">⚠️ {status[6:]}</div>',
                            unsafe_allow_html=True)

    for idx in sorted(ref_del, reverse=True):
        st.session_state.references.pop(idx)

# ══════════════════════════════════════════════
# TAB 5 — ANNEXES
# ══════════════════════════════════════════════
with tab5:
    st.markdown("""<div class="info-box">
    Même logique que les sections — configure les titres, figures et tableaux.<br>
    Chaque annexe a sa propre numérotation de pages (repart à 1).<br>
    Citer chaque annexe dans le corps du rapport (ex: <i>voir Annexe A</i>).
    </div>""", unsafe_allow_html=True)

    if "annexes" not in st.session_state:
        st.session_state.annexes = [
            {"title":"Données supplémentaires","figures":0,"tables":1,
             "fig_captions":[],"tbl_captions":["Données brutes mesurées."]}
        ]

    ann_letters = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")

    ca2, _ = st.columns([1,5])
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
            ac1, ac2, ac3 = st.columns([3,1,0.6])
            with ac1:
                new_at = st.text_input("Titre", value=ann["title"], key=f"at_{i}")
            with ac3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"ad_{i}"):
                    ann_del.append(i); st.rerun()

            ac4, ac5 = st.columns(2)
            with ac4:
                new_anf = st.number_input("Figures", min_value=0, max_value=10,
                                           value=ann.get("figures",0), key=f"anf_{i}")
            with ac5:
                new_ant = st.number_input("Tableaux", min_value=0, max_value=10,
                                           value=ann.get("tables",0), key=f"ant_{i}")

            new_afcaps, new_atcaps = [], []
            if new_anf > 0:
                st.markdown("**Légendes figures**")
                for fi in range(int(new_anf)):
                    ex = ann.get("fig_captions",[])[fi] if fi<len(ann.get("fig_captions",[])) else ""
                    c  = st.text_input(f"  Figure {fi+1}", value=ex, key=f"afc_{i}_{fi}",
                                        placeholder="Description…")
                    new_afcaps.append(c)
            if new_ant > 0:
                st.markdown("**Légendes tableaux**")
                for ti in range(int(new_ant)):
                    ex = ann.get("tbl_captions",[])[ti] if ti<len(ann.get("tbl_captions",[])) else ""
                    c  = st.text_input(f"  Tableau {ti+1}", value=ex, key=f"atc_{i}_{ti}",
                                        placeholder="Description…")
                    new_atcaps.append(c)

            st.session_state.annexes[i] = {
                "title":new_at, "figures":int(new_anf), "tables":int(new_ant),
                "fig_captions":new_afcaps, "tbl_captions":new_atcaps,
            }

    for idx in sorted(ann_del, reverse=True):
        st.session_state.annexes.pop(idx)

# ─────────────────────────────────────────────
# GÉNÉRATION
# ─────────────────────────────────────────────

st.markdown("---")
cb1, cb2 = st.columns([2,3])
with cb1:
    gen_clicked = st.button("⚡ Générer le rapport .docx",
                             type="primary", use_container_width=True)
with cb2:
    st.markdown("""<div class="warn-box">
    ⚠️ <b>Après ouverture dans Word :</b> Ctrl+A → F9 pour mettre à jour
    la table des matières, la liste des figures et la liste des tableaux.<br>
    Les zones de texte grises sont des <b>placeholders</b> — les remplacer par votre rédaction.
    </div>""", unsafe_allow_html=True)

if gen_clicked:
    # Construire les refs (seulement celles avec texte formaté)
    refs_clean = [
        {"formatted": r.get("formatted",""), "url": r.get("url","")}
        for r in st.session_state.get("references",[])
        if r.get("formatted","").strip()
    ]

    config = {
        "font":          font,
        "font_size":     font_size,
        "ref_style":     ref_style,
        "doc_number":    doc_number,
        "revision":      revision,
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
        "abstract_en":   st.session_state.get("abstract_field",""),
        "keywords_fr":   keywords_fr,
        "keywords_en":   keywords_en,
        "sections":      st.session_state.get("sections",[]),
        "references":    refs_clean,
        "annexes":       st.session_state.get("annexes",[]),
    }

    with st.spinner("Génération du rapport en cours…"):
        try:
            buf      = generate_poly_report(config)
            safe_t   = report_title[:40].replace(" ","_").replace("/","-")
            filename = f"{doc_number.replace(' ','_')}_{safe_t}.docx"
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
