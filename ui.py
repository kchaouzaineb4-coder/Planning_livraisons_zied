"""Composants d'interface : thème, formats de nombres, en-têtes, KPIs, tableaux."""
import html
import re

import pandas as pd
import streamlit as st

NNBSP = " "  # espace fine insécable (séparateur de milliers FR)

# =====================================================
# FORMATS DE NOMBRES (affichage écran, format français)
# =====================================================
def _vide(x):
    try:
        return x is None or pd.isna(x) or str(x).strip() in ("", "nan", "None")
    except (TypeError, ValueError):
        return False


def fmt_num(x, dec=1, unite=""):
    if _vide(x):
        return "—"
    try:
        x = float(x)
    except (TypeError, ValueError):
        return str(x)
    s = f"{x:,.{dec}f}".replace(",", "§").replace(".", ",").replace("§", NNBSP)
    return f"{s}{NNBSP}{unite}" if unite else s


def fmt_kg(x):
    return fmt_num(x, 1, "kg")


def fmt_m3(x):
    return fmt_num(x, 3, "m³")


def fmt_pct(x):
    return fmt_num(x, 1, "%")


def fmt_int(x):
    return fmt_num(x, 0)


def nettoyer_message(msg):
    """Retire les emojis de tête des messages du backend (✅, ❌, ⚠️…)."""
    return re.sub(r"^[^\w«(\"']+", "", str(msg or "")).strip()


# =====================================================
# STYLE GLOBAL
# =====================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root{--ink:#0F172A;--muted:#64748B;--line:#E2E8F0;--soft:#F1F5F9;--primary:#0369A1;--primary-50:#EFF6FB;
--ok:#15803D;--ok-50:#F0FDF4;--warn:#B45309;--warn-50:#FFFBEB;--danger:#B91C1C;--danger-50:#FEF2F2;}
html,body,.stApp,button,input,textarea,select{font-family:'Inter','Segoe UI',system-ui,sans-serif;}
.block-container{padding-top:1.6rem;padding-bottom:3rem;max-width:1360px;}
header[data-testid="stHeader"]{background:transparent;}
#MainMenu,footer{visibility:hidden;}
h1,h2,h3{letter-spacing:-.01em;}

/* En-tête de page */
.ph{display:flex;flex-direction:column;gap:2px;margin:0 0 1.1rem 0;padding-bottom:1rem;border-bottom:1px solid var(--line);}
.ph-step{font-size:.75rem;font-weight:600;color:var(--primary);text-transform:uppercase;letter-spacing:.08em;}
.ph-title{font-size:1.65rem;font-weight:700;color:var(--ink);line-height:1.25;}
.ph-sub{font-size:.95rem;color:var(--muted);}

/* Titre de section */
.sec{margin:1.4rem 0 .6rem 0;}
.sec-title{font-size:1.08rem;font-weight:650;color:var(--ink);}
.sec-sub{font-size:.85rem;color:var(--muted);margin-top:2px;}

/* KPIs */
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:.4rem 0 1rem 0;}
.kpi{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 16px;}
.kpi-label{font-size:.78rem;color:var(--muted);font-weight:500;line-height:1.3;}
.kpi-value{font-size:1.45rem;font-weight:700;color:var(--ink);margin-top:4px;font-variant-numeric:tabular-nums;white-space:nowrap;}
.kpi-sub{font-size:.75rem;color:var(--muted);margin-top:2px;}
.kpi.ok{border-left:4px solid var(--ok);} .kpi.warn{border-left:4px solid #F59E0B;}
.kpi.danger{border-left:4px solid var(--danger);} .kpi.danger .kpi-value{color:var(--danger);}
.kpi.primary{border-left:4px solid var(--primary);}

/* Tableaux */
.dt-wrap{border:1px solid var(--line);border-radius:12px;overflow:auto;background:#fff;margin:.3rem 0 .8rem 0;}
.dt{width:100%;border-collapse:separate;border-spacing:0;font-size:13px;}
.dt thead th{position:sticky;top:0;z-index:1;background:var(--soft);color:#334155;font-weight:600;font-size:12px;
text-align:left;padding:10px 12px;border-bottom:1px solid var(--line);white-space:nowrap;}
.dt td{padding:9px 12px;border-bottom:1px solid #EEF2F6;vertical-align:top;color:var(--ink);line-height:1.45;}
.dt tbody tr:last-child td{border-bottom:none;}
.dt tbody tr:hover td{background:#F8FAFC;}
.dt th.num,.dt td.num{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums;}
.dt td.nowrap{white-space:nowrap;}
.dt tr.r-src td{background:var(--danger-50);} .dt tr.r-dst td{background:var(--ok-50);}
.dt tr.r-alert td:first-child{box-shadow:inset 3px 0 0 var(--danger);}
.dt .muted{color:var(--muted);}

/* Badges et barres d'occupation */
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11.5px;font-weight:600;white-space:nowrap;
background:var(--soft);color:#334155;}
.badge.blue{background:var(--primary-50);color:var(--primary);} .badge.green{background:#DCFCE7;color:var(--ok);}
.badge.amber{background:#FEF3C7;color:var(--warn);} .badge.red{background:#FEE2E2;color:var(--danger);}
.occ{display:flex;align-items:center;gap:8px;min-width:120px;}
.occ-bar{flex:1;height:6px;background:#E2E8F0;border-radius:99px;overflow:hidden;}
.occ-bar span{display:block;height:100%;background:var(--primary);border-radius:99px;}
.occ-bar span.over{background:var(--danger);}
.occ-val{font-size:12px;font-variant-numeric:tabular-nums;white-space:nowrap;min-width:48px;text-align:right;}
.occ-val.over{color:var(--danger);font-weight:600;}

/* Jauges */
.jg{background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px 14px;margin:.4rem 0 .8rem 0;}
.jg-t{font-size:.82rem;font-weight:600;color:var(--ink);margin-bottom:8px;}
.jg-g{display:grid;grid-template-columns:auto 1fr;gap:6px 12px;align-items:center;}
.jg-l{font-size:.8rem;color:var(--muted);}

/* État vide */
.empty{border:1px dashed #CBD5E1;border-radius:12px;padding:28px;text-align:center;background:#fff;color:var(--muted);}
.empty-title{font-weight:600;color:var(--ink);font-size:1rem;margin-bottom:4px;}

/* Barre latérale */
[data-testid="stSidebar"]{background:#FFFFFF;border-right:1px solid var(--line);}
[data-testid="stSidebar"] .stButton>button{justify-content:flex-start;}
.brand{display:flex;align-items:center;gap:10px;margin:.2rem 0 .2rem 0;}
.brand img{width:40px;height:40px;object-fit:contain;}
.brand-name{font-weight:700;color:var(--ink);font-size:1rem;line-height:1.2;}
.brand-sub{font-size:.78rem;color:var(--muted);}
.side-label{font-size:.72rem;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin:.6rem 0 .4rem 0;}
.side-stats{display:grid;grid-template-columns:1fr auto;row-gap:6px;font-size:.85rem;color:#334155;}
.side-stats b{font-variant-numeric:tabular-nums;color:var(--ink);text-align:right;}
.side-foot{font-size:.75rem;color:var(--muted);line-height:1.5;}

/* Cartes de fichiers */
.up-title{font-weight:650;color:var(--ink);font-size:.98rem;}
.up-desc{font-size:.8rem;color:var(--muted);margin:2px 0 6px 0;min-height:2.7em;line-height:1.35;}
[data-testid="stFileUploader"] section{border-radius:10px;}

@media (max-width:768px){.ph-title{font-size:1.3rem;}.kpi-value{font-size:1.2rem;}}
</style>
"""


def appliquer_style():
    st.markdown(CSS, unsafe_allow_html=True)


# =====================================================
# COMPOSANTS
# =====================================================
def _h(x):
    return html.escape(str(x))


def page_header(titre, sous_titre="", etape=None, total=4):
    step = f'<div class="ph-step">Étape {etape} sur {total}</div>' if etape else ""
    sub = f'<div class="ph-sub">{_h(sous_titre)}</div>' if sous_titre else ""
    st.markdown(f'<div class="ph">{step}<div class="ph-title">{_h(titre)}</div>{sub}</div>',
                unsafe_allow_html=True)


def section(titre, sous_titre=None):
    sub = f'<div class="sec-sub">{_h(sous_titre)}</div>' if sous_titre else ""
    st.markdown(f'<div class="sec"><div class="sec-title">{_h(titre)}</div>{sub}</div>',
                unsafe_allow_html=True)


def kpis(items):
    """items : liste de tuples (libellé, valeur[, sous-texte[, ton]]). Ton : ok, warn, danger, primary."""
    cartes = []
    for it in items:
        label, valeur = it[0], it[1]
        sub = it[2] if len(it) > 2 and it[2] else ""
        ton = it[3] if len(it) > 3 and it[3] else ""
        sub_html = f'<div class="kpi-sub">{_h(sub)}</div>' if sub else ""
        cartes.append(f'<div class="kpi {ton}"><div class="kpi-label">{_h(label)}</div>'
                      f'<div class="kpi-value">{_h(valeur)}</div>{sub_html}</div>')
    st.markdown(f'<div class="kpi-grid">{"".join(cartes)}</div>', unsafe_allow_html=True)


def badge(texte, ton=""):
    return f'<span class="badge {ton}">{_h(texte)}</span>'


def occupation(pct):
    if _vide(pct):
        return '<span class="muted">—</span>'
    p = float(pct)
    over = " over" if p > 100 else ""
    return (f'<div class="occ"><div class="occ-bar"><span class="{over.strip()}" style="width:{min(max(p, 0), 100):.0f}%">'
            f'</span></div><span class="occ-val{over}">{fmt_pct(p)}</span></div>')


def jauges(titre, items):
    """Petit bloc de jauges : items = [(libellé, pourcentage), ...]."""
    lignes = "".join(f'<span class="jg-l">{_h(l)}</span>{occupation(p)}' for l, p in items)
    st.markdown(f'<div class="jg"><div class="jg-t">{_h(titre)}</div><div class="jg-g">{lignes}</div></div>',
                unsafe_allow_html=True)


def etat_vide(titre, texte=""):
    st.markdown(f'<div class="empty"><div class="empty-title">{_h(titre)}</div>{_h(texte)}</div>',
                unsafe_allow_html=True)


SEP_LISTE = r"\s*(?:<br\s*/?>|;|\n|\|\|)\s*"   # BL, désignations
SEP_VIRGULE = r"\s*(?:,|;|\n)\s*"               # articles, clients, représentants


def tableau(df, multiligne=None, numeriques=(), html_cols=(), classes_lignes=None,
            hauteur_max=520, nowrap=()):
    """Tableau HTML lisible (en-tête figé, défilement, cellules multi-lignes).

    multiligne : {colonne: regex de séparation} -> un élément par ligne dans la cellule.
    numeriques : colonnes alignées à droite (valeurs déjà formatées).
    html_cols  : colonnes contenant du HTML déjà construit (badges, barres).
    classes_lignes : liste de classes CSS par ligne (même ordre que df).
    """
    multiligne = multiligne or {}
    cols = list(df.columns)
    th = "".join(f'<th class="{"num" if c in numeriques else ""}">{_h(c)}</th>' for c in cols)
    lignes = []
    for i, (_, row) in enumerate(df.iterrows()):
        tds = []
        for c in cols:
            v = row[c]
            cls = "num" if c in numeriques else ("nowrap" if c in nowrap else "")
            if c in html_cols:
                contenu = "" if _vide(v) else str(v)
            elif _vide(v):
                contenu = '<span class="muted">—</span>'
            elif c in multiligne:
                elems = [e for e in re.split(multiligne[c], str(v)) if e.strip() and e.strip() != "nan"]
                contenu = "<br>".join(_h(e.strip()) for e in elems) or '<span class="muted">—</span>'
            else:
                contenu = _h(v)
            tds.append(f'<td class="{cls}">{contenu}</td>')
        rc = classes_lignes[i] if classes_lignes else ""
        lignes.append(f'<tr class="{rc}">{"".join(tds)}</tr>')
    style = f' style="max-height:{hauteur_max}px"' if hauteur_max else ""
    st.markdown(f'<div class="dt-wrap"{style}><table class="dt"><thead><tr>{th}</tr></thead>'
                f'<tbody>{"".join(lignes)}</tbody></table></div>', unsafe_allow_html=True)


def filtre_texte(df, key, placeholder="Rechercher (client, BL, ville…)"):
    """Champ de recherche plein texte sur toutes les colonnes."""
    q = st.text_input("Rechercher", key=key, placeholder=placeholder,
                      label_visibility="collapsed", icon=":material/search:")
    if not q:
        return df
    q = q.strip().lower()
    masque = df.astype(str).apply(lambda col: col.str.lower().str.contains(q, regex=False)).any(axis=1)
    return df[masque]


# =====================================================
# MESSAGES PERSISTANTS (affichés après un rerun)
# =====================================================
def flash(message, type_="success"):
    st.session_state["_flash"] = (type_, nettoyer_message(message))


def afficher_flash():
    f = st.session_state.pop("_flash", None)
    if not f:
        return
    type_, msg = f
    icones = {"success": ":material/check_circle:", "error": ":material/error:",
              "warning": ":material/warning:", "info": ":material/info:"}
    getattr(st, type_, st.info)(msg, icon=icones.get(type_))
    st.toast(msg, icon=icones.get(type_))  # visible même si l'utilisateur est en bas de page
