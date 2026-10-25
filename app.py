"""Planning des livraisons — préparation des chargements ADV (estafettes et camions loués)."""
import base64
import copy
import io
import json
import os
import re
import tempfile
from datetime import datetime, timedelta, timezone

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Planning Livraisons",
    page_icon=":material/local_shipping:",
    layout="wide",
    initial_sidebar_state="expanded",
)

import exports as xp  # noqa: E402
import metier as m  # noqa: E402
from backend import (SEUIL_POIDS, SEUIL_VOLUME, DeliveryProcessor,  # noqa: E402
                     TruckRentalProcessor, TruckTransferManager, get_capacite_poids_camion,
                     get_capacite_volume_camion)
from ui import (SEP_LISTE, SEP_VIRGULE, afficher_flash, appliquer_style, badge,  # noqa: E402
                etat_vide, filtre_texte, flash, jauges, fmt_int, fmt_kg, fmt_m3, fmt_num, fmt_pct, kpis,
                nettoyer_message, occupation, page_header, section, tableau)

appliquer_style()
S = st.session_state
TZ = timezone(timedelta(hours=1), "Africa/Tunis")  # Tunisie : UTC+1 toute l'année
BLEU = "#0369A1"
TAUX = "Taux d'occupation (%)"

# =====================================================
# ÉTAT DE SESSION
# =====================================================
ETAT_INITIAL = {
    "page": "import",
    "data_processed": False,
    "df_grouped": None, "df_city": None, "df_grouped_zone": None, "df_zone": None,
    "df_optimized_estafettes": None, "df_livraisons_original": None, "df_livraisons": None,
    "rental_processor": None, "propositions": None, "decisions": [],
    "df_voyages": None, "df_voyages_valides": None,
    "validations": {}, "attributions": {},
    "objets_historique": [], "transferts_historique": [],
    "df_poids_suspects": None, "df_volumes_nuls": None,
    "fichiers": [], "ver_editeur": 0, "ver_upload": 0,
    "opt_vue": "Location", "export_planning": None,
}
for _k, _v in ETAT_INITIAL.items():
    if _k not in S:
        S[_k] = copy.deepcopy(_v)


def aller(page, vue=None):
    if S.get("page") != page:
        S["_haut_de_page"] = True
    S.page = page
    if vue:
        S.opt_vue = vue
    elif page == "optimisation" and (S.propositions is None or S.propositions.empty):
        S.opt_vue = "Plan de chargement"  # plus rien à arbitrer : on ouvre directement le plan


def actualiser_propositions():
    rp = S.rental_processor
    S.propositions = rp.detecter_propositions() if rp else pd.DataFrame()


def reconstruire_plan():
    """Recalcule les voyages à partir des décisions de location (repart d'un plan propre)."""
    S.df_voyages = m.ajouter_taux(S.rental_processor.get_df_result())
    S.df_voyages_valides = None
    S.validations, S.attributions = {}, {}
    S.objets_historique, S.transferts_historique = [], []
    S.export_planning = None
    S.ver_editeur += 1
    actualiser_propositions()


def plan_modifie():
    return bool(S.transferts_historique or S.objets_historique or S.validations)


def invalider_validation():
    """Le plan a changé : la liste validée doit être refaite."""
    if S.df_voyages_valides is not None:
        S.df_voyages_valides = None
        S.export_planning = None
        return True
    return False


def cle_vehicule(v):
    mt = re.match(r"([A-Za-z]+)(\d+)", str(v))
    return (mt.group(1), int(mt.group(2)), str(v)) if mt else (str(v), 0, str(v))


def zones_triees(df):
    return sorted(df["Zone"].dropna().unique().tolist(), key=m.num_zone)


def bouton_telechargement(label, data, nom, mime=xp.MIME_XLSX, key=None, primaire=False):
    st.download_button(label, data=data, file_name=nom, mime=mime, key=key, width="stretch",
                       icon=":material/download:", type="primary" if primaire else "secondary")


def garde():
    if S.data_processed:
        return True
    etat_vide("Aucune donnée chargée", "Importez les fichiers du jour pour accéder à cette étape.")
    st.button("Aller à l'importation", type="primary", icon=":material/upload_file:",
              on_click=aller, args=("import",))
    return False


def style_fig(fig, titre, hauteur=330):
    fig.update_layout(
        title=dict(text=titre, x=0, font=dict(size=14, color="#0F172A")),
        template="plotly_white", height=hauteur, margin=dict(l=8, r=8, t=46, b=8),
        font=dict(family="Inter, Segoe UI, sans-serif", size=12, color="#334155"),
        xaxis_title=None, yaxis_title=None, showlegend=False, bargap=0.35,
    )
    fig.update_yaxes(gridcolor="#EEF2F6", zeroline=False)
    fig.update_xaxes(showgrid=False)
    return fig


def graphique(fig):
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})


def table_voyages(df, classes=None, hauteur=560):
    """Tableau standard d'un plan de voyages (occupation, type, BL)."""
    if df is None or df.empty:
        etat_vide("Aucun voyage à afficher")
        return
    vue = pd.DataFrame({
        "Zone": df["Zone"],
        "Véhicule": df["Véhicule N°"],
        "Type": df.apply(lambda r: badge(m.libelle_type(r), "amber" if m.est_camion(r) else "blue"), axis=1),
        "Clients": df["Client(s) inclus"],
        "Représentants": df["Représentant(s) inclus"],
        "BL": df["BL inclus"],
        "Poids": df["Poids total chargé"].map(fmt_kg),
        "Volume": df["Volume total chargé"].map(fmt_m3),
        "Occupation poids": df["Taux occupation poids (%)"].map(occupation),
        "Occupation volume": df["Taux occupation volume (%)"].map(occupation),
    })
    if classes is None:
        classes = ["r-alert" if t > 100 else "" for t in df["Taux d'occupation (%)"]]
    tableau(vue, multiligne={"Zone": SEP_LISTE, "Clients": SEP_VIRGULE, "Représentants": SEP_VIRGULE,
                             "BL": SEP_LISTE},
            numeriques=("Poids", "Volume"), html_cols=("Type", "Occupation poids", "Occupation volume"),
            classes_lignes=classes, hauteur_max=hauteur, nowrap=("Véhicule", "BL"))


# =====================================================
# SAUVEGARDE / REPRISE DE SESSION (fichier .json)
# =====================================================
CLES_DF = ["df_grouped", "df_city", "df_grouped_zone", "df_zone", "df_optimized_estafettes",
           "df_livraisons_original", "df_voyages", "df_voyages_valides", "df_poids_suspects", "df_volumes_nuls"]
CLES_OBJ = ["validations", "attributions", "decisions", "objets_historique", "transferts_historique", "fichiers"]


def exporter_session():
    rp = S.rental_processor
    data = {
        "format": "planning-livraisons", "version": 1, "date": datetime.now(TZ).strftime("%d/%m/%Y à %H:%M"),
        "df": {k: S[k].to_json(orient="table", date_format="iso", double_precision=15) for k in CLES_DF if S.get(k) is not None},
        "obj": {k: S.get(k) for k in CLES_OBJ},
        "location": {"df_base": rp.df_base.to_json(orient="table", date_format="iso", double_precision=15),
                     "prochain_camion": int(rp._next_camion_num)},
    }
    return json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")


def importer_session(fichier):
    data = json.load(fichier)
    if data.get("format") != "planning-livraisons":
        raise ValueError("ce fichier n'est pas une sauvegarde de Planning Livraisons")
    lire = lambda t: pd.read_json(io.StringIO(t), orient="table")  # noqa: E731
    for k in CLES_DF:
        S[k] = lire(data["df"][k]) if k in data["df"] else None
    for k in CLES_OBJ:
        S[k] = data["obj"].get(k) or copy.deepcopy(ETAT_INITIAL[k])
    rp = TruckRentalProcessor(S.df_optimized_estafettes, S.df_livraisons_original)
    rp.df_base = lire(data["location"]["df_base"])
    rp._next_camion_num = data["location"]["prochain_camion"]
    S.rental_processor = rp
    S.df_livraisons = S.df_grouped_zone
    S.data_processed = True
    S.export_planning = None
    S.ver_editeur += 1
    actualiser_propositions()
    return data.get("date", "")


# =====================================================
# DIALOGUES : CONTRÔLE QUALITÉ
# =====================================================
@st.dialog("Articles avec poids suspect (0 ou 1000)", width="large")
def dialogue_poids():
    df = S.df_poids_suspects
    st.markdown(f"**{len(df)} article(s)** répartis sur **{df['No livraison'].nunique()} BL** ont un poids "
                "égal à 0 ou 1000. Le brouillon d'e-mail s'ouvre dans Outlook avec la liste jointe.")
    st.dataframe(df, hide_index=True, width="stretch",
                 column_config={"Poids": st.column_config.NumberColumn("Poids", format="%.3f")})
    c1, c2 = st.columns(2)
    c1.download_button("E-mail Outlook (.eml)", data=xp.eml_poids_suspects(df), file_name="Email_poids_a_corriger.eml",
                       mime="message/rfc822", type="primary", width="stretch", icon=":material/mail:", on_click="ignore")
    c2.download_button("Liste Excel", data=xp.excel_poids_suspects(df), file_name="articles_poids_a_corriger.xlsx",
                       mime=xp.MIME_XLSX, width="stretch", icon=":material/download:", on_click="ignore")


@st.dialog("Articles avec volume nul ou non renseigné", width="large")
def dialogue_volumes():
    df = S.df_volumes_nuls
    st.markdown(f"**{len(df)} article(s)** ont un volume nul ou absent du fichier Volumes. "
                "Le brouillon d'e-mail s'ouvre dans Outlook avec la liste jointe.")
    st.dataframe(df, hide_index=True, width="stretch",
                 column_config={"Volume": st.column_config.NumberColumn("Volume", format="%.3f")})
    c1, c2 = st.columns(2)
    c1.download_button("E-mail Outlook (.eml)", data=xp.eml_volumes_nuls(df), file_name="Email_volume_a_corriger.eml",
                       mime="message/rfc822", type="primary", width="stretch", icon=":material/mail:", on_click="ignore")
    c2.download_button("Liste Excel", data=xp.excel_volumes_nuls(df), file_name="articles_volume_nul.xlsx",
                       mime=xp.MIME_XLSX, width="stretch", icon=":material/download:", on_click="ignore")


def nb_anomalies():
    n_p = 0 if S.df_poids_suspects is None else len(S.df_poids_suspects)
    n_v = 0 if S.df_volumes_nuls is None else len(S.df_volumes_nuls)
    return n_p, n_v


# =====================================================
# PAGE 1 : IMPORTATION
# =====================================================
FICHIERS = [
    ("Livraisons (BL)", "Extraction des bons de livraison : N° livraison, article, client, site, quantité, poids."),
    ("Articles / Volumes", "Fiche articles : volume de l'US, unité de volume, désignation."),
    ("Clients / Zones", "Référentiel clients : client, ville, représentant."),
]


def traiter(liv, art, cli):
    proc = DeliveryProcessor()
    (S.df_grouped, S.df_city, S.df_grouped_zone, S.df_zone,
     S.df_optimized_estafettes, S.df_livraisons_original) = proc.process_delivery_data(liv, art, cli)
    S.df_livraisons = S.df_grouped_zone
    S.rental_processor = TruckRentalProcessor(S.df_optimized_estafettes, S.df_livraisons_original)
    S.decisions = []
    S.data_processed = True
    reconstruire_plan()

    # Contrôles qualité (non bloquants)
    liv.seek(0)
    art.seek(0)
    df_liv_raw, df_art_raw = pd.read_excel(liv), pd.read_excel(art)
    for cle, fonction in (("df_poids_suspects", xp.detecter_poids_suspects),
                          ("df_volumes_nuls", xp.detecter_volumes_nuls)):
        try:
            S[cle] = fonction(df_liv_raw, df_art_raw)
        except Exception as e:  # le contrôle ne doit pas bloquer le planning
            S[cle] = None
            S["_alerte_qualite"] = f"Contrôle qualité non effectué : {e}"
    S.fichiers = [liv.name, art.name, cli.name]


def page_import():
    page_header("Importation des données",
                "Chargez les trois extractions Excel du jour pour calculer les voyages.", etape=1)
    afficher_flash()

    zone_fichiers = (st.expander("Fichiers importés : " + " · ".join(f.replace("_", "\\_") for f in S.fichiers)
                                 + " — cliquez pour en changer",
                                 icon=":material/folder_open:")
                     if S.data_processed else st.container())
    with zone_fichiers:
        fichiers = []
        for col, (titre, desc), i in zip(st.columns(3, gap="medium"), FICHIERS, range(3)):
            with col, st.container(border=True):
                st.markdown(f'<div class="up-title">{i + 1}. {titre}</div><div class="up-desc">{desc}</div>',
                            unsafe_allow_html=True)
                f = st.file_uploader(titre, type=["xlsx"], key=f"file{i + 1}_{S.ver_upload}",
                                     label_visibility="collapsed")
                fichiers.append(f)

        manquants = sum(f is None for f in fichiers)
        _, centre, _ = st.columns([1, 1.4, 1])
        with centre:
            libelle = "Relancer le traitement" if S.data_processed else "Lancer le traitement"
            lancer = st.button(libelle, type="secondary" if S.data_processed else "primary", icon=":material/play_arrow:", width="stretch",
                               disabled=manquants > 0)
            if manquants:
                st.caption(f"{manquants} fichier(s) manquant(s) sur 3.")
            elif S.data_processed:
                st.caption("Relancer recalcule tout le planning (les décisions en cours seront perdues).")
    if lancer:
        try:
            with st.spinner("Traitement des fichiers en cours…"):
                traiter(*fichiers)
        except Exception as e:
            S.data_processed = False
            st.error(f"Le traitement a échoué : {nettoyer_message(e)}", icon=":material/error:")
        else:
            alerte = S.pop("_alerte_qualite", None)
            flash(f"Données importées et traitées. {alerte}" if alerte else "Données importées et traitées.",
                  "warning" if alerte else "success")
            st.rerun()

    if not S.data_processed:
        with st.expander("Reprendre une session sauvegardée", icon=":material/history:"):
            st.caption("Rechargez le fichier .json téléchargé depuis « Sauvegarder la session » pour retrouver "
                       "décisions, transferts, validations et attributions.")
            sauvegarde = st.file_uploader("Sauvegarde", type=["json"], key=f"reprise_{S.ver_upload}",
                                          label_visibility="collapsed")
            if st.button("Reprendre la session", icon=":material/restore:", disabled=sauvegarde is None):
                try:
                    date = importer_session(sauvegarde)
                except Exception as e:
                    st.error(f"Impossible de reprendre la session : {nettoyer_message(e)}", icon=":material/error:")
                else:
                    flash(f"Session du {date} reprise.")
                    aller("optimisation")
                    st.rerun()
        return

    dfz = S.df_grouped_zone
    n_prop = 0 if S.propositions is None else len(S.propositions)
    section("Résumé du traitement")
    kpis([
        ("Bons de livraison", fmt_int(dfz["No livraison"].nunique())),
        ("Clients", fmt_int(dfz["Client de l'estafette"].nunique())),
        ("Villes", fmt_int(S.df_city["Ville"].nunique())),
        ("Zones", fmt_int(S.df_zone["Zone"].nunique())),
        ("Voyages proposés", fmt_int(len(S.df_voyages)), None, "primary"),
        ("Locations à arbitrer", fmt_int(n_prop), "client(s) au-dessus des seuils", "warn" if n_prop else "ok"),
    ])

    n_p, n_v = nb_anomalies()
    section("Contrôle qualité des fiches articles", "À faire corriger avant de relancer le planning.")
    c1, c2 = st.columns(2, gap="medium")
    for col, n, titre, dialogue, k in ((c1, n_p, "Poids égal à 0 ou 1000", dialogue_poids, "q_poids"),
                                       (c2, n_v, "Volume nul ou non renseigné", dialogue_volumes, "q_vol")):
        with col, st.container(border=True):
            a, b = st.columns([3, 2], vertical_alignment="center")
            a.markdown(f"**{titre}**  \n" + (f":orange[{n} article(s) à corriger]" if n else ":green[Aucun article concerné]"))
            if b.button("Voir et envoyer", key=k, disabled=not n, width="stretch", icon=":material/mail:"):
                dialogue()

    st.write("")
    _, droite = st.columns([3, 1])
    droite.button("Continuer vers l'analyse", type="primary", icon=":material/arrow_forward:",
                  width="stretch", on_click=aller, args=("analyse",))


# =====================================================
# PAGE 2 : ANALYSE
# =====================================================
def barre_outils(df, cle, data_excel, nom_fichier):
    a, b = st.columns([3, 1], vertical_alignment="bottom")
    with a:
        filtre = filtre_texte(df, f"q_{cle}")
    with b:
        bouton_telechargement("Exporter Excel", data_excel, nom_fichier, key=f"dl_{cle}")
    return filtre


def vue_livraisons(df, avec_zone):
    cle = "liv_zone" if avec_zone else "liv_ville"
    if not avec_zone:
        df = df.drop(columns=["Zone"], errors="ignore")
        df = df[~df["Ville"].isin(m.VILLES_EXCLUES_ANALYSE)]
    if df.empty:
        etat_vide("Aucune livraison à afficher")
        return
    kpis([
        ("Bons de livraison", fmt_int(df["No livraison"].nunique())),
        ("Clients", fmt_int(df["Client de l'estafette" if avec_zone else "Client"].nunique())),
        ("Poids total", fmt_kg(df["Poids total"].sum())),
        ("Volume total", fmt_m3(df["Volume total"].sum())),
    ])
    data = xp.excel_livraisons(S.df_grouped_zone if avec_zone else S.df_grouped, sans_zone=not avec_zone,
                               nom="Livraisons Client Ville Zone" if avec_zone else "Livraisons Client Ville")
    vue = barre_outils(df, cle, data, "Livraisons_Client_Ville_Zone.xlsx" if avec_zone else "Livraisons_Client_Ville.xlsx")
    detail = st.toggle("Afficher le détail des articles", key=f"detail_{cle}")
    affichage = vue.copy()
    if not detail:
        affichage["Article"] = affichage["Article"].map(lambda x: f"{len(m.liste(x, r','))} article(s)")
        affichage = affichage.drop(columns=[c for c in affichage.columns if "signation" in str(c)])
        affichage = affichage.rename(columns={"Article": "Articles"})
    if avec_zone and "Zone" in affichage.columns:
        cols = [c for c in affichage.columns if c != "Zone"]
        cols.insert(cols.index("Ville") + 1, "Zone")
        affichage = affichage[cols]
    affichage["Poids total"] = affichage["Poids total"].map(fmt_kg)
    affichage["Volume total"] = affichage["Volume total"].map(fmt_m3)
    st.caption(f"{len(vue)} ligne(s) affichée(s)" + ("" if avec_zone else " · TRIPOLI exclue de l'analyse"))
    tableau(affichage, multiligne={"Article": SEP_VIRGULE, "Désignation": SEP_LISTE},
            numeriques=("Poids total", "Volume total"), nowrap=("No livraison",), hauteur_max=620)


def vue_besoin(df, par_zone):
    d = df.rename(columns={"Nombre livraisons": "Nombre de BLs"}).copy()
    if not par_zone:
        d = d[~d["Ville"].isin(m.VILLES_EXCLUES_ANALYSE)]
    cle_lieu = "Zone" if par_zone else "Ville"
    kpis([
        ("Zones" if par_zone else "Villes", fmt_int(len(d))),
        ("Bons de livraison", fmt_int(d["Nombre de BLs"].sum())),
        ("Poids total", fmt_kg(d["Poids total"].sum())),
        ("Besoin en estafettes", fmt_num(m.total_besoin(d), 1), f"Zone 7 : {m.MAX_VOYAGES_ZONE7} voyages max / estafette", "primary"),
    ])
    data = xp.excel_besoin(df, par_zone=par_zone, nom="Besoin Estafette Zone" if par_zone else "Besoin Estafette Ville")
    vue = barre_outils(d, f"besoin_{cle_lieu}", data,
                       "Besoin_Estafette_Zone.xlsx" if par_zone else "Besoin_Estafette_Ville.xlsx")
    if par_zone:
        vue = m.trier_par_zone(vue)
    affichage = pd.DataFrame({
        cle_lieu: vue[cle_lieu],
        "BL": vue["Nombre de BLs"].map(fmt_int),
        "Poids total": vue["Poids total"].map(fmt_kg),
        "Volume total": vue["Volume total"].map(fmt_m3),
        "Besoin (poids)": vue["Besoin estafette (poids)"].map(fmt_int),
        "Besoin (volume)": vue["Besoin estafette (volume)"].map(fmt_int),
        "Besoin retenu": vue.apply(lambda r: m.libelle_besoin(r["Besoin estafette réel"], m.regle_zone7(r)), axis=1),
    })
    tableau(affichage, numeriques=("BL", "Poids total", "Volume total", "Besoin (poids)", "Besoin (volume)"))


def vue_graphiques():
    df = S.df_city[~S.df_city["Ville"].isin(m.VILLES_EXCLUES_ANALYSE)].copy()
    if df.empty:
        etat_vide("Aucune donnée pour les graphiques")
        return

    df["Besoin estafettes"] = df.apply(lambda r: m.estafettes_necessaires(r["Besoin estafette réel"], m.regle_zone7(r)), axis=1)
    reglages = [("Poids total", "Poids livré par ville (kg)"), ("Volume total", "Volume livré par ville (m³)"),
                ("Nombre de BLs", "Nombre de BL par ville"), ("Besoin estafettes", "Besoin en estafettes par ville")]
    for ligne in (reglages[:2], reglages[2:]):
        for col, (y, titre) in zip(st.columns(2, gap="medium"), ligne):
            with col, st.container(border=True):
                d = df.sort_values(y, ascending=False)
                fig = px.bar(d, x="Ville", y=y, color_discrete_sequence=[BLEU],
                             text=y if y == "Besoin estafettes" else None)
                if y == "Besoin estafettes":
                    fig.update_traces(textposition="outside", cliponaxis=False)
                graphique(style_fig(fig, titre))
    st.caption(f"Zone 7 / SFAX : {m.MAX_VOYAGES_ZONE7} voyages maximum par estafette (valeur provisoire).")


def page_analyse():
    page_header("Analyse des livraisons",
                "Vue consolidée des BL et du besoin en estafettes, par ville et par zone.", etape=2)
    if not garde():
        return

    n_p, n_v = nb_anomalies()
    if n_p or n_v:
        with st.container(border=True):
            a, b, c = st.columns([4, 1.2, 1.2], vertical_alignment="center")
            a.markdown(f":orange[:material/warning:] **Fiches articles à corriger** : {n_p} poids suspect(s), "
                       f"{n_v} volume(s) nul(s). Les poids et volumes concernés faussent le calcul des voyages.")
            if b.button("Poids", disabled=not n_p, width="stretch", key="ana_poids"):
                dialogue_poids()
            if c.button("Volumes", disabled=not n_v, width="stretch", key="ana_vol"):
                dialogue_volumes()

    onglets = st.tabs([":material/receipt_long: Livraisons par ville", ":material/location_city: Besoin par ville",
                       ":material/map: Livraisons par zone", ":material/hub: Besoin par zone",
                       ":material/bar_chart: Graphiques"])
    with onglets[0]:
        vue_livraisons(S.df_grouped, avec_zone=False)
    with onglets[1]:
        vue_besoin(S.df_city, par_zone=False)
    with onglets[2]:
        vue_livraisons(S.df_grouped_zone, avec_zone=True)
    with onglets[3]:
        vue_besoin(S.df_zone, par_zone=True)
    with onglets[4]:
        vue_graphiques()

    st.divider()
    a, _, c = st.columns([1.2, 2, 1.2])
    with a:
        bouton_telechargement("Analyse complète (Excel)",
                              xp.excel_analyse_complete(S.df_grouped, S.df_city, S.df_grouped_zone, S.df_zone),
                              "Analyse_Complete_Livraisons.xlsx", key="dl_analyse")
    c.button("Passer à l'optimisation", type="primary", icon=":material/arrow_forward:", width="stretch",
             on_click=aller, args=("optimisation",))


# =====================================================
# PAGE 3 : OPTIMISATION — sous-étapes
# =====================================================
def decider(client, accepter, camion):
    ok, msg, _ = S.rental_processor.appliquer_location(client, accepter=accepter, truck_type=camion)
    if ok:
        S.decisions.append({
            "Heure": datetime.now(TZ).strftime("%H:%M"), "Client": client,
            "Décision": f"Location acceptée ({camion})" if accepter else "Location refusée",
        })
        reconstruire_plan()
        S.pop("loc_client", None)
        S.pop("truck_type", None)
        flash(msg, "success" if accepter else "info")
    else:
        flash(msg, "error")


def _valeur(texte):
    """'1874.050 kg' -> 1874.05 (détails renvoyés formatés par le backend)."""
    try:
        return float(re.sub(r"[^\d.\-]", "", str(texte)))
    except ValueError:
        return None


def camion_suggere(poids, volume):
    """Plus petit camion qui convient à la commande."""
    for t in m.TYPES_CAMION:
        if poids <= get_capacite_poids_camion(t) and volume <= get_capacite_volume_camion(t):
            return t
    return m.TYPES_CAMION[-1]


def _suggerer_camion(props):
    ligne = props[props["Client"].astype(str) == S.get("loc_client")]
    if not ligne.empty:
        S.truck_type = camion_suggere(float(ligne["Poids total (kg)"].iloc[0]), float(ligne["Volume total (m³)"].iloc[0]))


def motif_location(r):
    raisons = []
    if r["Poids total (kg)"] >= SEUIL_POIDS:
        raisons.append(f"Poids ≥ {fmt_num(SEUIL_POIDS, 0, 'kg')}")
    if r["Volume total (m³)"] >= SEUIL_VOLUME:
        raisons.append(f"Volume ≥ {fmt_m3(SEUIL_VOLUME)}")
    return " et ".join(raisons)


def vue_location():
    st.caption(f"Une location est proposée lorsqu'un client dépasse {fmt_kg(SEUIL_POIDS)} "
               f"ou {fmt_m3(SEUIL_VOLUME)} au total.")
    props = S.propositions
    if props is None or props.empty:
        st.success("Aucune proposition de location en attente de décision.", icon=":material/check_circle:")
        st.button("Revérifier les propositions", icon=":material/refresh:", on_click=actualiser_propositions)
    else:
        gauche, droite = st.columns([2, 3], gap="large")
        with gauche:
            section("Propositions en attente", f"{len(props)} client(s) à arbitrer")
            tableau(pd.DataFrame({
                "Client": props["Client"],
                "Poids": props["Poids total (kg)"].map(fmt_kg),
                "Volume": props["Volume total (m³)"].map(fmt_m3),
                "Motif": props.apply(motif_location, axis=1),
            }), numeriques=("Poids", "Volume"), hauteur_max=260)

            client = st.selectbox("Client à traiter", props["Client"].astype(str).tolist(), key="loc_client",
                                  on_change=_suggerer_camion, args=(props,))
            if "truck_type" not in S:
                _suggerer_camion(props)
            camion = st.segmented_control("Camion à louer", m.TYPES_CAMION, key="truck_type",
                                          help="Le plus petit camion adapté est présélectionné.") or "5 tonnes"
            ligne = props[props["Client"].astype(str) == client].iloc[0]
            poids, volume = float(ligne["Poids total (kg)"]), float(ligne["Volume total (m³)"])
            cap_p, cap_v = get_capacite_poids_camion(camion), get_capacite_volume_camion(camion)
            compatible = poids <= cap_p and volume <= cap_v
            jauges(f"Remplissage du camion {camion}",
                   [("Poids", poids / cap_p * 100), ("Volume", volume / cap_v * 100)])
            if not compatible:
                st.error(f"La commande dépasse la capacité du camion {camion} "
                         f"({fmt_num(cap_p, 0, 'kg')} · {fmt_num(cap_v, 0, 'm³')}). Choisissez un camion plus grand.",
                         icon=":material/block:")
            if plan_modifie():
                st.warning("Décider maintenant recalcule les voyages : transferts, objets ajoutés et validations "
                           "en cours seront réinitialisés.", icon=":material/warning:")
            b1, b2 = st.columns(2)
            b1.button("Accepter la location", type="primary", icon=":material/check:", width="stretch",
                      disabled=not compatible, on_click=decider, args=(client, True, camion))
            b2.button("Refuser", icon=":material/close:", width="stretch",
                      on_click=decider, args=(client, False, camion))

        with droite:
            resume, details = S.rental_processor.get_details_client(client)
            etat = resume.split("État :")[-1].strip() if "État :" in resume else "—"
            section(f"Client {client}", "Commande consolidée et estafettes qui la transportent aujourd'hui")
            kpis([("Poids total réel", fmt_kg(poids)), ("Volume total réel", fmt_m3(volume)),
                  ("Estafettes concernées", fmt_int(len(details))), ("État", etat)])
            if not details.empty:
                tableau(pd.DataFrame({
                    "Zone": details["Zone"], "Véhicule": details["Camion N°"],
                    "Poids": details["Poids total"].map(lambda x: fmt_kg(_valeur(x))),
                    "Volume": details["Volume total"].map(lambda x: fmt_m3(_valeur(x))),
                    "BL": details["BL inclus"],
                    "Occupation": details["Taux d'occupation (%)"].map(lambda x: occupation(_valeur(x))),
                }), multiligne={"BL": SEP_LISTE}, numeriques=("Poids", "Volume"), html_cols=("Occupation",),
                    nowrap=("Véhicule", "BL"), hauteur_max=420)

    if S.decisions:
        with st.expander(f"Décisions prises ({len(S.decisions)})", icon=":material/history:"):
            st.dataframe(pd.DataFrame(S.decisions), hide_index=True, width="stretch")


def _choix_defaut(r):
    """Véhicule / chauffeur déjà choisis, sinon valeurs par défaut des camions loués."""
    a = S.attributions.get(r["Véhicule N°"], {})
    if a:
        return a.get("Véhicule"), a.get("Chauffeur_complet")
    if m.est_camion(r):
        return m.VEHICULE_CAMION, m.CHAUFFEUR_CAMION
    return None, None


def _maj_plan(cle, vehicules, vue):
    """Enregistre les cases 'Valider' et les attributions saisies dans le tableau du plan."""
    for i, ch in S[cle]["edited_rows"].items():
        i = int(i)
        v = vehicules[i]
        if "Valider" in ch:
            if ch["Valider"]:
                S.validations[v] = "Oui"
            else:
                S.validations.pop(v, None)
        if "Véhicule attribué" in ch or "Chauffeur" in ch:
            actuel = S.attributions.get(v) or {"Véhicule": vue["Véhicule attribué"].iloc[i],
                                                "Chauffeur_complet": vue["Chauffeur"].iloc[i]}
            if "Véhicule attribué" in ch:
                actuel["Véhicule"] = ch["Véhicule attribué"]
            if "Chauffeur" in ch:
                actuel["Chauffeur_complet"] = ch["Chauffeur"]
            S.attributions[v] = actuel
    invalider_validation()  # toute modification impose de regénérer les codes


def cocher_tout(vehicules, valeur):
    for v in vehicules:
        if valeur:
            S.validations[v] = "Oui"
        else:
            S.validations.pop(v, None)
    invalider_validation()
    S.ver_editeur += 1


def preremplir_attributions(df):
    """Propose un véhicule et un chauffeur distincts pour chaque voyage validé encore vide."""
    valides = df[df["Véhicule N°"].map(lambda v: S.validations.get(v) == "Oui")]
    pris_v = {a.get("Véhicule") for a in S.attributions.values()}
    pris_c = {a.get("Chauffeur_complet") for a in S.attributions.values()}
    libres_v = [v for v in m.VEHICULES_DISPONIBLES if v != m.VEHICULE_CAMION and v not in pris_v]
    libres_c = [c for c in m.OPTIONS_CHAUFFEURS if c != m.CHAUFFEUR_CAMION and c not in pris_c]
    for _, r in valides.iterrows():
        veh, ch = _choix_defaut(r)
        if veh and ch:
            continue
        if not veh and libres_v:
            veh = libres_v.pop(0)
        if not ch and libres_c:
            ch = libres_c.pop(0)
        S.attributions[r["Véhicule N°"]] = {"Véhicule": veh, "Chauffeur_complet": ch}
    invalider_validation()
    S.ver_editeur += 1


def etat_plan(df):
    """Compteurs du plan : validés, attribués parmi les validés."""
    valides = df[df["Véhicule N°"].map(lambda v: S.validations.get(v) == "Oui")]
    choix = valides.apply(_choix_defaut, axis=1) if not valides.empty else []
    n_attr = sum(1 for veh, ch in choix if veh and ch)
    return valides, n_attr


def generer_codes(date_voyage, debut):
    df = m.trier_par_zone(S.df_voyages)
    valides = df[df["Véhicule N°"].map(lambda v: S.validations.get(v) == "Oui")].copy()
    choix = valides.apply(_choix_defaut, axis=1)
    valides["Véhicule attribué"] = [c[0] for c in choix]
    valides["Chauffeur attribué"] = [str(c[1]).split(" - ", 1)[1] for c in choix]
    valides["Matricule chauffeur"] = [str(c[1]).split(" - ", 1)[0] for c in choix]
    for v, c in zip(valides["Véhicule N°"], choix):
        S.attributions[v] = {"Véhicule": c[0], "Chauffeur_complet": c[1]}
    valides["Date Voyage Format"] = date_voyage.strftime("%Y%m%d")
    valides["Numero Séquentiel"] = range(int(debut), int(debut) + len(valides))
    valides["Code voyage"] = (valides["Véhicule N°"].astype(str) + "/" + valides["Date Voyage Format"] + "/" +
                              valides["Numero Séquentiel"].map(lambda x: f"{x:03d}"))
    S.df_voyages_valides = valides.reset_index(drop=True)
    S.export_planning = None
    flash(f"{len(valides)} voyage(s) validé(s) et codifié(s). Le planning final est prêt.")


# ---------- Utilitaires : transfert de BL et ajout d'objet (fenêtres) ----------
@st.dialog("Transférer des BL entre véhicules", width="large")
def dialogue_transfert():
    df = S.df_voyages
    a, b, c = st.columns(3)
    zone = a.selectbox("Zone", zones_triees(df), key="tr_zone")
    dz = df[df["Zone"] == zone]
    vehicules = sorted(dz["Véhicule N°"].dropna().unique().tolist(), key=cle_vehicule)
    infos = {r["Véhicule N°"]: r for _, r in dz.iterrows()}

    def etiquette(v):
        r = infos[v]
        return f"{v} · {fmt_kg(r['Poids total chargé'])} · {fmt_pct(r[TAUX])}"

    if len(vehicules) < 2:
        st.info("Cette zone ne compte qu'un véhicule : aucun transfert possible.", icon=":material/info:")
        return
    # Source : les plus chargés d'abord ; cible : les moins chargés d'abord
    par_taux = sorted(vehicules, key=lambda v: infos[v][TAUX])
    source = b.selectbox("Depuis", par_taux[::-1], format_func=etiquette, key="tr_src")
    cible = c.selectbox("Vers", [v for v in par_taux if v != source], format_func=etiquette, key="tr_dst")

    liv = S.df_livraisons.assign(_bl=S.df_livraisons["No livraison"].astype(str).str.strip())
    lignes = []
    for bl in [x for x in m.bls_de(infos[source]["BL inclus"]) if not x.startswith("OBJ-")]:
        info = liv[liv["_bl"] == bl]
        lignes.append({
            "Transférer": False, "BL": bl,
            "Client": info["Client de l'estafette"].iloc[0] if not info.empty else "Client non trouvé",
            "Poids (kg)": float(pd.to_numeric(info["Poids total"], errors="coerce").sum()) if not info.empty else 0.0,
            "Volume (m³)": float(pd.to_numeric(info["Volume total"], errors="coerce").sum()) if not info.empty else 0.0,
        })
    if not lignes:
        st.info(f"Aucun BL transférable dans {source}.", icon=":material/info:")
        return
    edite = st.data_editor(
        pd.DataFrame(lignes), hide_index=True, width="stretch", key=f"tr_ed_{zone}_{source}_{S.ver_editeur}",
        disabled=["BL", "Client", "Poids (kg)", "Volume (m³)"],
        column_config={"Transférer": st.column_config.CheckboxColumn(width="small"),
                       "Poids (kg)": st.column_config.NumberColumn(format="%.1f"),
                       "Volume (m³)": st.column_config.NumberColumn(format="%.3f")})
    sel = edite[edite["Transférer"]]
    p_sel, v_sel = sel["Poids (kg)"].sum(), sel["Volume (m³)"].sum()
    rc = infos[cible]
    cap_p, cap_v = m.capacites(rc)
    tp = (rc["Poids total chargé"] + p_sel) / cap_p * 100
    tv = (rc["Volume total chargé"] + v_sel) / cap_v * 100
    jauges(f"{cible} après transfert de {len(sel)} BL ({fmt_kg(p_sel)} · {fmt_m3(v_sel)})",
           [("Poids", tp), ("Volume", tv)])
    depasse = tp > 100 or tv > 100
    if depasse and len(sel):
        st.error(f"Le transfert dépasserait la capacité de {cible}.", icon=":material/block:")
    if st.button(f"Transférer {len(sel)} BL vers {cible}", type="primary", icon=":material/swap_horiz:",
                 disabled=sel.empty or depasse, width="stretch"):
        tm = TruckTransferManager(S.df_voyages, S.df_livraisons)
        ok, msg, df_maj = tm.transferer_bls(zone, source, cible, sel["BL"].tolist())
        if ok:
            S.df_voyages = m.ajouter_taux(df_maj)
            S.transferts_historique.append({
                "Heure": datetime.now(TZ).strftime("%H:%M"), "Zone": zone, "De": source, "Vers": cible,
                "BL": ", ".join(sel["BL"]), "Poids (kg)": round(p_sel, 1), "Volume (m³)": round(v_sel, 3)})
            S.ver_editeur += 1
            invalider_validation()
            flash(msg)
            st.rerun()
        st.error(nettoyer_message(msg), icon=":material/error:")


@st.dialog("Ajouter un objet hors BL", width="medium")
def dialogue_objet():
    df = S.df_voyages
    a, b = st.columns(2)
    zone = a.selectbox("Zone", zones_triees(df), key="obj_zone")
    dz = df[df["Zone"] == zone]
    vehicule = b.selectbox("Véhicule", sorted(dz["Véhicule N°"].dropna().unique(), key=cle_vehicule), key="obj_veh")
    r = dz[dz["Véhicule N°"] == vehicule].iloc[0]
    cap_p, cap_v = m.capacites(r)
    jauges(f"{vehicule} ({m.libelle_type(r)}) — {fmt_kg(r['Poids total chargé'])} / {fmt_num(cap_p, 0, 'kg')}",
           [("Poids", r["Taux occupation poids (%)"]), ("Volume", r["Taux occupation volume (%)"])])
    nom = st.text_input("Désignation de l'objet", placeholder="Ex. matériel urgent, colis oublié…", key="obj_nom")
    c1, c2 = st.columns(2)
    poids = c1.number_input("Poids (kg)", min_value=0.0, max_value=10000.0, value=10.0, step=0.1, key="obj_poids")
    volume = c2.number_input("Volume (m³)", min_value=0.0, max_value=100.0, value=0.1, step=0.01, format="%.3f",
                             key="obj_vol")
    if st.button("Ajouter au véhicule", type="primary", icon=":material/add:", width="stretch"):
        if not nom.strip():
            st.error("Donnez une désignation à l'objet.", icon=":material/error:")
            return
        if poids <= 0 or volume <= 0:
            st.error("Le poids et le volume doivent être supérieurs à 0.", icon=":material/error:")
            return
        tm = TruckTransferManager(S.df_voyages, S.df_livraisons)
        ok, msg, df_maj = tm.add_manual_object(df_voyages=S.df_voyages, vehicle=vehicule, zone=zone,
                                               name=nom.strip(), weight=poids, volume=volume)
        if ok:
            S.df_voyages = m.ajouter_taux(df_maj)
            S.objets_historique.append({
                "Date": datetime.now(TZ).strftime("%d/%m/%Y %H:%M"), "Zone": zone, "Véhicule": vehicule,
                "Objet": nom.strip(), "Poids (kg)": poids, "Volume (m³)": volume})
            S.ver_editeur += 1
            invalider_validation()
            flash(msg)
            st.rerun()
        st.error(nettoyer_message(msg), icon=":material/error:")


# ---------- Onglet principal : le plan de chargement ----------
def vue_plan():
    df_all = m.trier_par_zone(S.df_voyages).reset_index(drop=True)
    valides, n_attr = etat_plan(df_all)
    n_sur = int((df_all[TAUX] > 100).sum())
    kpis([
        ("Voyages", fmt_int(len(df_all)), f"{m.nb_estafettes(df_all)} estafette(s) · {m.nb_camions(df_all)} camion(s)"),
        ("Poids total", fmt_kg(df_all["Poids total chargé"].sum())),
        ("En surcharge", fmt_int(n_sur), "taux > 100 %", "danger" if n_sur else "ok"),
        ("Validés", f"{len(valides)} / {len(df_all)}", None, "ok" if len(valides) == len(df_all) else "warn"),
        ("Véhicule et chauffeur", f"{n_attr} / {len(valides)}", "parmi les voyages validés",
         "ok" if len(valides) and n_attr == len(valides) else "warn"),
    ])
    if n_sur:
        noms = ", ".join(f"{r['Véhicule N°']} ({fmt_pct(r[TAUX])})" for _, r in df_all[df_all[TAUX] > 100].iterrows())
        st.warning(f"Capacité dépassée : {noms}. Utilisez « Transférer des BL » pour rééquilibrer.",
                   icon=":material/warning:")

    # Barre d'outils : filtre + utilitaires + export
    a, b, c, d = st.columns([1.6, 1.2, 1.2, 1.2], vertical_alignment="bottom")
    zone = a.selectbox("Zone", ["Toutes les zones"] + zones_triees(df_all), key="plan_zone",
                       label_visibility="collapsed")
    if b.button("Transférer des BL", icon=":material/swap_horiz:", width="stretch"):
        dialogue_transfert()
    if c.button("Ajouter un objet", icon=":material/add_box:", width="stretch"):
        dialogue_objet()
    with d:
        bouton_telechargement("Exporter Excel", xp.excel_voyages(S.df_voyages, "Voyages Optimisés"),
                              "Voyages_Estafette_Optimises.xlsx", key="dl_voyages")

    df = df_all if zone == "Toutes les zones" else df_all[df_all["Zone"] == zone].reset_index(drop=True)
    vehicules = df["Véhicule N°"].tolist()
    defauts = df.apply(_choix_defaut, axis=1) if not df.empty else []
    # Colonnes à saisir en premier (visibles sans défilement), informations ensuite
    vue = pd.DataFrame({
        "Valider": [S.validations.get(v) == "Oui" for v in vehicules],
        "Voyage": df["Véhicule N°"],
        "Zone": df["Zone"],
        "Véhicule attribué": [x[0] for x in defauts], "Chauffeur": [x[1] for x in defauts],
        "Occupation (%)": df[TAUX],
        "Poids (kg)": df["Poids total chargé"], "Volume (m³)": df["Volume total chargé"],
        "BL": df["BL inclus"].map(lambda x: len(m.bls_de(x))),
        "Clients": df["Client(s) inclus"].map(lambda x: ", ".join(m.liste(x))),
        "Type": df.apply(m.libelle_type, axis=1),
    })

    e1, e2, e3, _ = st.columns([1, 1, 1.5, 1.9])
    e1.button("Tout valider", icon=":material/done_all:", width="stretch", on_click=cocher_tout,
              args=(vehicules, True))
    e2.button("Tout décocher", icon=":material/remove_done:", width="stretch", on_click=cocher_tout,
              args=(vehicules, False))
    e3.button("Pré-remplir véhicules et chauffeurs", icon=":material/auto_fix_high:", width="stretch",
              on_click=preremplir_attributions, args=(df_all,), disabled=not len(valides),
              help="Propose un véhicule et un chauffeur différents pour chaque voyage validé encore vide.")
    cle = f"ed_plan_{zone}_{S.ver_editeur}"
    st.data_editor(
        vue, key=cle, hide_index=True, width="stretch", height=min(48 + 36 * len(vue), 660),
        on_change=_maj_plan, args=(cle, vehicules, vue),
        disabled=[col for col in vue.columns if col not in ("Valider", "Véhicule attribué", "Chauffeur")],
        column_config={
            "Valider": st.column_config.CheckboxColumn(width="small"),
            "Voyage": st.column_config.TextColumn(width="small"),
            "Clients": st.column_config.TextColumn(width="large"),
            "BL": st.column_config.NumberColumn("Nb BL", width="small"),
            "Poids (kg)": st.column_config.NumberColumn(format="%.1f"),
            "Volume (m³)": st.column_config.NumberColumn(format="%.3f"),
            "Occupation (%)": st.column_config.ProgressColumn(format="%.1f %%", min_value=0, max_value=100),
            "Véhicule attribué": st.column_config.SelectboxColumn(options=m.VEHICULES_DISPONIBLES, width="medium"),
            "Chauffeur": st.column_config.SelectboxColumn(options=m.OPTIONS_CHAUFFEURS, width="medium"),
        })
    st.caption("Cochez les voyages à valider, puis choisissez le véhicule et le chauffeur. "
               "Les camions loués sont pré-remplis.")

    with st.expander("Détail des BL, clients et représentants par voyage", icon=":material/list:"):
        table_voyages(df, hauteur=480)
    if S.transferts_historique or S.objets_historique:
        with st.expander(f"Journal des modifications ({len(S.transferts_historique)} transfert(s), "
                         f"{len(S.objets_historique)} objet(s))", icon=":material/history:"):
            if S.transferts_historique:
                st.markdown("**Transferts de BL**")
                st.dataframe(pd.DataFrame(S.transferts_historique), hide_index=True, width="stretch")
            if S.objets_historique:
                st.markdown("**Objets ajoutés**")
                st.dataframe(pd.DataFrame(S.objets_historique), hide_index=True, width="stretch")

    section_codes(df_all)


def section_codes(df_all):
    valides, n_attr = etat_plan(df_all)
    manque = len(valides) - n_attr
    section("Finaliser : codes voyage", "Format véhicule / date / séquence — ex. E1/20261022/001")
    with st.container(border=True):
        a, b, c = st.columns(3, vertical_alignment="bottom")
        date_voyage = a.date_input("Date du voyage", value=datetime.now(TZ).date(), format="DD/MM/YYYY")
        debut = b.number_input("Premier numéro de séquence", min_value=1, max_value=1000, value=1)
        pret = planning_pret()
        c.button(f"{'Regénérer' if pret else 'Générer'} les codes ({len(valides)} voyage(s))", type="primary",
                 icon=":material/qr_code_2:", width="stretch", disabled=not len(valides) or manque > 0,
                 on_click=generer_codes, args=(date_voyage, debut))
        if not len(valides):
            st.caption("Validez au moins un voyage dans le tableau ci-dessus.")
        elif manque:
            st.caption(f"{manque} voyage(s) validé(s) sans véhicule ou chauffeur.")

    if not planning_pret():
        return
    dfv = S.df_voyages_valides
    tableau(pd.DataFrame({
        "Code voyage": dfv["Code voyage"], "Zone": dfv["Zone"], "Voyage": dfv["Véhicule N°"],
        "Véhicule": dfv["Véhicule attribué"], "Chauffeur": dfv["Chauffeur attribué"],
        "Poids": dfv["Poids total chargé"].map(fmt_kg),
    }), multiligne={"Zone": SEP_LISTE}, numeriques=("Poids",), nowrap=("Code voyage", "Voyage"), hauteur_max=420)
    cols = [col for col in xp.COLS_ATTRIBUES if col in dfv.columns]
    a, b, c, d = st.columns(4)
    with a:
        bouton_telechargement("Excel avec codes", xp.excel_voyages(dfv, "Voyages_avec_Codes", colonnes=xp.COLS_CODES),
                              "voyages_avec_codes.xlsx", key="dl_codes_xlsx")
    for col, label, fn, nom, k in ((b, "PDF avec codes", lambda: xp.pdf_codes(dfv), "voyages_avec_codes.pdf", "dl_codes_pdf"),
                                   (c, "PDF attribution", lambda: xp.pdf_attribues(dfv[cols]), "Voyages_attribues.pdf",
                                    "dl_attr_pdf")):
        with col:
            try:
                bouton_telechargement(label, fn(), nom, mime="application/pdf", key=k)
            except Exception as e:
                st.caption(f"{label} indisponible : {e}")
    d.button("Voir le planning final", type="primary", icon=":material/arrow_forward:", width="stretch",
             on_click=aller, args=("finalisation",))


VUES_OPT = {
    "Location": (":material/local_shipping:", vue_location),
    "Plan de chargement": (":material/route:", vue_plan),
}


def libelle_vue(nom):
    # Libellés fixes : un libellé qui change réinitialiserait la sélection du contrôle.
    return f"{VUES_OPT[nom][0]} {nom}"


def page_optimisation():
    page_header("Optimisation des chargements",
                "1. Arbitrez les locations de camions.  2. Sur le plan de chargement : ajustez, validez, "
                "attribuez véhicules et chauffeurs, puis générez les codes voyage.", etape=3)
    if not garde():
        return
    vue = st.segmented_control("Étape", list(VUES_OPT), key="opt_vue", format_func=libelle_vue,
                               label_visibility="collapsed") or "Location"
    afficher_flash()
    VUES_OPT[vue][1]()
    if vue == "Location":
        st.write("")
        _, droite = st.columns([3, 1])
        droite.button("Passer au plan de chargement", type="primary", icon=":material/arrow_forward:",
                      width="stretch", on_click=aller, args=("optimisation", "Plan de chargement"))


# =====================================================
# PAGE 4 : PLANNING & INDICATEURS
# =====================================================
def planning_pret():
    dfv = S.df_voyages_valides
    return dfv is not None and not dfv.empty and "Code voyage" in dfv.columns


def generer_export(nom, df_final):
    from backend import exporter_planning_excel
    supp = {k: v for k, v in (("Livraisons_Client_Ville", S.df_grouped), ("Besoin_Estafette_Ville", S.df_city),
                              ("Besoin_Estafette_Zone", S.df_zone)) if v is not None}
    df_exp = df_final.copy()
    for c in m.COLS_LISTES:
        df_exp[c] = df_exp[c].map(lambda x: "\n".join(m.liste(x)))
    with tempfile.TemporaryDirectory() as dossier:
        chemin = os.path.join(dossier, f"{nom}.xlsx")
        ok, msg = exporter_planning_excel(df_exp, chemin, supp, S.df_livraisons_original)
        if ok:
            with open(chemin, "rb") as f:
                S.export_planning = (f"{nom}.xlsx", f.read())
            flash("Planning généré : le fichier est prêt à être téléchargé.")
        else:
            flash(msg, "error")


def reinitialiser_optimisation():
    S.rental_processor = TruckRentalProcessor(S.df_optimized_estafettes, S.df_livraisons_original)
    S.decisions = []
    reconstruire_plan()
    aller("optimisation", "Location")
    flash("Optimisation réinitialisée : vous repartez des voyages calculés à l'import.", "info")


def section_rapports():
    from backend import AdvancedReportGenerator, calculer_couts_estimation, verifier_integrite_donnees
    section("Rapports et analyses", "Rapports texte générés par le moteur de calcul")
    gen = AdvancedReportGenerator(S.df_voyages, S.df_livraisons_original)
    t1, t2, t3, t4 = st.tabs([":material/analytics: Rapport analytique", ":material/person: Rapport client",
                              ":material/payments: Estimation des coûts", ":material/verified: Intégrité des données"])
    with t1:
        if st.button("Générer le rapport analytique", key="btn_rapport_analytique"):
            st.code(gen.generer_rapport_analytique(), language=None)
    with t2:
        exclus = {"AMECAP", "SANA", "SOPAL", "SOPALGAZ", "SOPALSERV", "SOPALTEC", "SOPALALG", "AQUAPLAST", "WINOX",
                  "QUIVEM", "SANISTONE", "SOPAMAR", "SOPALAFR", "SOPALINTER", "PERSOGSO", "SIAM"}
        d = S.df_livraisons_original
        d = d[~d["Client de l'estafette"].astype(str).str.strip().str.upper().isin(exclus)]
        if "Client commande" in d.columns:
            d = d[~d["Client commande"].astype(str).str.strip().str.upper().isin(exclus)]
        a, b = st.columns([3, 1], vertical_alignment="bottom")
        client = a.selectbox("Client", sorted(d["Client de l'estafette"].dropna().unique()), key="select_client_rapport")
        if b.button("Générer", key="btn_rapport_client", width="stretch"):
            st.code(gen.generer_rapport_client(client), language=None)
    with t3:
        a, b, c = st.columns([1, 1, 1], vertical_alignment="bottom")
        ce = a.number_input("Coût unitaire estafette (TND)", value=150, min_value=50, max_value=500)
        cc = b.number_input("Coût unitaire camion (TND)", value=800, min_value=300, max_value=2000)
        if c.button("Calculer", key="btn_calcul_couts", width="stretch"):
            couts = calculer_couts_estimation(S.df_voyages, ce, cc)
            if "erreur" in couts:
                st.error(nettoyer_message(couts["erreur"]))
            else:
                kpis([("Estafettes", f"{couts['estafettes']} × {ce} TND"),
                      ("Camions", f"{couts['camions']} × {cc} TND"),
                      ("Coût total estimé", fmt_num(couts["cout_total"], 0, "TND"), None, "primary")])
    with t4:
        if st.button("Vérifier l'intégrité des données", key="btn_verif_integrite"):
            res = verifier_integrite_donnees(S.df_voyages, S.df_livraisons_original)
            (st.success if "✅" in res else st.warning)(nettoyer_message(res))


def page_finalisation():
    page_header("Planning & indicateurs", "Synthèse du planning validé, export final et rapports.", etape=4)
    if not garde():
        return
    afficher_flash()

    pret = planning_pret()
    df_plan = S.df_voyages_valides if pret else S.df_voyages
    taux_moyen = df_plan["Taux d'occupation (%)"].mean() if not df_plan.empty else 0
    efficacite = "bonne" if taux_moyen > 70 else "moyenne" if taux_moyen > 50 else "faible"
    kpis([
        ("Voyages planifiés" if pret else "Voyages calculés", fmt_int(len(df_plan)),
         f"sur {len(S.df_voyages)} calculés" if pret else "planning non finalisé"),
        ("Estafettes", fmt_int(m.nb_estafettes(df_plan))),
        ("Camions loués", fmt_int(m.nb_camions(df_plan))),
        ("Poids total", fmt_num(df_plan["Poids total chargé"].sum(), 0, "kg")),
        ("Volume total", fmt_num(df_plan["Volume total chargé"].sum(), 1, "m³")),
        ("Taux d'occupation moyen", fmt_pct(taux_moyen), f"efficacité {efficacite}", "primary"),
    ])

    if not pret:
        dfv = S.df_voyages_valides
        valides, n_attr = etat_plan(S.df_voyages)
        etapes = [
            ("Voyages validés sur le plan de chargement", len(valides) > 0),
            ("Véhicule et chauffeur attribués aux voyages validés", len(valides) > 0 and n_attr == len(valides)),
            ("Codes voyage générés", False),
        ]
        with st.container(border=True):
            st.markdown("**Le planning final n'est pas encore prêt.** Étapes restantes :")
            for libelle, fait in etapes:
                st.markdown((":green[:material/check_circle:] " if fait else ":gray[:material/radio_button_unchecked:] ")
                            + libelle)
            st.button("Ouvrir le plan de chargement", type="primary", icon=":material/arrow_forward:",
                      on_click=aller, args=("optimisation", "Plan de chargement"))
    else:
        df_final = S.df_voyages_valides.copy()
        if "Chauffeur" not in df_final.columns:
            df_final["Chauffeur"] = df_final["Chauffeur attribué"] if "Chauffeur attribué" in df_final.columns \
                else "À attribuer"

        g1, g2 = st.columns([3, 2], gap="medium")
        with g1, st.container(border=True):
            rep = df_plan.groupby("Zone").size().reset_index(name="Véhicules")
            rep = rep.assign(_n=rep["Zone"].map(m.num_zone)).sort_values("_n", ascending=False)
            fig = px.bar(rep, x="Véhicules", y="Zone", orientation="h", text="Véhicules",
                         color_discrete_sequence=[BLEU])
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig = style_fig(fig, "Véhicules par zone", hauteur=max(260, 46 + 34 * len(rep)))
            fig.update_xaxes(visible=False)
            fig.update_yaxes(showgrid=False)
            graphique(fig)
        with g2, st.container(border=True):
            types = df_plan.apply(lambda r: "Camion loué" if m.est_camion(r) else "Estafette", axis=1)
            rep_t = types.value_counts().reset_index()
            rep_t.columns = ["Type", "Nombre"]
            fig = px.pie(rep_t, values="Nombre", names="Type", hole=0.6,
                         color="Type", color_discrete_map={"Estafette": BLEU, "Camion loué": "#F59E0B"})
            fig.update_traces(textinfo="value", textposition="inside", sort=False)
            fig = style_fig(fig, "Répartition par type de véhicule", hauteur=max(260, 46 + 34 * len(rep)))
            fig.update_layout(showlegend=True, legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center"))
            graphique(fig)

        section("Planning détaillé", f"{len(df_final)} voyage(s) validé(s)")
        df_aff = m.trier_par_zone(df_final)
        tableau(pd.DataFrame({
            "Code voyage": df_aff["Code voyage"], "Zone": df_aff["Zone"], "Voyage": df_aff["Véhicule N°"],
            "Véhicule": df_aff.get("Véhicule attribué"), "Chauffeur": df_aff["Chauffeur"],
            "Clients": df_aff["Client(s) inclus"], "BL": df_aff["BL inclus"],
            "Poids": df_aff["Poids total chargé"].map(fmt_kg), "Volume": df_aff["Volume total chargé"].map(fmt_m3),
            "Occupation": df_aff["Taux d'occupation (%)"].map(occupation),
        }), multiligne={"Zone": SEP_LISTE, "Clients": SEP_VIRGULE, "BL": SEP_LISTE},
            numeriques=("Poids", "Volume"), html_cols=("Occupation",), nowrap=("Code voyage", "Voyage", "Véhicule", "BL"),
            hauteur_max=560)

        section("Export du planning complet", "Fichier Excel multi-onglets (planning, analyses, synthèse)")
        with st.container(border=True):
            a, b, c = st.columns([2, 1, 1], vertical_alignment="bottom")
            nom = a.text_input("Nom du fichier", value=f"Planning_Livraisons_{datetime.now(TZ):%Y%m%d_%H%M}")
            b.button("Générer le fichier", icon=":material/build:", width="stretch",
                     on_click=generer_export, args=(nom.strip() or "Planning_Livraisons", df_final))
            with c:
                if S.export_planning:
                    bouton_telechargement("Télécharger", S.export_planning[1], S.export_planning[0],
                                          key="dl_planning", primaire=True)
                else:
                    st.button("Télécharger", icon=":material/download:", width="stretch", disabled=True,
                              key="dl_planning_off")

    st.divider()
    try:
        section_rapports()
    except ImportError as e:
        st.info(f"Rapports avancés indisponibles : {e}")

    st.divider()
    with st.popover("Réinitialiser l'optimisation", icon=":material/restart_alt:"):
        st.markdown("Les décisions de location, transferts, validations, attributions et codes seront effacés. "
                    "Les fichiers importés sont conservés.")
        st.button("Confirmer la réinitialisation", type="primary", on_click=reinitialiser_optimisation)


# =====================================================
# BARRE LATÉRALE ET NAVIGATION
# =====================================================
PAGES = [
    ("import", "Importation", ":material/upload_file:", page_import),
    ("analyse", "Analyse", ":material/query_stats:", page_analyse),
    ("optimisation", "Optimisation", ":material/local_shipping:", page_optimisation),
    ("finalisation", "Planning & KPIs", ":material/event_available:", page_finalisation),
]


@st.cache_data
def logo_base64(chemin="logo.png"):
    with open(chemin, "rb") as f:
        return base64.b64encode(f.read()).decode()


def nouvelle_importation():
    ver = S.ver_upload + 1
    for k in list(S.keys()):
        del S[k]
    for k, v in ETAT_INITIAL.items():
        S[k] = copy.deepcopy(v)
    S.ver_upload = ver


def etape_faite(cle):
    if cle in ("import", "analyse"):
        return S.data_processed
    if cle == "optimisation":
        return planning_pret()
    return S.export_planning is not None


def barre_laterale():
    with st.sidebar:
        logo = f'<img src="data:image/png;base64,{logo_base64()}" alt="">' if os.path.exists("logo.png") else ""
        st.markdown(f'<div class="brand">{logo}<div><div class="brand-name">Planning Livraisons</div>'
                    f'<div class="brand-sub">Préparation des chargements ADV</div></div></div>',
                    unsafe_allow_html=True)
        st.markdown('<div class="side-label">Parcours</div>', unsafe_allow_html=True)
        for i, (cle, libelle, icone, _) in enumerate(PAGES, start=1):
            courant = S.page == cle
            if not courant and etape_faite(cle):
                icone = ":material/check_circle:"
            st.button(f"{i}. {libelle}", key=f"nav_{cle}", icon=icone, width="stretch",
                      type="primary" if courant else "tertiary",
                      disabled=cle != "import" and not S.data_processed,
                      on_click=aller, args=(cle,))

        if S.data_processed:
            n_val = sum(S.validations.get(v) == "Oui" for v in S.df_voyages["Véhicule N°"])
            codes = "oui" if planning_pret() else "non"
            st.markdown('<div class="side-label">Session en cours</div>', unsafe_allow_html=True)
            st.markdown(
                '<div class="side-stats">'
                f'<span>Bons de livraison</span><b>{fmt_int(S.df_grouped_zone["No livraison"].nunique())}</b>'
                f'<span>Voyages</span><b>{fmt_int(len(S.df_voyages))}</b>'
                f'<span>Camions loués</span><b>{fmt_int(m.nb_camions(S.df_voyages))}</b>'
                f'<span>Voyages validés</span><b>{fmt_int(n_val)}</b>'
                f'<span>Codes voyage</span><b>{codes}</b></div>', unsafe_allow_html=True)
            st.write("")
            st.download_button("Sauvegarder la session", data=exporter_session(), width="stretch",
                               file_name=f"session_planning_{datetime.now(TZ):%Y%m%d_%H%M}.json",
                               mime="application/json", icon=":material/save:", on_click="ignore",
                               help="Télécharge l'état complet pour le reprendre plus tard (page Importation).")
            with st.popover("Nouvelle importation", icon=":material/add:"):
                st.markdown("Toutes les données et décisions de la session seront effacées.")
                st.button("Tout effacer et recommencer", type="primary", on_click=nouvelle_importation)

        st.divider()
        st.markdown('<div class="side-foot">Développé par Zaineb KCHAOU<br>Zaineb.KCHAOU@sopal.com</div>',
                    unsafe_allow_html=True)


def remonter_en_haut():
    """Après un changement de page, revient en haut (Streamlit garde sinon la position de défilement)."""
    if S.pop("_haut_de_page", False):
        import streamlit.components.v1 as components
        components.html("<script>const d=window.parent.document;"
                        "(d.querySelector('[data-testid=\"stMain\"]')||d.querySelector('section.main')||d.body)"
                        ".scrollTo({top:0});window.parent.scrollTo({top:0});</script>", height=0)


def main():
    barre_laterale()
    remonter_en_haut()
    pages = {cle: fn for cle, _, _, fn in PAGES}
    pages.get(S.page, page_import)()


if __name__ == "__main__":
    main()
