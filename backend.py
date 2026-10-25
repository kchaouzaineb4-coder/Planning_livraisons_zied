import pandas as pd
import math
import numpy as np
import unicodedata

# --- Constantes pour la location de camion ---
SEUIL_POIDS = 3000.0    # kg
SEUIL_VOLUME = 9.216    # m³
CAPACITE_POIDS_ESTAFETTE = 1550  # kg
CAPACITE_VOLUME_ESTAFETTE = 4.608  # m³

# --- Sites autorisés ---
SITES_AUTORISES = ["SA1", "AP1", "SO1", "TC1", "SO2"]

# NOUVELLES CONSTANTES POUR LES TYPES DE CAMIONS
CAPACITE_POIDS_CAMION_5T = 5000  # kg
CAPACITE_VOLUME_CAMION_5T = 20.0  # m³
CAPACITE_POIDS_CAMION_10T = 10000  # kg
CAPACITE_VOLUME_CAMION_10T = 40.0  # m³

CAMION_CODE = "CAMION-LOUE"

# REMPLACER LES ANCIENNES CONSTANTES PAR DES FONCTIONS
# Les fonctions suivantes retournent les capacités selon le type de camion
def get_capacite_poids_camion(truck_type="5 tonnes"):
    """Retourne la capacité poids selon le type de camion."""
    if truck_type == "10 tonnes":
        return CAPACITE_POIDS_CAMION_10T
    else:  # 5 tonnes par défaut
        return CAPACITE_POIDS_CAMION_5T

def get_capacite_volume_camion(truck_type="5 tonnes"):
    """Retourne la capacité volume selon le type de camion."""
    if truck_type == "10 tonnes":
        return CAPACITE_VOLUME_CAMION_10T
    else:  # 5 tonnes par défaut
        return CAPACITE_VOLUME_CAMION_5T


# =====================================================
# CLASSE PRINCIPALE DE TRAITEMENT DES LIVRAISONS
# =====================================================
class DeliveryProcessor:
    def __init__(self):
        self.df_livraisons_original = None
    
    def process_delivery_data(self, liv_file, ydlogist_file, wcliegps_file):
        """Traite les fichiers d'entrée et retourne les DataFrames résultants."""
        try:
            # Lecture des fichiers
            df_liv = self._load_livraisons(liv_file)
            df_yd = self._load_ydlogist(ydlogist_file)
            df_clients = self._load_wcliegps(wcliegps_file)

            # Filtrage des données
            df_liv = self._filter_initial_data(df_liv)

            # Calcul Poids & Volume
            df_poids = self._calculate_weights(df_liv)
            df_vol = self._calculate_volumes(df_liv, df_yd)

            # Fusionner poids + volume
            df_merged = self._merge_delivery_data(df_poids, df_vol)

            # Ajouter Client, Ville et Représentant
            df_final = self._add_city_client_info(df_merged, df_clients)

            # Calcul Volume total en m3
            df_final["Volume de l'US"] = pd.to_numeric(df_final["Volume de l'US"], errors='coerce').fillna(0) / 1_000_000
            df_final["Volume total"] = df_final["Volume de l'US"] * df_final["Quantité livrée US"]

            # Regroupement par ville et client (pour l'affichage "Livraisons Client/Ville")
            df_grouped, df_city = self._group_data(df_final)

            # Calcul du besoin en estafette par ville
            df_city = self._calculate_estafette_need(df_city)

            # Nouveau tableau : ajout Zone
            df_grouped_zone = self._add_zone(df_grouped.rename(columns={"Client": "Client de l'estafette"}))

            # Filtrer les livraisons avec "Zone inconnue"
            df_grouped_zone = df_grouped_zone[df_grouped_zone["Zone"] != "Zone inconnue"].copy()
            
            # Groupement par zone
            df_zone = self._group_by_zone(df_grouped_zone)
            
            # Calcul du besoin en estafette par zone
            df_zone = self._calculate_estafette_need(df_zone)

            # Calcul des voyages optimisés 
            df_optimized_estafettes = self._calculate_optimized_estafette(df_grouped_zone)

            # 🆕 CORRECTION : Stocker les données originales du tableau "Livraisons par Client & Ville + Zone"
            self.df_livraisons_original = df_grouped_zone.copy()

            # 🆕 CORRECTION : Retourner 6 valeurs
            return df_grouped, df_city, df_grouped_zone, df_zone, df_optimized_estafettes, self.df_livraisons_original

        except Exception as e:
            raise Exception(f"❌ Erreur lors du traitement des données : {str(e)}")

    # =====================================================
    # MÉTHODES AUXILIAIRES
    # =====================================================
    def _load_livraisons(self, liv_file):
        df = pd.read_excel(liv_file)
        if 'N° BON LIVRAISON' in df.columns:
            df.rename(columns={'N° BON LIVRAISON': 'No livraison'}, inplace=True)
        if len(df.columns) > 4:
            df.rename(columns={df.columns[4]: "Quantité livrée US"}, inplace=True)
        return df

    def _load_ydlogist(self, file_path):
        df = pd.read_excel(file_path)
        if len(df.columns) > 16:
            df.rename(columns={df.columns[16]: "Unité Volume"}, inplace=True)
        if len(df.columns) > 13:
            df.rename(columns={df.columns[13]: "Poids de l'US"}, inplace=True)
        return df

    def _load_wcliegps(self, wcliegps_file):
        df_clients = pd.read_excel(wcliegps_file)
        if len(df_clients.columns) > 16:
            df_clients.rename(columns={df_clients.columns[16]: "Représentant"}, inplace=True)
        required_cols = ["Client", "Ville", "Représentant"]
        for col in required_cols:
            if col not in df_clients.columns:
                raise ValueError(f"La colonne '{col}' est manquante dans le fichier clients.")
        return df_clients[["Client", "Ville", "Représentant"]].copy()

    def _filter_initial_data(self, df):
        clients_exclus = [
            "AMECAP", "SANA", "SOPAL", "SOPALGAZ", "SOPALSERV", "SOPALTEC",
            "SOPALALG", "AQUAPLAST", "WINOX", "QUIVEM", "SANISTONE",
            "SOPAMAR", "SOPALAFR", "SOPALINTER", "PERSOGSO"
        ]

        # Vérifier que la colonne existe
        if "Site vente" not in df.columns:
            raise ValueError(
                "La colonne 'Site vente' est introuvable dans le fichier Livraisons. "
                f"Colonnes disponibles : {list(df.columns)}"
            )

        # Filtre sur les sites : on garde seulement SA1, AP1, SO1, TC1, SO2
        sites_normalises = df["Site vente"].astype(str).str.strip().str.upper()
        df = df[sites_normalises.isin(SITES_AUTORISES)]

        # Filtres existants
        df = df[(df["Type livraison"] != "SDC") & (~df["Client commande"].isin(clients_exclus))]
        return df.copy()

    def _calculate_weights(self, df):
        df["Poids de l'US"] = pd.to_numeric(df["Poids de l'US"].astype(str).str.replace(",", ".")
                                           .str.replace(r"[^\d.]", "", regex=True), errors="coerce").fillna(0)
        df["Quantité livrée US"] = pd.to_numeric(df["Quantité livrée US"], errors="coerce").fillna(0)
        df["Poids total"] = df["Quantité livrée US"] * df["Poids de l'US"]
        return df[["No livraison", "Article", "Client commande", "Poids total", "Quantité livrée US", "Poids de l'US"]]

    def _calculate_volumes(self, df_liv, df_art):
        colonnes = ["Article", "Volume de l'US", "Unité Volume"]
        if "Désignation 1" in df_art.columns:
            colonnes.append("Désignation 1")

        df_art_sel = df_art[colonnes].copy()
        if "Désignation 1" not in df_art_sel.columns:
            df_art_sel["Désignation 1"] = ""

        df_art_sel["Volume de l'US"] = pd.to_numeric(
            df_art_sel["Volume de l'US"].astype(str).str.replace(",", "."), errors="coerce")
        # Un seul enregistrement par article, pour ne pas multiplier les lignes
        df_art_sel = df_art_sel.drop_duplicates(subset="Article")

        df_liv_sel = df_liv[["No livraison", "Article", "Client commande"]].copy()
        df_liv_sel = df_liv_sel.reset_index(drop=True)
        df_vol = pd.merge(df_liv_sel, df_art_sel, on="Article", how="left")
        df_vol.index = df_liv_sel.index

        df_vol["Désignation 1"] = df_vol["Désignation 1"].fillna("").astype(str)
        return df_vol[["Volume de l'US", "Unité Volume", "Désignation 1"]]

    def _merge_delivery_data(self, df_poids, df_vol):
        # Même nombre de lignes, même ordre : on colle les colonnes côte à côte
        df_poids = df_poids.reset_index(drop=True)
        df_vol = df_vol.reset_index(drop=True)
        df_poids = df_poids.drop(columns=["Poids de l'US"], errors="ignore")
        return pd.concat([df_poids, df_vol], axis=1)

    def _add_city_client_info(self, df, df_clients):
        return pd.merge(df, df_clients[["Client", "Ville", "Représentant"]],
                       left_on="Client commande", right_on="Client", how="left")

    def _group_data(self, df):
        df_grouped = df.groupby(["No livraison", "Client", "Ville", "Représentant"], as_index=False).agg({
            "Article": lambda x: ", ".join(x.astype(str)),
            "Désignation 1": lambda x: "<br>".join(
                x.astype(str).str.replace(r"[,\r\n]+", " ", regex=True).str.strip()
            ),
            "Poids total": "sum",
            "Volume total": "sum"
        }).rename(columns={"Désignation 1": "Désignation"})

        df_city = df_grouped.groupby("Ville", as_index=False).agg({
            "Poids total": "sum",
            "Volume total": "sum",
            "No livraison": "count"
        }).rename(columns={"No livraison": "Nombre de BLs"})  # ← MODIFICATION ICI
        return df_grouped, df_city

    def _calculate_estafette_need(self, df):
        if "Poids total" in df.columns and "Volume total" in df.columns:
            df["Besoin estafette (poids)"] = df["Poids total"].apply(lambda p: math.ceil(p / CAPACITE_POIDS_ESTAFETTE))
            df["Besoin estafette (volume)"] = df["Volume total"].apply(lambda v: math.ceil(v / CAPACITE_VOLUME_ESTAFETTE))
            df["Besoin estafette réel"] = df[["Besoin estafette (poids)", "Besoin estafette (volume)"]].max(axis=1)
        return df

    def _add_zone(self, df):
        """
        Ajoute la colonne Zone (numéro + nom).
        - Villes normalisées (majuscules, sans accents, tirets/espaces uniformisés).
        - Zone 7_Sud : liste de clients précise (quelle que soit la ville).
        - Zone 7_Nord : les autres clients de Sfax.
        """
        ZONE_7_SUD = "Zone 7_Sud - Sfax Sud"
        ZONE_7_NORD = "Zone 7_Nord - Sfax Nord"

        zones = {
            "Zone 1 - Tunis Centre": ["BARDO", "TUNIS", "ZAHROUNI", "MANOUBA"],
            "Zone 2 - Ben Arous": ["BEN AROUS"],
            "Zone 3 - Ariana": ["ARIANA", "SOUKRA", "AOUINA", "CHARGUIA"],
            "Zone 4 - Sousse-Kairouan": ["MSAKEN", "SOUSSE", "KAIROUAN", "KAIROUAN SUD"],
            "Zone 5 - Nabeul": ["NABEUL"],
            "Zone 6 - Bizerte": ["BIZERTE"],
            "Zone 8 - Monastir-Mahdia": ["MAHDIA", "MONASTIR", "JAMMEL"],
            "Zone 9 - Intérieur": ["JANDOUBA", "JENDOUBA", "GAFSA", "SIDI BOUZID"],
            "Zone 10 - Grand Sud": ["DJERBA", "GABES", "MEDENINE", "DJERBA MIDOUN -MEDENINE"],
        }

        CLIENTS_ZONE_7_SUD = {"CEGA2", "CBHC", "EBCH", "PYRAMIDE", "STEMSEG", "TUNISACE1"}
        VILLES_SFAX = {"SFAX", "SFAX VILLE", "SFAX SUD", "SFAX NORD"}

        def normaliser(texte):
            """Majuscules, sans accents, tirets/espaces uniformisés."""
            texte = str(texte).strip().upper()
            texte = unicodedata.normalize("NFD", texte)
            texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
            texte = texte.replace("-", " ")
            return " ".join(texte.split())

        ville_vers_zone = {normaliser(v): z for z, villes in zones.items() for v in villes}
        clients_sud_norm = {normaliser(c) for c in CLIENTS_ZONE_7_SUD}
        villes_sfax_norm = {normaliser(v) for v in VILLES_SFAX}

        def get_zone(row):
            ville_norm = normaliser(row.get("Ville", ""))
            client_norm = normaliser(row.get("Client de l'estafette", row.get("Client", "")))

            # 1) Clients de la liste Sud -> Zone 7_Sud
            if client_norm in clients_sud_norm:
                return ZONE_7_SUD

            # 2) Autres clients de Sfax -> Zone 7_Nord
            if ville_norm in villes_sfax_norm:
                return ZONE_7_NORD

            # 3) Cas standard : ville -> zone
            return ville_vers_zone.get(ville_norm, "Zone inconnue")

        df["Zone"] = df.apply(get_zone, axis=1)
        return df

    

    def _group_by_zone(self, df_grouped_zone):
        df_zone = df_grouped_zone.groupby("Zone", as_index=False).agg({
            "Poids total": "sum",
            "Volume total": "sum",
            "No livraison": "count"
        }).rename(columns={"No livraison": "Nombre livraisons"})
        return df_zone

    def _calculate_optimized_estafette(self, df_grouped_zone):
        """
        Calcule les voyages optimisés avec règle spéciale pour Zone 7.
        
        MODIFICATION MAJEURE : Regroupement STRICT par client.
        Un client ne peut JAMAIS être réparti sur plusieurs estafettes de la même zone,
        SAUF si son poids/volume total dépasse la capacité d'une estafette.
        """
        resultats = []
        estafette_num = 1
        CAPACITE_POIDS_ESTAFETTE = 1550  # kg
        CAPACITE_VOLUME_ESTAFETTE = 4.608  # m³
        MAX_VOYAGES_ZONE7 = 3

        for zone, group in df_grouped_zone.groupby("Zone"):
            # =====================================================
            # ÉTAPE 1 : Regrouper les BLs par client (DONNÉES CONSOLIDÉES)
            # =====================================================
            clients_data = {}
            for idx, row in group.iterrows():
                client = str(row["Client de l'estafette"])
                if client not in clients_data:
                    clients_data[client] = {
                        "bls": [],
                        "poids_total": 0.0,
                        "volume_total": 0.0,
                        "representants": set()
                    }
                
                clients_data[client]["bls"].append({
                    "bl": str(row["No livraison"]),
                    "poids": float(row["Poids total"]),
                    "volume": float(row["Volume total"])
                })
                clients_data[client]["poids_total"] += float(row["Poids total"])
                clients_data[client]["volume_total"] += float(row["Volume total"])
                clients_data[client]["representants"].add(str(row["Représentant"]))
            
            # Convertir en liste
            clients_list = []
            for client, data in clients_data.items():
                clients_list.append({
                    "client": client,
                    "bls": data["bls"],
                    "poids_total": data["poids_total"],
                    "volume_total": data["volume_total"],
                    "representants": data["representants"]
                })
            
            # Trier les clients par poids total décroissant
            clients_list.sort(key=lambda x: x["poids_total"], reverse=True)
            
            # =====================================================
            # RÈGLE SPÉCIALE POUR ZONE 7
            # =====================================================
            if zone == "Zone 7" and len(clients_list) > 0:
                voyages = []
                voyage_actuel = {
                    "poids": 0.0, "volume": 0.0, "bls": [],
                    "clients": set(), "representants": set()
                }
                
                for client_data in clients_list:
                    client = client_data["client"]
                    
                    # ✅ VÉRIFIER SI LE CLIENT ENTIER RENTRE DANS LE VOYAGE ACTUEL
                    if (voyage_actuel["poids"] + client_data["poids_total"] <= CAPACITE_POIDS_ESTAFETTE and 
                        voyage_actuel["volume"] + client_data["volume_total"] <= CAPACITE_VOLUME_ESTAFETTE):
                        # Ajouter TOUT le client d'un coup
                        voyage_actuel["poids"] += client_data["poids_total"]
                        voyage_actuel["volume"] += client_data["volume_total"]
                        voyage_actuel["clients"].add(client)
                        voyage_actuel["representants"].update(client_data["representants"])
                        for bl_info in client_data["bls"]:
                            voyage_actuel["bls"].append(bl_info["bl"])
                    else:
                        # ✅ SI LE CLIENT NE RENTRE PAS EN ENTIER : NE PAS LE RÉPARTIR
                        # Créer un nouveau voyage pour ce client
                        if voyage_actuel["bls"]:
                            voyages.append(voyage_actuel)
                        
                        # Vérifier si on a atteint le max de voyages
                        if len(voyages) >= MAX_VOYAGES_ZONE7:
                            # Forcer l'ajout dans le dernier voyage (cas extrême)
                            dernier_voyage = voyages[-1]
                            dernier_voyage["poids"] += client_data["poids_total"]
                            dernier_voyage["volume"] += client_data["volume_total"]
                            dernier_voyage["clients"].add(client)
                            dernier_voyage["representants"].update(client_data["representants"])
                            for bl_info in client_data["bls"]:
                                dernier_voyage["bls"].append(bl_info["bl"])
                            voyage_actuel = {"poids": 0, "volume": 0, "bls": [], "clients": set(), "representants": set()}
                        else:
                            # Nouveau voyage avec TOUT le client
                            voyage_actuel = {
                                "poids": client_data["poids_total"],
                                "volume": client_data["volume_total"],
                                "bls": [bl_info["bl"] for bl_info in client_data["bls"]],
                                "clients": {client},
                                "representants": set(client_data["representants"])
                            }
                
                # Ajouter le dernier voyage
                if voyage_actuel["bls"]:
                    voyages.append(voyage_actuel)
                
                for i, voyage in enumerate(voyages, 1):
                    if i > MAX_VOYAGES_ZONE7:
                        break
                    clients_str = ", ".join(sorted(list(voyage["clients"])))
                    representants_str = ", ".join(sorted(list(voyage["representants"])))
                    resultats.append([
                        zone, estafette_num, voyage["poids"], voyage["volume"],
                        clients_str, representants_str, ";".join(voyage["bls"]), i
                    ])
                
                estafette_num += 1
                
            else:
                # =====================================================
                # ALGORITHME STANDARD AVEC REGROUPEMENT STRICT PAR CLIENT
                # =====================================================
                estafettes = []
                # Track les clients déjà placés pour éviter les doublons
                clients_placed = set()
                
                for client_data in clients_list:
                    client = client_data["client"]
                    
                    if client in clients_placed:
                        continue  # Sécurité : ne pas traiter 2 fois le même client
                    
                    poids_client = client_data["poids_total"]
                    volume_client = client_data["volume_total"]
                    
                    # ✅ CAS 1 : Le client tient EN ENTIER dans une estafette existante
                    placed_entier = False
                    for e in estafettes:
                        if (e["poids"] + poids_client <= CAPACITE_POIDS_ESTAFETTE and 
                            e["volume"] + volume_client <= CAPACITE_VOLUME_ESTAFETTE):
                            # Ajouter TOUT le client
                            e["poids"] += poids_client
                            e["volume"] += volume_client
                            for bl_info in client_data["bls"]:
                                e["bls"].append(bl_info["bl"])
                            e["clients"].add(client)
                            e["representants"].update(client_data["representants"])
                            placed_entier = True
                            break
                    
                    if placed_entier:
                        clients_placed.add(client)
                        continue
                    
                    # ✅ CAS 2 : Le client ne tient pas en entier
                    # Vérifier si c'est parce qu'il dépasse la capacité d'UNE estafette
                    if poids_client > CAPACITE_POIDS_ESTAFETTE or volume_client > CAPACITE_VOLUME_ESTAFETTE:
                        # Le client est trop gros → répartition INÉVITABLE
                        # On essaie de minimiser le nombre d'estafettes
                        bls_restants = client_data["bls"].copy()
                        
                        while bls_restants:
                            # Chercher une estafette existante avec de la place
                            placed_bl = False
                            for e in estafettes:
                                if (e["poids"] + bls_restants[0]["poids"] <= CAPACITE_POIDS_ESTAFETTE and 
                                    e["volume"] + bls_restants[0]["volume"] <= CAPACITE_VOLUME_ESTAFETTE):
                                    bl_info = bls_restants.pop(0)
                                    e["poids"] += bl_info["poids"]
                                    e["volume"] += bl_info["volume"]
                                    e["bls"].append(bl_info["bl"])
                                    e["clients"].add(client)
                                    e["representants"].update(client_data["representants"])
                                    placed_bl = True
                                    break
                            
                            if not placed_bl:
                                # Créer une nouvelle estafette
                                bl_info = bls_restants.pop(0)
                                estafettes.append({
                                    "poids": bl_info["poids"],
                                    "volume": bl_info["volume"],
                                    "bls": [bl_info["bl"]],
                                    "clients": {client},
                                    "representants": set(client_data["representants"]),
                                    "num_global": estafette_num
                                })
                                estafette_num += 1
                        
                        clients_placed.add(client)
                    else:
                        # ✅ CAS 3 : Le client pourrait tenir dans une estafette VIDE
                        # → Créer une nouvelle estafette dédiée à ce client
                        estafettes.append({
                            "poids": poids_client,
                            "volume": volume_client,
                            "bls": [bl_info["bl"] for bl_info in client_data["bls"]],
                            "clients": {client},
                            "representants": set(client_data["representants"]),
                            "num_global": estafette_num
                        })
                        estafette_num += 1
                        clients_placed.add(client)
                
                # Formater les résultats
                for e in estafettes:
                    clients_str = ", ".join(sorted(list(e["clients"])))
                    representants_str = ", ".join(sorted(list(e["representants"])))
                    resultats.append([
                        zone, e["num_global"], e["poids"], e["volume"],
                        clients_str, representants_str, ";".join(e["bls"]), 1
                    ])
        
        # Créer le DataFrame
        df_estafettes = pd.DataFrame(resultats, columns=[
            "Zone", "Estafette N°", "Poids total chargé", "Volume total chargé", 
            "Client(s) inclus", "Représentant(s) inclus", "BL inclus", "Numéro Voyage"
        ])
        
        # Calcul du taux d'occupation
        df_estafettes["Taux Poids (%)"] = (df_estafettes["Poids total chargé"] / CAPACITE_POIDS_ESTAFETTE) * 100
        df_estafettes["Taux Volume (%)"] = (df_estafettes["Volume total chargé"] / CAPACITE_VOLUME_ESTAFETTE) * 100
        df_estafettes["Taux d'occupation (%)"] = df_estafettes[["Taux Poids (%)", "Taux Volume (%)"]].max(axis=1).round(2)
        
        df_estafettes["Location_camion"] = False
        df_estafettes["Location_proposee"] = False
        df_estafettes["Code Véhicule"] = "ESTAFETTE"
        
        def get_vehicle_name(row):
            if row["Zone"] == "Zone 7":
                return f"E{row['Estafette N°']}-Voyage {row['Numéro Voyage']}"
            else:
                return f"E{row['Estafette N°']}"
        
        df_estafettes["Camion N°"] = df_estafettes.apply(get_vehicle_name, axis=1)
        df_estafettes = df_estafettes.drop(columns=["Taux Poids (%)", "Taux Volume (%)", "Numéro Voyage"]) 
        
        return df_estafettes

# =====================================================
# CLASSE DE GESTION DE LA LOCATION DE CAMIONS
# =====================================================
class TruckRentalProcessor:
    def __init__(self, df_optimized, df_livraisons_original):
        """Initialise avec le DataFrame optimisé ET les données originales."""
        self.df_base = self._initialize_rental_columns(df_optimized.copy())
        self.df_livraisons_original = df_livraisons_original.copy()
        self._next_camion_num = self.df_base[self.df_base["Code Véhicule"] == CAMION_CODE].shape[0] + 1
        self.truck_type = "5 tonnes"  # Valeur par défaut
    
    def _get_capacites_camion(self, truck_type="5 tonnes"):
        """Retourne les capacités selon le type de camion."""
        return get_capacite_poids_camion(truck_type), get_capacite_volume_camion(truck_type)
    
    def _initialize_rental_columns(self, df):
        """Initialise les colonnes pour la gestion de la location."""
        df.rename(columns={
            "Poids total chargé": "Poids total",
            "Volume total chargé": "Volume total"
        }, inplace=True)

        if "Location_camion" not in df.columns:
            df["Location_camion"] = False
        if "Location_proposee" not in df.columns:
            df["Location_proposee"] = False
        if "Code Véhicule" not in df.columns:
            df["Code Véhicule"] = "ESTAFETTE"
        if "Camion N°" not in df.columns:
            df["Camion N°"] = df["Estafette N°"].apply(lambda x: f"E{int(x)}" if pd.notna(x) and x != 0 else "À Optimiser")
        
        mask_camion_loue = df["Code Véhicule"] == CAMION_CODE
        if mask_camion_loue.any():
            df.loc[mask_camion_loue, "Camion N°"] = [f"C{i+1}" for i in range(mask_camion_loue.sum())]

        df['BL inclus'] = df['BL inclus'].astype(str)
        df["Estafette N°"] = pd.to_numeric(df["Estafette N°"], errors='coerce').fillna(99999).astype(int)

        return df

    def _get_client_totals_from_original_data(self):
        """Extrait les totaux réels des clients depuis le tableau 'Livraisons par Client & Ville + Zone'."""
        try:
            # Vérifier que les colonnes nécessaires existent
            required_cols = ["Client de l'estafette", "Poids total", "Volume total"]
            missing_cols = [col for col in required_cols if col not in self.df_livraisons_original.columns]
            
            if missing_cols:
                print(f"⚠️ Colonnes manquantes dans les données originales: {missing_cols}")
                return pd.DataFrame(columns=["Client", "Poids total (kg)", "Volume total (m³)"])
            
            # Grouper par client pour obtenir les totaux RÉELS
            df_client_totals = self.df_livraisons_original.groupby("Client de l'estafette").agg({
                "Poids total": "sum",
                "Volume total": "sum"
            }).reset_index()
            
            df_client_totals = df_client_totals.rename(columns={
                "Client de l'estafette": "Client",
                "Poids total": "Poids total (kg)",
                "Volume total": "Volume total (m³)"
            })
            
            return df_client_totals
            
        except Exception as e:
            print(f"❌ Erreur lors de l'extraction des totaux clients: {e}")
            return pd.DataFrame(columns=["Client", "Poids total (kg)", "Volume total (m³)"])

    def detecter_propositions(self):
        """Détecte les propositions en utilisant les totaux RÉELS du tableau original."""
        # Récupérer les totaux réels des clients
        df_client_totals = self._get_client_totals_from_original_data()
        
        if df_client_totals.empty:
            return pd.DataFrame()

        # Exclure les clients déjà traités
        processed_clients = self.df_base[self.df_base["Location_proposee"]]["Client(s) inclus"].unique()
        
        # Filtrer les clients non traités
        df_pending = df_client_totals[~df_client_totals["Client"].isin(processed_clients)].copy()
        
        if df_pending.empty:
            return pd.DataFrame()

        # Appliquer les seuils sur les totaux RÉELS
        propositions = df_pending[
            (df_pending["Poids total (kg)"] >= SEUIL_POIDS) | 
            (df_pending["Volume total (m³)"] >= SEUIL_VOLUME)
        ].copy()

        if propositions.empty:
            return pd.DataFrame()

        # Ajouter la colonne Raison
        def get_raison(row):
            raisons = []
            if row["Poids total (kg)"] >= SEUIL_POIDS:
                raisons.append(f"Poids ≥ {SEUIL_POIDS} kg")
            if row["Volume total (m³)"] >= SEUIL_VOLUME:
                raisons.append(f"Volume ≥ {SEUIL_VOLUME:.3f} m³")
            return " & ".join(raisons)

        propositions["Raison"] = propositions.apply(get_raison, axis=1)

        return propositions.sort_values(["Poids total (kg)", "Volume total (m³)"], ascending=False).reset_index(drop=True)

    def get_details_client(self, client):
        """Affiche les détails avec les totaux RÉELS du tableau original."""
        try:
            # Récupérer les totaux RÉELS du client depuis les données originales
            client_data_original = self.df_livraisons_original[
                self.df_livraisons_original["Client de l'estafette"] == client
            ]
            
            if client_data_original.empty:
                return f"Aucune donnée pour {client}", pd.DataFrame()

            # Calculer les totaux RÉELS
            total_poids_reel = client_data_original["Poids total"].sum()
            total_volume_reel = client_data_original["Volume total"].sum()
            
            # Récupérer les BLs du client
            bls_client = client_data_original["No livraison"].unique()
            
            # Trouver les estafettes qui contiennent ces BLs
            details_estafettes = []
            for _, row in self.df_base.iterrows():
                bls_in_vehicle = str(row["BL inclus"]).split(';')
                bls_commun = set(map(str, bls_client)) & set(bls_in_vehicle)
                
                if bls_commun:
                    details_estafettes.append({
                        'Zone': row['Zone'],
                        'Camion N°': row['Camion N°'],
                        'Poids total': f"{row['Poids total']:.3f} kg",
                        'Volume total': f"{row['Volume total']:.3f} m³",
                        'BL inclus': row['BL inclus'],
                        'Taux d\'occupation (%)': f"{row['Taux d\'occupation (%)']:.2f}%"
                    })
            
            # Déterminer l'état
            etat = "Non décidée"
            client_in_base = self.df_base[self.df_base["Client(s) inclus"].str.contains(client, na=False)]
            
            if not client_in_base.empty:
                if client_in_base["Location_camion"].any():
                    etat = "Location ACCEPTÉE"
                elif client_in_base["Location_proposee"].any():
                    etat = "Proposition REFUSÉE"
            
            resume = f"Client {client} — Poids total RÉEL : {total_poids_reel:.1f} kg ; Volume total RÉEL : {total_volume_reel:.3f} m³ | État : {etat}"
            
            df_details = pd.DataFrame(details_estafettes)
            return resume, df_details
            
        except Exception as e:
            print(f"❌ Erreur dans get_details_client: {e}")
            return f"Erreur avec le client {client}", pd.DataFrame()

    def appliquer_location(self, client, accepter, truck_type="5 tonnes"):
        """Applique la décision de location pour un client avec réoptimisation automatique."""
        try:
            # Stocker le type de camion
            self.truck_type = truck_type
            
            # Récupérer les capacités pour ce type de camion
            capacite_poids, capacite_volume = self._get_capacites_camion(truck_type)
            
            # Utiliser les données originales pour trouver tous les BLs du client
            client_data_original = self.df_livraisons_original[
                self.df_livraisons_original["Client de l'estafette"] == client
            ]
            
            if client_data_original.empty:
                return False, "Client introuvable dans les données originales.", self.df_base

            # Récupérer tous les BLs du client
            bls_client = client_data_original["No livraison"].unique()
            
            df = self.df_base.copy()
            
            if accepter:
                # Récupérer les données consolidées pour le camion
                poids_total = client_data_original["Poids total"].sum()
                volume_total = client_data_original["Volume total"].sum()
                bl_concat = ";".join([str(bl) for bl in bls_client])
                representants = ";".join(sorted(client_data_original["Représentant"].astype(str).unique().tolist()))
                zones = ";".join(sorted(client_data_original["Zone"].astype(str).unique().tolist()))
                
                # Vérifier que les totaux ne dépassent pas la capacité du camion
                if poids_total > capacite_poids:
                    return False, f"❌ Le poids total ({poids_total:.1f} kg) dépasse la capacité du camion {truck_type} ({capacite_poids} kg).", self.df_base
                
                if volume_total > capacite_volume:
                    return False, f"❌ Le volume total ({volume_total:.3f} m³) dépasse la capacité du camion {truck_type} ({capacite_volume} m³).", self.df_base
                
                # Calcul du taux d'occupation du camion avec les capacités appropriées
                taux_poids = (poids_total / capacite_poids) * 100
                taux_volume = (volume_total / capacite_volume) * 100
                taux_occu = max(taux_poids, taux_volume)
                
                # Créer un nouveau voyage pour le camion loué
                camion_num_final = f"C{self._next_camion_num}"
                new_row = pd.DataFrame([{
                    "Zone": zones,
                    "Estafette N°": 0,
                    "Poids total": poids_total,
                    "Volume total": volume_total,
                    "BL inclus": bl_concat,
                    "Client(s) inclus": client,
                    "Représentant(s) inclus": representants,
                    "Location_camion": True,
                    "Location_proposee": True,
                    "Code Véhicule": CAMION_CODE,
                    "Camion N°": camion_num_final,
                    "Taux d'occupation (%)": taux_occu,
                    "Type_Camion": truck_type,  # Ajouter le type de camion
                    "Capacite_Poids": capacite_poids,
                    "Capacite_Volume": capacite_volume
                }])
                
                self._next_camion_num += 1
                
                # ÉTAPE 1: Identifier tous les BLs à garder (non transférés)
                bls_a_garder_total = []
                zones_affectees = set()
                
                for idx, row in df.iterrows():
                    if pd.notna(row["BL inclus"]):
                        bls_actuels = str(row["BL inclus"]).split(';')
                        # Garder seulement les BLs qui ne sont PAS du client à transférer
                        bls_a_garder = [bl for bl in bls_actuels if bl not in [str(b) for b in bls_client]]
                        bls_a_garder_total.extend(bls_a_garder)
                        
                        # Noter les zones affectées
                        if bls_a_garder:
                            zones_affectees.add(row["Zone"])
                
                # ÉTAPE 2: Réoptimiser COMPLÈTEMENT les estafettes pour chaque zone affectée
                df_estafettes_optimisees = self._reoptimiser_estafettes_par_zone(bls_a_garder_total, zones_affectees)
                
                # ÉTAPE 3: Combiner camions existants + nouvelles estafettes optimisées
                df_camions_existants = df[df["Code Véhicule"] == CAMION_CODE].copy()
                df_final = pd.concat([df_camions_existants, df_estafettes_optimisees, new_row], ignore_index=True)
                
                self.df_base = df_final
                return True, f"✅ Location ACCEPTÉE pour {client} avec camion {truck_type}. Commandes transférées vers {camion_num_final}. Réoptimisation des estafettes effectuée.", self.detecter_propositions()
            else:
                # Refuser la proposition - pas de changement dans l'optimisation
                mask_original = df["BL inclus"].apply(
                    lambda x: any(str(bl) in str(x).split(';') for bl in bls_client)
                )
                df.loc[mask_original, ["Location_proposee", "Location_camion", "Code Véhicule"]] = [True, False, "ESTAFETTE"]
                df.loc[mask_original, "Camion N°"] = df.loc[mask_original, "Estafette N°"].apply(lambda x: f"E{int(x)}")
                
                self.df_base = df
                return True, f"❌ Proposition REFUSÉE pour {client}. Les commandes restent en Estafettes.", self.detecter_propositions()
                
        except Exception as e:
            return False, f"❌ Erreur lors de l'application de la décision: {str(e)}", self.df_base

    def _reoptimiser_estafettes_par_zone(self, bls_a_garder, zones_affectees):
        """
        Réoptimise complètement les estafettes pour les BLs restants après transfert.
        
        MODIFICATION : Regroupement STRICT par client.
        Un client ne peut JAMAIS être réparti sur plusieurs estafettes de la même zone,
        SAUF si son poids/volume total dépasse la capacité d'une estafette.
        """
        try:
            if not bls_a_garder:
                return pd.DataFrame()
            
            # Récupérer les données complètes des BLs à garder
            df_bls_data = self.df_livraisons_original[
                self.df_livraisons_original["No livraison"].isin(bls_a_garder)
            ]
            
            if df_bls_data.empty:
                return pd.DataFrame()
            
            resultats_optimises = []
            CAPACITE_POIDS_ESTAFETTE = 1550  # kg
            CAPACITE_VOLUME_ESTAFETTE = 4.608  # m³
            MAX_VOYAGES_ZONE7 = 3
            
            # Pour chaque zone affectée
            for zone in zones_affectees:
                df_zone = df_bls_data[df_bls_data["Zone"] == zone]
                
                if df_zone.empty:
                    continue
                
                # =====================================================
                # ÉTAPE 1 : REGROUPER LES BLS PAR CLIENT
                # =====================================================
                clients_data = {}
                for idx, row in df_zone.iterrows():
                    client = str(row["Client de l'estafette"])
                    if client not in clients_data:
                        clients_data[client] = {
                            "bls": [],
                            "poids_total": 0.0,
                            "volume_total": 0.0,
                            "representants": set()
                        }
                    
                    clients_data[client]["bls"].append({
                        "bl": str(row["No livraison"]),
                        "poids": float(row["Poids total"]),
                        "volume": float(row["Volume total"])
                    })
                    clients_data[client]["poids_total"] += float(row["Poids total"])
                    clients_data[client]["volume_total"] += float(row["Volume total"])
                    clients_data[client]["representants"].add(str(row["Représentant"]))
                
                clients_list = [
                    {
                        "client": client,
                        "bls": data["bls"],
                        "poids_total": data["poids_total"],
                        "volume_total": data["volume_total"],
                        "representants": data["representants"]
                    }
                    for client, data in clients_data.items()
                ]
                
                # Trier par poids décroissant
                clients_list.sort(key=lambda x: x["poids_total"], reverse=True)
                
                # =====================================================
                # DÉTERMINER LE NUMÉRO D'ESTAFETTE DE DÉPART
                # =====================================================
                estafettes_existantes = self.df_base[
                    (self.df_base["Zone"] == zone) & 
                    (self.df_base["Code Véhicule"] != CAMION_CODE)
                ].copy()
                
                numeros_existants = []
                if not estafettes_existantes.empty:
                    for _, row in estafettes_existantes.iterrows():
                        camion_nom = row.get("Camion N°", "")
                        if "E" in str(camion_nom):
                            try:
                                num = int(str(camion_nom).replace("E", "").split("-")[0])
                                numeros_existants.append(num)
                            except:
                                pass
                    numeros_existants = sorted(set(numeros_existants))
                
                # =====================================================
                # RÈGLE SPÉCIALE POUR ZONE 7
                # =====================================================
                if zone == "Zone 7":
                    if numeros_existants:
                        num_estafette = numeros_existants[0]
                    else:
                        all_nums = []
                        for _, r in self.df_base.iterrows():
                            camion_nom = r.get("Camion N°", "")
                            if "E" in str(camion_nom):
                                try:
                                    n = int(str(camion_nom).replace("E", "").split("-")[0])
                                    all_nums.append(n)
                                except:
                                    pass
                        num_estafette = max(all_nums + [0]) + 1
                    
                    voyages = []
                    voyage_actuel = {
                        "poids": 0.0, "volume": 0.0, "bls": [],
                        "clients": set(), "representants": set()
                    }
                    
                    for client_data in clients_list:
                        client = client_data["client"]
                        
                        # Le client ENTIER rentre dans le voyage actuel
                        if (voyage_actuel["poids"] + client_data["poids_total"] <= CAPACITE_POIDS_ESTAFETTE and 
                            voyage_actuel["volume"] + client_data["volume_total"] <= CAPACITE_VOLUME_ESTAFETTE):
                            voyage_actuel["poids"] += client_data["poids_total"]
                            voyage_actuel["volume"] += client_data["volume_total"]
                            voyage_actuel["clients"].add(client)
                            voyage_actuel["representants"].update(client_data["representants"])
                            for bl_info in client_data["bls"]:
                                voyage_actuel["bls"].append(bl_info["bl"])
                        else:
                            # Sauvegarder le voyage actuel
                            if voyage_actuel["bls"]:
                                voyages.append(voyage_actuel)
                            
                            # Si on a atteint MAX_VOYAGES, forcer dans le dernier
                            if len(voyages) >= MAX_VOYAGES_ZONE7:
                                dernier = voyages[-1]
                                dernier["poids"] += client_data["poids_total"]
                                dernier["volume"] += client_data["volume_total"]
                                dernier["clients"].add(client)
                                dernier["representants"].update(client_data["representants"])
                                for bl_info in client_data["bls"]:
                                    dernier["bls"].append(bl_info["bl"])
                                voyage_actuel = {"poids": 0, "volume": 0, "bls": [], "clients": set(), "representants": set()}
                            else:
                                # Nouveau voyage avec le client ENTIER
                                voyage_actuel = {
                                    "poids": client_data["poids_total"],
                                    "volume": client_data["volume_total"],
                                    "bls": [bl_info["bl"] for bl_info in client_data["bls"]],
                                    "clients": {client},
                                    "representants": set(client_data["representants"])
                                }
                    
                    if voyage_actuel["bls"]:
                        voyages.append(voyage_actuel)
                    
                    for i, voyage in enumerate(voyages, 1):
                        if i > MAX_VOYAGES_ZONE7:
                            break
                        
                        clients_str = ", ".join(sorted(list(voyage["clients"])))
                        representants_str = ", ".join(sorted(list(voyage["representants"])))
                        taux_poids = (voyage["poids"] / CAPACITE_POIDS_ESTAFETTE) * 100
                        taux_volume = (voyage["volume"] / CAPACITE_VOLUME_ESTAFETTE) * 100
                        
                        resultats_optimises.append({
                            "Zone": zone,
                            "Estafette N°": num_estafette,
                            "Poids total": voyage["poids"],
                            "Volume total": voyage["volume"],
                            "Client(s) inclus": clients_str,
                            "Représentant(s) inclus": representants_str,
                            "BL inclus": ";".join(voyage["bls"]),
                            "Taux d'occupation (%)": max(taux_poids, taux_volume),
                            "Location_camion": False,
                            "Location_proposee": False,
                            "Code Véhicule": "ESTAFETTE",
                            "Camion N°": f"E{num_estafette}-Voyage {i}"
                        })
                
                else:
                    # =====================================================
                    # ALGORITHME STANDARD AVEC REGROUPEMENT STRICT PAR CLIENT
                    # =====================================================
                    estafettes_zone = []
                    
                    for client_data in clients_list:
                        client = client_data["client"]
                        poids_client = client_data["poids_total"]
                        volume_client = client_data["volume_total"]
                        
                        # ✅ CAS 1 : Le client ENTIER rentre dans une estafette existante
                        placed_entier = False
                        for e in estafettes_zone:
                            if (e["poids"] + poids_client <= CAPACITE_POIDS_ESTAFETTE and 
                                e["volume"] + volume_client <= CAPACITE_VOLUME_ESTAFETTE):
                                e["poids"] += poids_client
                                e["volume"] += volume_client
                                for bl_info in client_data["bls"]:
                                    e["bls"].append(bl_info["bl"])
                                e["clients"].add(client)
                                e["representants"].update(client_data["representants"])
                                placed_entier = True
                                break
                        
                        if placed_entier:
                            continue
                        
                        # ✅ CAS 2 : Le client dépasse la capacité d'UNE estafette → répartition inévitable
                        if poids_client > CAPACITE_POIDS_ESTAFETTE or volume_client > CAPACITE_VOLUME_ESTAFETTE:
                            bls_restants = client_data["bls"].copy()
                            while bls_restants:
                                placed_bl = False
                                for e in estafettes_zone:
                                    if (e["poids"] + bls_restants[0]["poids"] <= CAPACITE_POIDS_ESTAFETTE and 
                                        e["volume"] + bls_restants[0]["volume"] <= CAPACITE_VOLUME_ESTAFETTE):
                                        bl_info = bls_restants.pop(0)
                                        e["poids"] += bl_info["poids"]
                                        e["volume"] += bl_info["volume"]
                                        e["bls"].append(bl_info["bl"])
                                        e["clients"].add(client)
                                        e["representants"].update(client_data["representants"])
                                        placed_bl = True
                                        break
                                
                                if not placed_bl:
                                    if numeros_existants:
                                        num = numeros_existants.pop(0)
                                    else:
                                        num = (max([e["num_global"] for e in estafettes_zone], default=0) + 1
                                            if estafettes_zone else 1)
                                    
                                    bl_info = bls_restants.pop(0)
                                    estafettes_zone.append({
                                        "poids": bl_info["poids"],
                                        "volume": bl_info["volume"],
                                        "bls": [bl_info["bl"]],
                                        "clients": {client},
                                        "representants": set(client_data["representants"]),
                                        "num_global": num
                                    })
                        
                        # ✅ CAS 3 : Le client pourrait tenir dans une estafette VIDE → créer une nouvelle
                        else:
                            if numeros_existants:
                                num = numeros_existants.pop(0)
                            else:
                                num = (max([e["num_global"] for e in estafettes_zone], default=0) + 1
                                    if estafettes_zone else 1)
                            
                            estafettes_zone.append({
                                "poids": poids_client,
                                "volume": volume_client,
                                "bls": [bl_info["bl"] for bl_info in client_data["bls"]],
                                "clients": {client},
                                "representants": set(client_data["representants"]),
                                "num_global": num
                            })
                    
                    # Formater les résultats
                    for e in estafettes_zone:
                        clients_str = ", ".join(sorted(list(e["clients"])))
                        representants_str = ", ".join(sorted(list(e["representants"])))
                        taux_poids = (e["poids"] / CAPACITE_POIDS_ESTAFETTE) * 100
                        taux_volume = (e["volume"] / CAPACITE_VOLUME_ESTAFETTE) * 100
                        
                        resultats_optimises.append({
                            "Zone": zone,
                            "Estafette N°": e["num_global"],
                            "Poids total": e["poids"],
                            "Volume total": e["volume"],
                            "Client(s) inclus": clients_str,
                            "Représentant(s) inclus": representants_str,
                            "BL inclus": ";".join(e["bls"]),
                            "Taux d'occupation (%)": max(taux_poids, taux_volume),
                            "Location_camion": False,
                            "Location_proposee": False,
                            "Code Véhicule": "ESTAFETTE",
                            "Camion N°": f"E{e['num_global']}"
                        })
            
            if resultats_optimises:
                return pd.DataFrame(resultats_optimises)
            else:
                return pd.DataFrame()
                
        except Exception as e:
            print(f"❌ Erreur lors de la réoptimisation: {e}")
            return pd.DataFrame()

    def get_df_result(self):
        """Retourne le DataFrame optimisé final."""
        df_result = self.df_base.copy()
        
        # Renommer les colonnes si nécessaire
        rename_mapping = {
            "Poids total": "Poids total chargé",
            "Volume total": "Volume total chargé", 
            "Représentant": "Représentant(s) inclus"
        }
        
        # Appliquer seulement les renommages qui existent
        rename_mapping = {k: v for k, v in rename_mapping.items() if k in df_result.columns}
        if rename_mapping:
            df_result.rename(columns=rename_mapping, inplace=True)
        
        # S'assurer que "Véhicule N°" existe
        if "Camion N°" in df_result.columns and "Véhicule N°" not in df_result.columns:
            df_result["Véhicule N°"] = df_result["Camion N°"]
        
        # Préparer le tri
        df_result['Code_Tri'] = df_result['Code Véhicule'].apply(lambda x: 0 if x == CAMION_CODE else 1)
        
        # Définir l'ordre de tri par défaut
        sort_columns = ["Code_Tri", "Zone"]
        sort_ascending = [True, True]
        
        # Ajouter les colonnes de tri si elles existent
        if "Estafette N°" in df_result.columns:
            sort_columns.insert(1, "Estafette N°")
            sort_ascending.insert(1, True)
        
        if "Véhicule N°" in df_result.columns:
            sort_columns.append("Véhicule N°")
            sort_ascending.append(True)
        
        # Appliquer le tri
        df_result = df_result.sort_values(by=sort_columns, ascending=sort_ascending)
        
        # Nettoyer les colonnes temporaires
        df_result = df_result.drop(columns=['Code_Tri'], errors='ignore')
        
        # Ordre d'affichage final (ajouter Type_Camion si existe)
        final_columns = [
            "Zone", "Véhicule N°", "Poids total chargé", "Volume total chargé",
            "Client(s) inclus", "Représentant(s) inclus", "BL inclus", "Taux d'occupation (%)",
            "Location_camion", "Location_proposee", "Code Véhicule", "Type_Camion"
        ]
        
        # Filtrer seulement les colonnes qui existent
        available_columns = [col for col in final_columns if col in df_result.columns]
        return df_result[available_columns]
    # =====================================================
    # NOUVELLES MÉTHODES POUR LE COMPTAGE CORRECT
    # =====================================================
    def get_estafette_unique_count(self):
        """
        Retourne le nombre unique d'estafettes en tenant compte que Zone 7
        peut avoir plusieurs voyages pour une même estafette.
        """
        if self.df_base.empty:
            return 0
        
        # Filtrer les camions loués
        df_estafettes = self.df_base[self.df_base["Code Véhicule"] != CAMION_CODE].copy()
        
        if df_estafettes.empty:
            return 0
        
        # Pour Zone 7, compter les estafettes uniques (sans les numéros de voyage)
        df_zone7 = df_estafettes[df_estafettes["Zone"] == "Zone 7"]
        df_autres = df_estafettes[df_estafettes["Zone"] != "Zone 7"]
        
        # Compter les estafettes uniques pour Zone 7
        zone7_estafettes = 0
        if not df_zone7.empty:
            # ✅ Vérifier les deux noms possibles
            if "Camion N°" in df_zone7.columns:
                col_vehicule = "Camion N°"
            elif "Véhicule N°" in df_zone7.columns:
                col_vehicule = "Véhicule N°"
            else:
                # Si aucune colonne n'existe, utiliser l'index comme fallback
                return len(df_autres)
            
            zone7_estafettes = df_zone7[col_vehicule].apply(
                lambda x: str(x).split("-")[0] if "-Voyage" in str(x) else str(x)
            ).nunique()
        
        # Pour les autres zones, compter chaque ligne
        autres_estafettes = len(df_autres)
        
        return zone7_estafettes + autres_estafettes

    def get_camions_count(self):
        """Retourne le nombre de camions loués."""
        return len(self.df_base[self.df_base["Code Véhicule"] == CAMION_CODE])
# =====================================================
# CLASSE DE GESTION DES TRANSFERTS DE BL
# =====================================================
class TruckTransferManager:
    def __init__(self, df_voyages, df_livraisons):
        self.df_voyages = df_voyages.copy()
        self.df_livraisons = df_livraisons.copy()
        self.MAX_POIDS_ESTAFETTE = CAPACITE_POIDS_ESTAFETTE
        self.MAX_VOLUME_ESTAFETTE = CAPACITE_VOLUME_ESTAFETTE
    
    def _get_capacites_vehicule(self, vehicule, df_voyages):
        """Retourne les capacités max selon le type de véhicule."""
        # Rechercher le véhicule dans le DataFrame
        if "Véhicule N°" in df_voyages.columns:
            veh_col = "Véhicule N°"
        elif "Camion N°" in df_voyages.columns:
            veh_col = "Camion N°"
        else:
            return CAPACITE_POIDS_ESTAFETTE, CAPACITE_VOLUME_ESTAFETTE
        
        vehicule_data = df_voyages[df_voyages[veh_col] == vehicule]
        
        if vehicule_data.empty:
            return CAPACITE_POIDS_ESTAFETTE, CAPACITE_VOLUME_ESTAFETTE
        
        row = vehicule_data.iloc[0]
        
        # Déterminer si c'est un camion loué
        is_camion = row.get("Code Véhicule", "") == CAMION_CODE or str(vehicule).upper().startswith("C")
        
        if is_camion:
            # Vérifier le type de camion
            truck_type = row.get("Type_Camion", "5 tonnes")
            # Utiliser les fonctions pour obtenir les capacités
            return get_capacite_poids_camion(truck_type), get_capacite_volume_camion(truck_type)
        else:
            return CAPACITE_POIDS_ESTAFETTE, CAPACITE_VOLUME_ESTAFETTE

    def transferer_bls(self, zone, source, cible, bls_a_transferer):
        """Transfère des BLs d'un véhicule source à un véhicule cible."""
        try:
            # Vérifier que les BLs existent dans la source
            df_source = self.df_voyages[
                (self.df_voyages["Zone"] == zone) & 
                (self.df_voyages["Véhicule N°"] == source)
            ]
            
            if df_source.empty:
                return False, f"❌ Véhicule source {source} non trouvé dans la zone {zone}", self.df_voyages
            
            bls_source = df_source["BL inclus"].iloc[0].split(';')
            bls_existants = [bl for bl in bls_a_transferer if bl in bls_source]
            
            if not bls_existants:
                return False, f"❌ Aucun des BLs sélectionnés n'est présent dans le véhicule source {source}", self.df_voyages
            
            # Calculer le poids et volume des BLs à transférer depuis les données originales
            df_bls_transfert = self.df_livraisons[
                self.df_livraisons["No livraison"].isin(bls_existants)
            ]
            poids_transfert = df_bls_transfert["Poids total"].sum()
            volume_transfert = df_bls_transfert["Volume total"].sum()
            
            # Vérifier la capacité du véhicule cible avec les capacités dynamiques
            df_cible = self.df_voyages[
                (self.df_voyages["Zone"] == zone) & 
                (self.df_voyages["Véhicule N°"] == cible)
            ]
            
            if df_cible.empty:
                return False, f"❌ Véhicule cible {cible} non trouvé dans la zone {zone}", self.df_voyages
            
            # Obtenir les capacités max du véhicule cible
            max_poids_cible, max_volume_cible = self._get_capacites_vehicule(cible, self.df_voyages)
            
            poids_cible_actuel = df_cible["Poids total chargé"].iloc[0]
            volume_cible_actuel = df_cible["Volume total chargé"].iloc[0]
            
            if (poids_cible_actuel + poids_transfert > max_poids_cible or 
                volume_cible_actuel + volume_transfert > max_volume_cible):
                return False, f"❌ Le transfert dépasse les capacités du véhicule cible {cible} (Max: {max_poids_cible}kg, {max_volume_cible}m³)", self.df_voyages
            
            # Appliquer le transfert
            for idx, row in self.df_voyages.iterrows():
                if row["Zone"] == zone and row["Véhicule N°"] == source:
                    bls_restants = [bl for bl in row["BL inclus"].split(';') if bl not in bls_existants]
                    self.df_voyages.at[idx, "BL inclus"] = ';'.join(bls_restants)
                    self.df_voyages.at[idx, "Poids total chargé"] -= poids_transfert
                    self.df_voyages.at[idx, "Volume total chargé"] -= volume_transfert
                    
                    # Recalculer le taux d'occupation pour la source
                    max_poids_source, max_volume_source = self._get_capacites_vehicule(source, self.df_voyages)
                    new_poids_source = max(0, self.df_voyages.at[idx, "Poids total chargé"])
                    new_volume_source = max(0, self.df_voyages.at[idx, "Volume total chargé"])
                    taux_source = max((new_poids_source / max_poids_source) * 100, (new_volume_source / max_volume_source) * 100)
                    self.df_voyages.at[idx, "Taux d'occupation (%)"] = taux_source
                
                elif row["Zone"] == zone and row["Véhicule N°"] == cible:
                    bls_actuels = row["BL inclus"].split(';')
                    # Combiner et éliminer les doublons
                    bls_nouveaux = list(dict.fromkeys(bls_actuels + bls_existants))
                    self.df_voyages.at[idx, "BL inclus"] = ';'.join(bls_nouveaux)
                    self.df_voyages.at[idx, "Poids total chargé"] += poids_transfert
                    self.df_voyages.at[idx, "Volume total chargé"] += volume_transfert
                    
                    # Recalculer le taux d'occupation pour la cible
                    max_poids_cible, max_volume_cible = self._get_capacites_vehicule(cible, self.df_voyages)
                    new_poids_cible = self.df_voyages.at[idx, "Poids total chargé"]
                    new_volume_cible = self.df_voyages.at[idx, "Volume total chargé"]
                    taux_cible = max((new_poids_cible / max_poids_cible) * 100, (new_volume_cible / max_volume_cible) * 100)
                    self.df_voyages.at[idx, "Taux d'occupation (%)"] = taux_cible
            
            message = f"✅ Transfert réussi : {len(bls_existants)} BL(s) déplacé(s) de {source} vers {cible}"
            return True, message, self.df_voyages
            
        except Exception as e:
            return False, f"❌ Erreur lors du transfert : {str(e)}", self.df_voyages

    def get_voyages_actuels(self):
        return self.df_voyages

    def add_manual_object(self, df_voyages, vehicle, zone, name, weight, volume):
        """
        Ajoute un objet manuel (objet virtuel) dans le véhicule sélectionné.
        """
        try:
            # Validation inputs
            weight = float(weight)
            volume = float(volume)
            if weight < 0 or volume < 0:
                return False, "Poids et volume doivent être >= 0", df_voyages

            df = df_voyages.copy()

            # Rechercher la ligne du véhicule dans df
            if "Véhicule N°" in df.columns:
                veh_col = "Véhicule N°"
            elif "Camion N°" in df.columns:
                veh_col = "Camion N°"
            else:
                return False, "Structure du DataFrame inattendue.", df

            mask = (df[veh_col] == vehicle) & (df["Zone"] == zone)
            if not mask.any():
                return False, f"Véhicule {vehicle} non trouvé dans la zone {zone}.", df

            idx = df[mask].index[0]
            row = df.loc[idx].copy()

            # Obtenir les capacités max selon le type de véhicule
            max_poids, max_volume = self._get_capacites_vehicule(vehicle, df)

            current_poids = float(row.get("Poids total chargé", row.get("Poids total", 0)) or 0)
            current_volume = float(row.get("Volume total chargé", row.get("Volume total", 0)) or 0)

            new_poids = current_poids + weight
            new_volume = current_volume + volume

            # Vérifier les capacités
            if new_poids > max_poids or new_volume > max_volume:
                # Récupérer le type de véhicule pour le message d'erreur
                is_camion = row.get("Code Véhicule", "") == CAMION_CODE or str(vehicle).upper().startswith("C")
                vehicle_type = "camion" if is_camion else "estafette"
                truck_type = row.get("Type_Camion", "") if is_camion else ""
                type_info = f" ({truck_type})" if truck_type else ""
                
                return False, f"❌ Capacité dépassée pour {vehicle_type}{type_info} {vehicle} : {new_poids:.1f}kg/{max_poids}kg, {new_volume:.3f}m³/{max_volume}m³", df

            # Générer code unique pour l'objet
            obj_code = f"OBJ-{name}"

            # Mettre à jour BL inclus
            bls_current = str(row.get("BL inclus", "")).strip()
            if bls_current == "nan" or bls_current == "":
                new_bls = obj_code
            else:
                new_bls = bls_current + ";" + obj_code

            # Appliquer modifications
            df.at[idx, "BL inclus"] = new_bls
            if "Poids total chargé" in df.columns:
                df.at[idx, "Poids total chargé"] = new_poids
            else:
                df.at[idx, "Poids total"] = new_poids

            if "Volume total chargé" in df.columns:
                df.at[idx, "Volume total chargé"] = new_volume
            else:
                df.at[idx, "Volume total"] = new_volume

            # Recalculer taux d'occupation avec les bonnes capacités
            taux = max((new_poids / max_poids) * 100, (new_volume / max_volume) * 100)
            df.at[idx, "Taux d'occupation (%)"] = taux

            return True, f"✅ Objet '{name}' ajouté à {vehicle} en zone {zone}", df

        except Exception as e:
            return False, f"❌ Erreur lors de l'ajout de l'objet : {str(e)}", df_voyages

# =====================================================
# CLASSE DE GESTION DES RAPPORTS AVANCÉS
# =====================================================
class AdvancedReportGenerator:
    def __init__(self, df_voyages, df_livraisons_original):
        self.df_voyages = df_voyages.copy()
        self.df_livraisons_original = df_livraisons_original.copy()
    
    def generer_rapport_analytique(self):
        """Génère un rapport analytique complet."""
        try:
            analyses = []
            
            # 1. Analyse par type de véhicule
            estafettes = self.df_voyages[self.df_voyages["Code Véhicule"] == "ESTAFETTE"]
            camions = self.df_voyages[self.df_voyages["Code Véhicule"] == CAMION_CODE]
            
            # ✅ CORRECTION : Compter les estafettes uniques (avec règle Zone 7)
            df_estafettes = self.df_voyages[self.df_voyages["Code Véhicule"] != CAMION_CODE].copy()
            
            # Pour Zone 7, compter les estafettes uniques (sans les numéros de voyage)
            df_zone7 = df_estafettes[df_estafettes["Zone"] == "Zone 7"]
            df_autres = df_estafettes[df_estafettes["Zone"] != "Zone 7"]
            
            zone7_estafettes = 0
            if not df_zone7.empty:
                # Vérifier les deux noms possibles
                if "Véhicule N°" in df_zone7.columns:
                    col_vehicule = "Véhicule N°"
                elif "Camion N°" in df_zone7.columns:
                    col_vehicule = "Camion N°"
                else:
                    col_vehicule = None
                
                if col_vehicule:
                    zone7_estafettes = df_zone7[col_vehicule].apply(
                        lambda x: str(x).split("-")[0] if "-Voyage" in str(x) else str(x)
                    ).nunique()
            
            nb_estafettes_uniques = zone7_estafettes + len(df_autres)
            nb_camions = len(camions)
            
            analyses.append("📊 ANALYSE PAR TYPE DE VÉHICULE")
            analyses.append(f"• Nombre total d'estafettes : {nb_estafettes_uniques}")
            analyses.append(f"• Nombre total de camions : {nb_camions}")
            analyses.append(f"• Poids total transporté par estafettes : {estafettes['Poids total chargé'].sum():.1f} kg")
            analyses.append(f"• Volume total transporté par estafettes : {estafettes['Volume total chargé'].sum():.3f} m³")
            analyses.append(f"• Poids total transporté par camions : {camions['Poids total chargé'].sum():.1f} kg")
            analyses.append(f"• Volume total transporté par camions : {camions['Volume total chargé'].sum():.3f} m³")
            
            # 2. Analyse par zone
            analyses.append("\n🌍 ANALYSE PAR ZONE GÉOGRAPHIQUE")
            for zone in self.df_voyages["Zone"].unique():
                df_zone = self.df_voyages[self.df_voyages["Zone"] == zone]
                analyses.append(f"• {zone} : {len(df_zone)} véhicules, {df_zone['Poids total chargé'].sum():.1f} kg, {df_zone['Volume total chargé'].sum():.3f} m³")
            
            # 3. Analyse d'efficacité
            analyses.append("\n⚡ ANALYSE D'EFFICACITÉ")
            taux_moyen_estafettes = estafettes["Taux d'occupation (%)"].mean() if len(estafettes) > 0 else 0
            taux_moyen_camions = camions["Taux d'occupation (%)"].mean() if len(camions) > 0 else 0
            
            analyses.append(f"• Taux d'occupation moyen des estafettes : {taux_moyen_estafettes:.1f}%")
            analyses.append(f"• Taux d'occupation moyen des camions : {taux_moyen_camions:.1f}%")
            
            # Véhicules sous-utilisés (< 60%)
            vehicules_sous_utilises = self.df_voyages[self.df_voyages["Taux d'occupation (%)"] < 60]
            if len(vehicules_sous_utilises) > 0:
                analyses.append(f"• Véhicules sous-utilisés (< 60%) : {len(vehicules_sous_utilises)}")
                for idx, row in vehicules_sous_utilises.iterrows():
                    analyses.append(f"  - {row['Véhicule N°']} (Zone {row['Zone']}) : {row['Taux d\'occupation (%)']:.1f}%")
            
            # Véhicules sur-utilisés (> 95%)
            vehicules_sur_utilises = self.df_voyages[self.df_voyages["Taux d'occupation (%)"] > 95]
            if len(vehicules_sur_utilises) > 0:
                analyses.append(f"• Véhicules très chargés (> 95%) : {len(vehicules_sur_utilises)}")
                for idx, row in vehicules_sur_utilises.iterrows():
                    analyses.append(f"  - {row['Véhicule N°']} (Zone {row['Zone']}) : {row['Taux d\'occupation (%)']:.1f}%")
            
            # 4. Analyse économique
            analyses.append("\n💰 ANALYSE ÉCONOMIQUE")
            analyses.append(f"• Coût estimé des estafettes : {nb_estafettes_uniques} x [coût unitaire]")
            analyses.append(f"• Coût estimé des camions : {nb_camions} x [coût unitaire camion]")
            
            # 5. Recommandations
            analyses.append("\n🎯 RECOMMANDATIONS")
            if len(vehicules_sous_utilises) > len(vehicules_sur_utilises):
                analyses.append("• Optimisation possible : regrouper certains voyages sous-utilisés")
            
            if len(camions) > 0 and camions["Taux d'occupation (%)"].mean() < 70:
                analyses.append("• Attention : les camions sont sous-utilisés, envisager plus d'estafettes")
            
            if len(vehicules_sur_utilises) > 0:
                analyses.append("• Vigilance : certains véhicules sont à pleine capacité")
            
            return "\n".join(analyses)
            
        except Exception as e:
            return f"❌ Erreur lors de la génération du rapport analytique : {str(e)}"

    def generer_rapport_client(self, client):
        """Génère un rapport spécifique pour un client."""
        try:
            # Trouver tous les BLs du client dans les données originales
            bls_client = self.df_livraisons_original[
                self.df_livraisons_original["Client de l'estafette"] == client
            ]["No livraison"].unique()
            
            # Trouver les véhicules qui transportent ces BLs
            vehicules_client = []
            for idx, row in self.df_voyages.iterrows():
                bls_vehicule = str(row["BL inclus"]).split(';')
                if any(str(bl) in bls_vehicule for bl in bls_client):
                    vehicules_client.append({
                        'Véhicule': row['Véhicule N°'],
                        'Zone': row['Zone'],
                        'Type': 'Camion' if row['Code Véhicule'] == CAMION_CODE else 'Estafette',
                        'Poids': row['Poids total chargé'],
                        'Volume': row['Volume total chargé'],
                        'Taux Occupation': row['Taux d\'occupation (%)'],
                        'Date Livraison Estimée': 'À planifier'  # Peut être enrichi avec des données de planning
                    })
            
            if not vehicules_client:
                return f"Aucune livraison trouvée pour le client {client}"
            
            # Générer le rapport
            rapport = [f"📦 RAPPORT LIVRAISON - CLIENT {client}"]
            rapport.append(f"Nombre de véhicules concernés : {len(vehicules_client)}")
            rapport.append("\nDétails des véhicules :")
            
            for veh in vehicules_client:
                rapport.append(
                    f"• {veh['Type']} {veh['Véhicule']} ({veh['Zone']}) : "
                    f"{veh['Poids']:.1f}kg, {veh['Volume']:.3f}m³, "
                    f"Taux {veh['Taux Occupation']:.1f}%"
                )
            
            # Calcul des totaux
            total_poids = sum(veh['Poids'] for veh in vehicules_client)
            total_volume = sum(veh['Volume'] for veh in vehicules_client)
            
            rapport.append(f"\n📊 TOTAUX CLIENT {client}:")
            rapport.append(f"• Poids total : {total_poids:.1f} kg")
            rapport.append(f"• Volume total : {total_volume:.3f} m³")
            rapport.append(f"• Nombre de véhicules : {len(vehicules_client)}")
            
            return "\n".join(rapport)
            
        except Exception as e:
            return f"❌ Erreur lors de la génération du rapport client : {str(e)}"

# =====================================================
# FONCTIONS UTILITAIRES GLOBALES
# =====================================================
def calculer_couts_estimation(df_voyages, cout_estafette=150, cout_camion=800):
    """Estime les coûts de transport basés sur les véhicules utilisés."""
    try:
        nb_estafettes = len(df_voyages[df_voyages["Code Véhicule"] == "ESTAFETTE"])
        nb_camions = len(df_voyages[df_voyages["Code Véhicule"] == CAMION_CODE])
        
        cout_total = (nb_estafettes * cout_estafette) + (nb_camions * cout_camion)
        
        return {
            'estafettes': nb_estafettes,
            'camions': nb_camions,
            'cout_estafette_unitaire': cout_estafette,
            'cout_camion_unitaire': cout_camion,
            'cout_total': cout_total,
            'cout_estimation': f"💰 Estimation des coûts : {nb_estafettes} estafettes × {cout_estafette} TND + {nb_camions} camions × {cout_camion} TND = {cout_total} TND"
        }
    except Exception as e:
        return {'erreur': f"❌ Erreur dans le calcul des coûts : {str(e)}"}

def exporter_planning_excel(df_voyages, file_path, donnees_supplementaires=None, df_livraisons_original=None):
    """Exporte le planning complet vers Excel avec prise en compte des types de camions."""
    try:
        import openpyxl
        from openpyxl.styles import Alignment
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            # =====================================================
            # ORDRE EXACT DES COLONNES DEMANDÉ AVEC VILLE
            # =====================================================
            colonnes_demandees = [
                "Code voyage", "Zone", "Ville", "Véhicule N°", "Chauffeur", 
                "BL inclus", "Client(s) inclus", "Poids total chargé", 
                "Volume total chargé"
            ]
            
            # =====================================================
            # CORRECTION : UTILISATION DES COLONNES EXISTANTES
            # =====================================================
            
            # Faire une copie pour éviter les modifications sur l'original
            df_voyages_working = df_voyages.copy()
            
            # 1. Vérifier et mapper "Chauffeur" vers les colonnes existantes
            if "Chauffeur" not in df_voyages_working.columns:
                # Priorité 1 : Utiliser "Chauffeur attribué"
                if "Chauffeur attribué" in df_voyages_working.columns:
                    df_voyages_working["Chauffeur"] = df_voyages_working["Chauffeur attribué"]
                # Priorité 2 : Utiliser "Nom_chauffeur" 
                elif "Nom_chauffeur" in df_voyages_working.columns:
                    df_voyages_working["Chauffeur"] = df_voyages_working["Nom_chauffeur"]
                # Priorité 3 : Utiliser "Matricule chauffeur" avec format
                elif "Matricule chauffeur" in df_voyages_working.columns:
                    df_voyages_working["Chauffeur"] = df_voyages_working["Matricule chauffeur"].apply(
                        lambda x: f"Chauffeur {x}" if pd.notna(x) and x != "" else "À attribuer"
                    )
                # Fallback : Colonne vide
                else:
                    df_voyages_working["Chauffeur"] = "À attribuer"
            
            # 2. AJOUT DE LA COLONNE "TYPE VÉHICULE" - NOUVELLE FONCTIONNALITÉ
            if 'Code Véhicule' in df_voyages_working.columns:
                def get_info_capacite(row):
                    if row['Code Véhicule'] == CAMION_CODE:
                        truck_type = row.get('Type_Camion', '5 tonnes')
                        poids_max = get_capacite_poids_camion(truck_type)
                        volume_max = get_capacite_volume_camion(truck_type)
                        return f"Camion {truck_type} ({poids_max} kg, {volume_max} m³)"
                    else:
                        return f"Estafette ({CAPACITE_POIDS_ESTAFETTE} kg, {CAPACITE_VOLUME_ESTAFETTE} m³)"
                
                df_voyages_working['Type_Véhicule'] = df_voyages_working.apply(get_info_capacite, axis=1)
                # Ajouter cette colonne à la liste des colonnes demandées
                colonnes_demandees.insert(4, "Type_Véhicule")  # Après "Véhicule N°"
            
            # 3. AJOUT DE LA COLONNE "VILLE" - NOUVELLE FONCTIONNALITÉ
            if "Ville" not in df_voyages_working.columns and df_livraisons_original is not None:
                print("🔄 Ajout de la colonne Ville depuis les données originales...")
                
                # Créer un mapping BL -> Ville depuis les données originales
                mapping_ville = {}
                
                # Parcourir les données originales pour créer le mapping
                for idx, row in df_livraisons_original.iterrows():
                    bl = str(row.get("No livraison", ""))
                    ville = str(row.get("Ville", ""))
                    if bl and bl != "nan" and ville and ville != "nan":
                        mapping_ville[bl] = ville
                
                # Fonction pour extraire les villes à partir des BLs d'un véhicule
                def get_villes_from_bls(bls_inclus):
                    if pd.isna(bls_inclus) or bls_inclus == "":
                        return ""
                    
                    bls_list = str(bls_inclus).split(';')
                    villes_trouvees = set()
                    
                    for bl in bls_list:
                        bl_clean = bl.strip()
                        if bl_clean in mapping_ville:
                            villes_trouvees.add(mapping_ville[bl_clean])
                        # Ignorer les objets manuels (OBJ-)
                        elif not bl_clean.startswith('OBJ-'):
                            # Chercher le BL dans les données originales
                            for original_bl, original_ville in mapping_ville.items():
                                if bl_clean == original_bl:
                                    villes_trouvees.add(original_ville)
                                    break
                    
                    return ", ".join(sorted(villes_trouvees)) if villes_trouvees else "Ville inconnue"
                
                # Appliquer la fonction pour créer la colonne Ville
                df_voyages_working["Ville"] = df_voyages_working["BL inclus"].apply(get_villes_from_bls)
                print("✅ Colonne 'Ville' ajoutée avec succès")
            
            # 4. FORMATER LES COLONNES AVEC RETOURS À LIGNE
            colonnes_retours_ligne = ['BL inclus', 'Client(s) inclus', 'Représentant(s) inclus']
            for col in colonnes_retours_ligne:
                if col in df_voyages_working.columns:
                    df_voyages_working[col] = df_voyages_working[col].apply(
                        lambda x: '\n'.join([elem.strip() for elem in str(x).replace(';', ',').split(',') if elem.strip()]) 
                        if pd.notna(x) else ""
                    )
            
            # 5. Filtrer seulement les colonnes qui existent
            colonnes_finales = [col for col in colonnes_demandees if col in df_voyages_working.columns]
            
            # 6. Vérifier qu'on a au moins les colonnes de base et que le DataFrame n'est pas vide
            if df_voyages_working.empty:
                # Créer une feuille vide avec les colonnes demandées pour éviter l'erreur
                df_voyages_ordered = pd.DataFrame(columns=colonnes_finales)
                print("⚠️ DataFrame vide - création d'une structure vide")
            else:
                colonnes_requises = ["Zone", "Véhicule N°", "BL inclus", "Client(s) inclus"]
                colonnes_manquantes = [col for col in colonnes_requises if col not in colonnes_finales]
                
                if colonnes_manquantes:
                    print(f"❌ Colonnes manquantes : {', '.join(colonnes_manquantes)}")
                    # Créer quand même l'export avec les colonnes disponibles
                    df_voyages_ordered = df_voyages_working[colonnes_finales].copy()
                else:
                    # Réorganiser le DataFrame avec l'ordre exact demandé
                    df_voyages_ordered = df_voyages_working[colonnes_finales].copy()
            
            # =====================================================
            # FORMATAGE DES VALEURS NUMÉRIQUES
            # =====================================================
            if "Poids total chargé" in df_voyages_ordered.columns and not df_voyages_ordered.empty:
                df_voyages_ordered["Poids total chargé"] = df_voyages_ordered["Poids total chargé"].round(3)
            
            if "Volume total chargé" in df_voyages_ordered.columns and not df_voyages_ordered.empty:
                df_voyages_ordered["Volume total chargé"] = df_voyages_ordered["Volume total chargé"].round(3)
            
            
            # =====================================================
            # FEUILLE PRINCIPALE - PLANNING LIVRAISONS
            # =====================================================

            # Copier le DataFrame pour ne pas modifier les données originales
            df_planning_excel = df_voyages_ordered.copy()

            # =====================================================
            # SUPPRIMER LA COLONNE VILLE
            # UNIQUEMENT POUR LA FEUILLE PLANNING LIVRAISONS
            # =====================================================
            df_planning_excel = df_planning_excel.drop(
                columns=["Ville"],
                errors="ignore"
            )

            # =====================================================
            # SIMPLIFIER LE TYPE DE VÉHICULE
            # UNIQUEMENT POUR LA FEUILLE PLANNING LIVRAISONS
            # =====================================================
            if "Type_Véhicule" in df_planning_excel.columns:
                df_planning_excel["Type_Véhicule"] = (
                    df_planning_excel["Type_Véhicule"]
                    .fillna("")
                    .astype(str)
                    .str.replace(r"\s*\([^)]*\)", "", regex=True)
                    .str.strip()
                )

            # =====================================================
            # EXPORTER LA FEUILLE PRINCIPALE
            # =====================================================
            if not df_planning_excel.empty:
                df_planning_excel.to_excel(
                    writer,
                    sheet_name="Planning Livraisons",
                    index=False
                )
            else:
                pd.DataFrame(
                    columns=[
                        col for col in colonnes_finales
                        if col != "Ville"
                    ]
                ).to_excel(
                    writer,
                    sheet_name="Planning Livraisons",
                    index=False
                )

            # =====================================================
            # APPLIQUER LE FORMATAGE DES RETOURS À LA LIGNE DANS EXCEL
            # =====================================================
            workbook = writer.book
            worksheet = writer.sheets["Planning Livraisons"]

            # Style avec retours à la ligne et centrage
            wrap_alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            # Appliquer le formatage aux colonnes avec retours à la ligne
            for col_idx, col_name in enumerate(df_planning_excel.columns, 1):
                if col_name in colonnes_retours_ligne:
                    for row_idx in range(2, len(df_planning_excel) + 2):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        cell.alignment = wrap_alignment

            # Ajuster la hauteur des lignes
            for row in range(2, len(df_planning_excel) + 2):
                worksheet.row_dimensions[row].height = 40

            # Ajuster automatiquement la largeur des colonnes
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter

                for cell in column:
                    try:
                        if cell.value:
                            lines = str(cell.value).split("\n")
                            max_line_length = max(len(line) for line in lines)
                            max_length = max(max_length, max_line_length)
                    except Exception:
                        pass

                adjusted_width = min(50, max_length + 2)
                worksheet.column_dimensions[column_letter].width = adjusted_width

            
            # =====================================================
            # FEUILLE DE SYNTHÈSE DÉTAILLÉE
            # =====================================================
            try:
                nb_estafettes = len(df_voyages_working[df_voyages_working["Code Véhicule"] == "ESTAFETTE"]) if "Code Véhicule" in df_voyages_working.columns else 0
                nb_camions = len(df_voyages_working[df_voyages_working["Code Véhicule"] == CAMION_CODE]) if "Code Véhicule" in df_voyages_working.columns else 0
                
                # Compter les types de camions
                camions_5t = 0
                camions_10t = 0
                if nb_camions > 0 and "Type_Camion" in df_voyages_working.columns:
                    camions_5t = len(df_voyages_working[(df_voyages_working["Code Véhicule"] == CAMION_CODE) & (df_voyages_working["Type_Camion"] == "5 tonnes")])
                    camions_10t = len(df_voyages_working[(df_voyages_working["Code Véhicule"] == CAMION_CODE) & (df_voyages_working["Type_Camion"] == "10 tonnes")])
                
                poids_total = df_voyages_working['Poids total chargé'].sum() if 'Poids total chargé' in df_voyages_working.columns else 0
                volume_total = df_voyages_working['Volume total chargé'].sum() if 'Volume total chargé' in df_voyages_working.columns else 0
                taux_moyen = df_voyages_working['Taux d\'occupation (%)'].mean() if 'Taux d\'occupation (%)' in df_voyages_working.columns else 0
                
                synthèse_data = {
                    'Métrique': [
                        'Total Véhicules', 
                        'Estafettes', 
                        'Camions (Total)',
                        'Camions 5 tonnes',
                        'Camions 10 tonnes',
                        'Poids Total', 
                        'Volume Total', 
                        'Taux Occupation Moyen'
                    ],
                    'Valeur': [
                        len(df_voyages_working),
                        nb_estafettes,
                        nb_camions,
                        camions_5t,
                        camions_10t,
                        f"{poids_total:.1f} kg",
                        f"{volume_total:.3f} m³",
                        f"{taux_moyen:.1f}%" if taux_moyen > 0 else "N/A"
                    ]
                }
                pd.DataFrame(synthèse_data).to_excel(writer, sheet_name='Synthèse', index=False)
            except Exception as e:
                print(f"⚠️ Erreur lors de la création de la synthèse : {e}")
                pd.DataFrame({'Métrique': ['Erreur'], 'Valeur': ['Données non disponibles']}).to_excel(writer, sheet_name='Synthèse', index=False)
            
            # =====================================================
            # FEUILLE STATS PAR ZONE DÉTAILLÉE
            # =====================================================
            try:
                if 'Zone' in df_voyages_working.columns and not df_voyages_working.empty:
                    # Ajouter une colonne Type pour les statistiques
                    df_stats = df_voyages_working.copy()
                    df_stats['Type_Véhicule_Simple'] = df_stats['Code Véhicule'].apply(
                        lambda x: 'CAMION' if x == CAMION_CODE else 'ESTAFETTE'
                    )
                    
                    # Statistiques par zone
                    stats_zone = df_stats.groupby('Zone').agg({
                        'Véhicule N°': 'count',
                        'Type_Véhicule_Simple': lambda x: (x == 'CAMION').sum(),
                        'Poids total chargé': ['sum', 'mean'],
                        'Volume total chargé': ['sum', 'mean'],
                        'Taux d\'occupation (%)': 'mean'
                    }).round(2)
                    
                    # Renommer les colonnes
                    stats_zone.columns = [
                        'Nombre Véhicules',
                        'Nombre Camions',
                        'Poids Total (kg)',
                        'Poids Moyen (kg)',
                        'Volume Total (m³)',
                        'Volume Moyen (m³)',
                        'Taux Occupation Moyen (%)'
                    ]
                    
                    # Ajouter le nombre d'estafettes
                    stats_zone['Nombre Estafettes'] = stats_zone['Nombre Véhicules'] - stats_zone['Nombre Camions']
                    
                    stats_zone.to_excel(writer, sheet_name='Stats par Zone')
                else:
                    pd.DataFrame(columns=['Zone', 'Nombre_Véhicules']).to_excel(writer, sheet_name='Stats par Zone', index=False)
            except Exception as e:
                print(f"⚠️ Erreur lors de la création des stats par zone : {e}")
                pd.DataFrame(columns=['Zone', 'Nombre_Véhicules']).to_excel(writer, sheet_name='Stats par Zone', index=False)
            
            # =====================================================
            # FEUILLE CAPACITÉS VÉHICULES
            # =====================================================
            try:
                if not df_voyages_working.empty:
                    capacites_data = []
                    
                    for idx, row in df_voyages_working.iterrows():
                        if row.get('Code Véhicule') == CAMION_CODE:
                            truck_type = row.get('Type_Camion', '5 tonnes')
                            poids_max = get_capacite_poids_camion(truck_type)
                            volume_max = get_capacite_volume_camion(truck_type)
                            type_veh = f"Camion {truck_type}"
                        else:
                            poids_max = CAPACITE_POIDS_ESTAFETTE
                            volume_max = CAPACITE_VOLUME_ESTAFETTE
                            type_veh = "Estafette"
                        
                        capacites_data.append({
                            'Véhicule': row.get('Véhicule N°', ''),
                            'Type': type_veh,
                            'Zone': row.get('Zone', ''),
                            'Poids Chargé (kg)': row.get('Poids total chargé', 0),
                            'Capacité Max Poids (kg)': poids_max,
                            'Utilisation Poids (%)': round((row.get('Poids total chargé', 0) / poids_max) * 100, 1),
                            'Volume Chargé (m³)': row.get('Volume total chargé', 0),
                            'Capacité Max Volume (m³)': volume_max,
                            'Utilisation Volume (%)': round((row.get('Volume total chargé', 0) / volume_max) * 100, 1)
                        })
                    
                    df_capacites = pd.DataFrame(capacites_data)
                    df_capacites.to_excel(writer, sheet_name='Capacités Véhicules', index=False)
            except Exception as e:
                print(f"⚠️ Erreur lors de la création des capacités véhicules : {e}")
            
            
            # =====================================================
            # DONNÉES SUPPLÉMENTAIRES
            # AVEC RETOURS À LA LIGNE POUR ARTICLE ET DÉSIGNATION
            # =====================================================
            if donnees_supplementaires:
                for nom_feuille, data in donnees_supplementaires.items():

                    nom_feuille = nom_feuille[:31]

                    if isinstance(data, pd.DataFrame) and not data.empty:

                        # Copier les données pour ne pas modifier l'original
                        data_export = data.copy()

                        # =========================================
                        # ARTICLE : un article par ligne dans la cellule
                        # =========================================
                        if "Article" in data_export.columns:
                            data_export["Article"] = (
                                data_export["Article"]
                                .fillna("")
                                .astype(str)
                                .apply(
                                    lambda x: "\n".join(
                                        article.strip()
                                        for article in x.split(",")
                                        if article.strip()
                                    )
                                )
                            )

                        # =========================================
                        # DÉSIGNATION : une désignation par ligne
                        # =========================================
                        for col in data_export.columns:
                            if "signation" in str(col).lower():

                                data_export[col] = (
                                    data_export[col]
                                    .fillna("")
                                    .astype(str)
                                    .apply(
                                        lambda x: "\n".join(
                                            designation.strip()
                                            for designation in (
                                                x.replace("<br />", "||")
                                                .replace("<br/>", "||")
                                                .replace("<br>", "||")
                                            ).split("||")
                                            if designation.strip()
                                        )
                                    )
                                )

                        # Exporter la feuille
                        data_export.to_excel(
                            writer,
                            sheet_name=nom_feuille,
                            index=False
                        )

                        # =========================================
                        # ACTIVER LE RETOUR À LA LIGNE DANS EXCEL
                        # =========================================
                        worksheet = writer.sheets[nom_feuille]

                        for col_idx, col_name in enumerate(data_export.columns, 1):

                            # Colonnes Article et Désignation
                            if (
                                col_name == "Article"
                                or "signation" in str(col_name).lower()
                            ):

                                # Largeur de la colonne
                                from openpyxl.utils import get_column_letter
                                lettre = get_column_letter(col_idx)
                                worksheet.column_dimensions[lettre].width = 35

                                # Activer le retour à la ligne
                                for row_idx in range(2, len(data_export) + 2):

                                    cellule = worksheet.cell(
                                        row=row_idx,
                                        column=col_idx
                                    )

                                    cellule.alignment = Alignment(
                                        horizontal="center",
                                        vertical="center",
                                        wrap_text=True
                                    )

                        # =========================================
                        # AJUSTER LA HAUTEUR DES LIGNES
                        # =========================================
                        for row_idx in range(2, len(data_export) + 2):

                            max_lignes = 1

                            for cellule in worksheet[row_idx]:
                                if cellule.value is not None:
                                    nb_lignes = str(cellule.value).count("\n") + 1
                                    max_lignes = max(max_lignes, nb_lignes)

                            worksheet.row_dimensions[row_idx].height = min(
                                120,
                                max(22, max_lignes * 18)
                            )

                    else:
                        pd.DataFrame({
                            "Info": [
                                f"Données non disponibles pour {nom_feuille}"
                            ]
                        }).to_excel(
                            writer,
                            sheet_name=nom_feuille,
                            index=False
                        )

                        
                        # =====================================================
                        # FEUILLE COMPLÈTE (toutes les colonnes) - pour référence
                        # =====================================================
                        try:
                            if not df_voyages_working.empty:
                                df_voyages_complet = df_voyages_working.copy()
                                # Formater les valeurs numériques
                                if "Poids total chargé" in df_voyages_complet.columns:
                                    df_voyages_complet["Poids total chargé"] = df_voyages_complet["Poids total chargé"].round(3)
                                if "Volume total chargé" in df_voyages_complet.columns:
                                    df_voyages_complet["Volume total chargé"] = df_voyages_complet["Volume total chargé"].round(3)
                                
                                df_voyages_complet.to_excel(writer, sheet_name='Données Complètes', index=False)
                            else:
                                pd.DataFrame(columns=list(df_voyages_working.columns)).to_excel(writer, sheet_name='Données Complètes', index=False)
                        except Exception as e:
                            print(f"⚠️ Erreur lors de la création de la feuille complète : {e}")
                            pd.DataFrame({'Erreur': ['Impossible de créer la feuille complète']}).to_excel(writer, sheet_name='Données Complètes', index=False)
            
            # =====================================================
            # MISE EN FORME PROFESSIONNELLE DE TOUTES LES FEUILLES
            # =====================================================
            from openpyxl.styles import (
                Font, PatternFill, Alignment, Border, Side
            )
            from openpyxl.utils import get_column_letter

            couleur_entete = "0369A1"
            couleur_bordure = "B0C4DE"

            bordure_fine = Side(
                style="thin",
                color=couleur_bordure
            )

            bordure_cellule = Border(
                left=bordure_fine,
                right=bordure_fine,
                top=bordure_fine,
                bottom=bordure_fine
            )

            alignement_centre = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            # Parcourir toutes les feuilles du fichier Excel
            for nom_feuille in writer.sheets:
                worksheet = writer.sheets[nom_feuille]

                # Mise en forme de toutes les cellules
                for ligne in worksheet.iter_rows():
                    for cellule in ligne:
                        cellule.alignment = alignement_centre
                        cellule.border = bordure_cellule
                        cellule.font = Font(
                            name="Arial",
                            size=10
                        )

                # Mise en forme de l'en-tête
                if worksheet.max_column > 0:
                    for cellule in worksheet[1]:
                        cellule.fill = PatternFill(
                            fill_type="solid",
                            fgColor=couleur_entete
                        )
                        cellule.font = Font(
                            name="Arial",
                            size=10,
                            bold=True,
                            color="FFFFFF"
                        )

                    worksheet.row_dimensions[1].height = 35

                # Adapter la largeur de chaque colonne
                for col_idx in range(1, worksheet.max_column + 1):
                    lettre = get_column_letter(col_idx)
                    entete = worksheet.cell(
                        row=1,
                        column=col_idx
                    ).value

                    max_longueur = 0

                    for ligne_idx in range(
                        1, worksheet.max_row + 1
                    ):
                        valeur = worksheet.cell(
                            row=ligne_idx,
                            column=col_idx
                        ).value

                        if valeur is not None:
                            lignes_texte = str(valeur).split("\n")
                            longueur = max(
                                len(texte) for texte in lignes_texte
                            )
                            max_longueur = max(
                                max_longueur, longueur
                            )

                    # Largeurs spécifiques aux colonnes importantes
                    largeurs = {
                        "Zone": 12,
                        "Véhicule N°": 15,
                        "Véhicule": 15,
                        "Code voyage": 16,
                        "Chauffeur": 22,
                        "Client(s) inclus": 32,
                        "Représentant(s) inclus": 30,
                        "BL inclus": 35,
                        "Poids total chargé": 20,
                        "Volume total chargé": 21,
                        "Type_Véhicule": 20,
                        "Type_Camion": 18,
                        "Métrique": 30,
                        "Valeur": 22,
                        "Nombre Véhicules": 20,
                        "Nombre Camions": 18,
                        "Nombre Estafettes": 20,
                        "Poids Total (kg)": 20,
                        "Poids Moyen (kg)": 20,
                        "Volume Total (m³)": 21,
                        "Volume Moyen (m³)": 21,
                        "Taux Occupation Moyen (%)": 25,
                        "Utilisation Poids (%)": 22,
                        "Utilisation Volume (%)": 23,
                    }

                    if entete in largeurs:
                        largeur = largeurs[entete]
                    else:
                        largeur = max(15, min(35, max_longueur + 2))

                    worksheet.column_dimensions[lettre].width = largeur

                # Adapter la hauteur de chaque ligne au contenu
                for ligne_idx in range(2, worksheet.max_row + 1):
                    max_lignes = 1

                    for cellule in worksheet[ligne_idx]:
                        if cellule.value is not None:
                            nb_lignes = (
                                str(cellule.value).count("\n") + 1
                            )
                            max_lignes = max(
                                max_lignes, nb_lignes
                            )

                    worksheet.row_dimensions[ligne_idx].height = min(
                        120,
                        max(22, max_lignes * 18)
                    )

                # Activer les filtres si la feuille contient des colonnes
                if worksheet.max_row >= 1 and worksheet.max_column >= 1:
                    worksheet.auto_filter.ref = worksheet.dimensions

                # Figer la première ligne
                worksheet.freeze_panes = "A2"

                # Masquer le quadrillage
                worksheet.sheet_view.showGridLines = False

                # Paramètres d'impression
                worksheet.page_setup.orientation = "landscape"
                worksheet.page_setup.fitToWidth = 1
                worksheet.page_setup.fitToHeight = 0
                worksheet.sheet_properties.pageSetUpPr.fitToPage = True
                worksheet.print_title_rows = "1:1"

        return True, f"✅ Planning exporté avec succès : {file_path}"
    
    except Exception as e:
        return False, f"❌ Erreur lors de l'export Excel : {str(e)}"
# =====================================================
# GARDEZ CETTE FONCTION INTACTE - NE PAS MODIFIER
# =====================================================
def verifier_integrite_donnees(df_voyages, df_livraisons_original):
    """Vérifie l'intégrité des données entre les voyages optimisés et les données originales."""
    try:
        problèmes = []
        
        # Vérifier que tous les BLs originaux sont présents dans les voyages
        bls_originaux = set(df_livraisons_original["No livraison"].astype(str).unique())
        bls_voyages = set()
        
        for bls in df_voyages["BL inclus"]:
            if pd.notna(bls):
                # EXCLURE les objets manuels des vérifications
                bls_filtres = [bl for bl in str(bls).split(';') if not bl.startswith('OBJ-')]
                bls_voyages.update(bls_filtres)
        
        bls_manquants = bls_originaux - bls_voyages
        bls_ajoutes = bls_voyages - bls_originaux
        
        if bls_manquants:
            problèmes.append(f"❌ BLs manquants dans les voyages : {len(bls_manquants)} BLs")
        
        if bls_ajoutes:
            problèmes.append(f"⚠️ BLs supplémentaires dans les voyages : {len(bls_ajoutes)} BLs (objets manuels exclus)")
        
        # Vérifier la cohérence des poids et volumes (EXCLURE les objets manuels)
        poids_total_originel = df_livraisons_original["Poids total"].sum()
        volume_total_originel = df_livraisons_original["Volume total"].sum()
        
        # Calculer les totaux des voyages SANS les objets manuels
        poids_total_voyages_sans_objets = 0
        volume_total_voyages_sans_objets = 0
        
        for idx, row in df_voyages.iterrows():
            bls = str(row.get("BL inclus", ""))
            if pd.notna(bls):
                # Identifier les objets manuels dans ce véhicule
                objets_manuels = [bl for bl in bls.split(';') if bl.startswith('OBJ-')]
                
                if objets_manuels:
                    # Estimer le poids/volume des objets manuels (approximatif)
                    # Ou simplement utiliser les valeurs actuelles comme référence
                    pass
            
            # Pour simplifier, utilisons les données originales comme référence
            poids_total_voyages_sans_objets += row.get("Poids total chargé", 0)
            volume_total_voyages_sans_objets += row.get("Volume total chargé", 0)
        
        # Ajuster pour les objets manuels (estimation)
        # Pour l'instant, utilisons une comparaison directe avec un message explicatif
        
        poids_total_voyages = df_voyages["Poids total chargé"].sum()
        volume_total_voyages = df_voyages["Volume total chargé"].sum()
        
        # Vérifier les écarts avec tolérance
        tolerance = 0.01  # 1%
        
        ecart_poids = abs(poids_total_originel - poids_total_voyages) / poids_total_originel
        ecart_volume = abs(volume_total_originel - volume_total_voyages) / volume_total_originel
        
        if ecart_poids > tolerance:
            problèmes.append(
                f"⚠️ Écart de poids : Original {poids_total_originel:.1f}kg vs "
                f"Voyages {poids_total_voyages:.1f}kg (diff: {poids_total_voyages-poids_total_originel:.1f}kg)"
            )
        
        if ecart_volume > tolerance:
            problèmes.append(
                f"⚠️ Écart de volume : Original {volume_total_originel:.3f}m³ vs "
                f"Voyages {volume_total_voyages:.3f}m³ (diff: {volume_total_voyages-volume_total_originel:.3f}m³)"
            )
        
        # Ajouter une note sur les objets manuels
        objets_count = sum(1 for bls in df_voyages["BL inclus"] if 'OBJ-' in str(bls))
        if objets_count > 0:
            problèmes.append(f"📦 Note : {objets_count} objet(s) manuel(s) inclus dans la planification")
        
        if not problèmes:
            return "✅ Intégrité des données vérifiée - Aucun problème détecté"
        else:
            return "\n".join(problèmes)
            
    except Exception as e:
        return f"❌ Erreur lors de la vérification d'intégrité : {str(e)}"
