# ==============================================================================
# SCRIPT AUTOMATISÉ DQAT - VERSION 19 (MOTEUR DE NETTOYAGE MULTI-FEUILLES)
# Configuration : GitHub Actions avec Sauvegarde Directe Google Drive API
# Zone Horaire : Afrique/Kinshasa (UTC+1)
# ==============================================================================

import os
import json
import base64
import time
from datetime import datetime, timedelta, timezone
import requests
import pandas as pd
import urllib3

# Importations pour l'API Google Drive
import google.auth
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Désactivation des alertes de sécurité SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuration du fuseau horaire de Kinshasa (UTC+1)
TZ_KINSHASA = timezone(timedelta(hours=1))

# ID de votre dossier Google Drive partagé
ID_DOSSIER_RACINE_DRIVE = "1RMriGUzLVa_O1Cf28OUPTlF2Jl0QEXn-"

# Dictionnaire officiel des favoris validés (UID PEC actualisé)
FAVORIS_MAPPING = {
    "Wy4zf2qLSI6": "DQAT_2025_Depistage_Pos",
    "cT2cgP86CRc": "DQAT_2025_Depistage",
    "hfko1E1ezWT": "DQAT_2025_PEC",
    "HslmfN4uaAj": "DQAT_2025_PTME",
    "eU2NiPPtbKv": "Suivi FA (6 derniers mois)"
}

def initialiser_drive_service():
    try:
        # Authentification automatique via la clé GCP_SA_KEY configurée sur GitHub
        creds, _ = google.auth.default(scopes=['https://www.googleapis.com/auth/drive'])
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        print(f"[X] Erreur d'initialisation de l'API Google Drive : {str(e)}")
        raise e

def creer_dossier_drive(service, nom_dossier, parent_id):
    metadata = {
        'name': nom_dossier,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    dossier = service.files().create(body=metadata, fields='id').execute()
    return dossier.get('id')

def uploader_fichier_drive(service, chemin_local, nom_fichier, dossier_destination_id):
    metadata = {'name': nom_fichier, 'parents': [dossier_destination_id]}
    
    if chemin_local.endswith('.xlsx'):
        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    elif chemin_local.endswith('.json'):
        mime = 'application/json'
    else:
        mime = 'text/plain'
        
    media = MediaFileUpload(chemin_local, mimetype=mime, resumable=True)
    service.files().create(body=metadata, media_body=media, fields='id').execute()
    print(f"  [➔] Transféré sur Google Drive : {nom_fichier}")

def initialiser_environnement_local():
    print("====================================================================")
    print("          LANCEMENT DU MOTEUR AUTOMATISÉ DQAT ONLINE - V19")
    print("====================================================================\n")
    
    chemin_racine = './DQAT_online_temp'
    os.makedirs(chemin_racine, exist_ok=True)
    
    nom_session = datetime.now(TZ_KINSHASA).strftime("Session_%Y-%m-%d_%Hh%Mm%Ss")
    chemin_destination = os.path.join(chemin_racine, nom_session)
    os.makedirs(chemin_destination, exist_ok=True)
    return chemin_destination, nom_session

def recuperer_identifiants():
    try:
        username = os.environ.get('DHIS2_USERG')
        password = os.environ.get('DHIS2_PASSWORDG')
        if not username or not password:
            raise ValueError("Les variables d'environnement sont vides.")
        return username, password
    except Exception as e:
        print("\n[X] ERREUR : Clés 'DHIS2_USERG' ou 'DHIS2_PASSWORDG' introuvables.\n")
        raise e

def recuperer_metadonnees_geo_ciblees(session, base_url, liste_uids, username, password):
    mapping_geo = {}
    if not liste_uids:
        return mapping_geo

    taille_paquet = 100
    for i in range(0, len(liste_uids), taille_paquet):
        paquet = liste_uids[i:i+taille_paquet]
        filtre_ids = ",".join(paquet)
        url = f"{base_url}/api/organisationUnits.json?filter=id:in:[{filtre_ids}]&fields=id,name,level,ancestors[level,name]&paging=false"
        try:
            response = session.get(url, timeout=60, verify=False, auth=(username, password))
            if response.status_code == 200:
                data = response.json()
                for ou in data.get('organisationUnits', []):
                    ou_id = ou.get('id')
                    ou_name = ou.get('name')
                    province = "Inconnue"
                    zone_sante = "Inconnue"

                    for ancestor in ou.get('ancestors', []):
                        if ancestor.get('level') == 2:
                            province = ancestor.get('name')
                        elif ancestor.get('level') == 3:
                            zone_sante = ancestor.get('name')

                    if ou.get('level') == 2:
                        province = ou_name
                    elif ou.get('level') == 3:
                        zone_sante = ou_name

                    mapping_geo[ou_id] = {
                        "Province": province, "Zone de Sante": zone_sante, "Nom du Site": ou_name
                    }
        except Exception:
            continue
    return mapping_geo

def extraire_et_convertir_favoris(base_url, mapping_favoris, username, password, dossier_local, drive_service, id_session_drive):
    session = requests.Session()
    auth_str = f"{username}:{password}"
    b64_auth = base64.b64encode(auth_str.encode('utf-8')).decode('utf-8')

    session.headers.update({
        'Authorization': f'Basic {b64_auth}',
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/json'
    })

    rapport_extraction = []
    dataframes_sauvegardes = {}

    for index, (uid, nom_personnalise) in enumerate(mapping_favoris.items(), start=1):
        print(f"\n--- Traitement du Favori {index}/{len(mapping_favoris)} [{nom_personnalise}] ---")

        endpoints_donnees = {
            "Moteur Analytique Universel" : f"{base_url}/api/analytics.json?visualization={uid}",
            "Visualisation Analytique"    : f"{base_url}/api/visualizations/{uid}/data.json",
            "Tableau Pivot Traditionnel"  : f"{base_url}/api/reportTables/{uid}/data.json"
        }

        donnees_json = None
        succes_donnees = False
        nb_lignes = 0
        derniere_erreur = "Aucune"
        temps_debut = time.time()

        for nom_route, url in endpoints_donnees.items():
            try:
                print(f"  -> Requête via : {nom_route}...")
                response = session.get(url, timeout=180, verify=False, auth=(username, password))
                if response.status_code == 200:
                    donnees_json = response.json()
                    if 'rows' in donnees_json and len(donnees_json['rows']) > 0:
                        succes_donnees = True
                        nb_lignes = len(donnees_json['rows'])
                        break
                else:
                    derniere_erreur = f"Erreur HTTP {response.status_code}"
            except Exception as e:
                derniere_erreur = f"Incident réseau : {type(e).__name__}."
                continue

        temps_fin = time.time()
        duree_minutes = round((temps_fin - temps_debut) / 60, 2)
        rapport_extraction.append({
            "Nom": nom_personnalise, "UID": uid, "Réussite": "yes" if succes_donnees else "no",
            "Temps": duree_minutes, "Lignes": nb_lignes, "Erreur": derniere_erreur
        })

        if not succes_donnees:
            continue

        # Sauvegarde et Envoi JSON
        nom_fichier_json = f"{nom_personnalise}.json"
        chemin_json = os.path.join(dossier_local, nom_fichier_json)
        with open(chemin_json, 'w', encoding='utf-8') as f:
            json.dump(donnees_json, f, ensure_ascii=False, indent=4)
        uploader_fichier_drive(drive_service, chemin_json, nom_fichier_json, id_session_drive)

        # Nettoyage et géographie
        try:
            colonnes = [h.get('column', h.get('name')) for h in donnees_json['headers']]
            lignes = donnees_json['rows']
            df = pd.DataFrame(lignes, columns=colonnes)

            colonne_uid_cible = None
            for potentiel_header in ['organisationunitid', 'ou', 'organisationUnit', 'organisationunit']:
                if potentiel_header in df.columns:
                    colonne_uid_cible = potentiel_header
                    break

            if colonne_uid_cible:
                uids_uniques = df[colonne_uid_cible].dropna().unique().tolist()
                mapping_geo = recuperer_metadonnees_geo_ciblees(session, base_url, uids_uniques, username, password)

                if mapping_geo:
                    df['Province'] = df[colonne_uid_cible].map(lambda x: mapping_geo.get(x, {}).get('Province', 'Inconnue'))
                    df['Zone de Sante'] = df[colonne_uid_cible].map(lambda x: mapping_geo.get(x, {}).get('Zone de Sante', 'Inconnue'))
                    df['Nom du Site'] = df[colonne_uid_cible].map(lambda x: mapping_geo.get(x, {}).get('Nom du Site', 'Inconnu'))
                    df['UID Site'] = df[colonne_uid_cible]

                    colonnes_geo = ['Province', 'Zone de Sante', 'UID Site', 'Nom du Site']
                    colonnes_a_exclure = set(colonnes_geo + ['organisationunitid', 'organisationunitname', 'organisationunitcode', 'organisationunitdescription', 'ou'])
                    autres_colonnes = [c for c in df.columns if c not in colonnes_a_exclure]
                    df = df[colonnes_geo + autres_colonnes]

            nom_fichier_excel = f"{nom_personnalise}.xlsx"
            chemin_excel = os.path.join(dossier_local, nom_fichier_excel)
            with pd.ExcelWriter(chemin_excel, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Données DQAT', index=False)
            
            uploader_fichier_drive(drive_service, chemin_excel, nom_fichier_excel, id_session_drive)
            dataframes_sauvegardes[nom_personnalise] = df

        except Exception as err:
            print(f"  [X] Erreur Excel sur {nom_personnalise} : {str(err)}")

    # Consolidation Master
    print("\n====================================================================")
    print("  ÉTAPE 3 : COMPILATION MULTI-FEUILLES MASTER")
    print("====================================================================")

    ref_name = "DQAT_2025_Depistage_Pos"
    if ref_name in dataframes_sauvegardes:
        df_master = dataframes_sauvegardes[ref_name].copy()
        favoris_a_greffer = ["DQAT_2025_Depistage", "DQAT_2025_PEC", "DQAT_2025_PTME"]
        colonnes_gestion_doublons = ['Province', 'Zone de Sante', 'Nom du Site', 'periodname', 'periodcode', 'perioddescription', 'reporting_month_name', 'param_organisationunit_name', 'organisation_unit_is_parent', 'organisationunitname', 'A27']

        for cible in favoris_a_greffer:
            if cible in dataframes_sauvegardes:
                df_cible = dataframes_sauvegardes[cible].copy().drop_duplicates(subset=['UID Site', 'periodid'])
                colonnes_indicateurs_propres = [c for c in df_cible.columns if c not in colonnes_gestion_doublons and c not in ['UID Site', 'periodid']]
                if colonnes_indicateurs_propres:
                    df_cible_filtre = df_cible[['UID Site', 'periodid'] + colonnes_indicateurs_propres]
                    df_master = pd.merge(df_master, df_cible_filtre, on=['UID Site', 'periodid'], how='left')

        df_cleaned = df_master.copy()
        df_cleaned.columns = [c.strip() for c in df_cleaned.columns]
        df_cleaned = df_cleaned.drop(columns=[c for c in ['periodcode', 'perioddescription', 'reporting_month_name', 'param_organisationunit_name', 'organisation_unit_is_parent', 'A27', 'PNLS_CU_4.7_FE VIH+ mises sous TARV SA/PP_y'] if c in df_cleaned.columns], errors='ignore')

        mapping_etape_1 = {
            "PNLS_CU_4.7_FE VIH+ informés résultats SA/PP": "TEMP_47_DIAGNOSTIQUEES",
            "PNLS_CU_4.7_FE retiré les résultats SA/PP": "TEMP_47_INFORMES",
            "PNLS_CU_4.7_Diagnostiqués VIH+ SA/PP": "TEMP_47_CONSEILLES",
            "PNLS_CU_4.7_Diagnostiquées VIH+ SA/PP": "TEMP_47_CONSEILLES",
            "PNLS_CU_4.7_FE Conseillés et testés SA/PP": "TEMP_47_RETIRE",
            "PNLS_CU_4.7_FE VIH+ mises sous TARV SA/PP_x": "PNLS_CU_4.7_FE VIH+ mises sous TARV SA/PP",
            "B114": "PNLS_CU_3.4.4_AZT/3TC +ATV/r"
        }
        df_cleaned = df_cleaned.rename(columns=mapping_etape_1)
        
        mapping_etape_2 = {
            "TEMP_47_DIAGNOSTIQUEES": "PNLS_CU_4.7_Diagnostiquées VIH+ SA/PP",
            "TEMP_47_INFORMES": "PNLS_CU_4.7_FE VIH+ informés résultats SA/PP",
            "TEMP_47_CONSEILLES": "PNLS_CU_4.7_Conseillés et testés SA/PP",
            "TEMP_47_RETIRE": "PNLS_CU_4.7_FE retiré les résultats SA/PP"
        }
        df_cleaned = df_cleaned.rename(columns=mapping_etape_2)

        nom_master_excel = "DQAT_2025_Compilation_Master.xlsx"
        chemin_master_excel = os.path.join(dossier_local, nom_master_excel)
        with pd.ExcelWriter(chemin_master_excel, engine='openpyxl') as writer:
            df_master.to_excel(writer, sheet_name='Compilation_Brute', index=False)
            df_cleaned.to_excel(writer, sheet_name='Compilation_Nettoyée', index=False)
        
        uploader_fichier_drive(drive_service, chemin_master_excel, nom_master_excel, id_session_drive)

    # Résumé txt
    chemin_resume_txt = os.path.join(dossier_local, "resume_extraction.txt")
    with open(chemin_resume_txt, 'w', encoding='utf-8') as f_txt:
        f_txt.write(f"RÉSUMÉ DQAT V19 - Date : {datetime.now(TZ_KINSHASA).strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        for r in rapport_extraction:
            f_txt.write(f"Favori : {r['Nom']} | Réussite : {r['Réussite']} | Lignes : {r['Lignes']}\n")
    uploader_fichier_drive(drive_service, chemin_resume_txt, "resume_extraction.txt", id_session_drive)
    
    session.close()

if __name__ == "__main__":
    URL_BASE_DHIS2 = "https://snisrdc.com"

    try:
        drive_serv = initialiser_drive_service()
        dossier_local, nom_session = initialiser_environnement_local()
        
        print(f"[-] Création du dossier de session sur Google Drive...")
        id_session_drive = creer_dossier_drive(drive_serv, nom_session, ID_DOSSIER_RACINE_DRIVE)
        
        nom_utilisateur, mot_de_passe = recuperer_identifiants()
        extraire_et_convertir_favoris(URL_BASE_DHIS2, FAVORIS_MAPPING, nom_utilisateur, mot_de_passe, dossier_local, drive_serv, id_session_drive)

        print("\n====================================================================")
        print("[+] TERMINÉ ! TOUS LES FICHIERS ONT ÉTÉ ENVOYÉS SUR GOOGLE DRIVE.")
        print("====================================================================")
    except Exception as e_main:
        print(f"\n[X] Le processus général a échoué : {str(e_main)}")