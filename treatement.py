import pandas as pd

def process_file(input_file, date_rapport, date_fin=None):
    """Traite le fichier Excel pour une plage de dates.

    Args:
        input_file: chemin vers le fichier Excel source.
        date_rapport: date de début (str YYYY-MM-DD ou date).
        date_fin: date de fin inclusive (str YYYY-MM-DD ou date).
                  Si None, la plage porte uniquement sur date_rapport (rapport journalier).
    """
    # Charger le fichier Excel
    print(f"Chargement du fichier: {input_file}")
    df = pd.read_excel(input_file)
    
    # Filtrage des alarmes
    alarmes_a_garder = [
        "BTS O&M LINK FAILURE / WCDMA BASE STATION OUT OF USE",
        "WCDMA BASE STATION OUT OF USE",
        "BTS O&M LINK FAILURE",
        "ALL RFMS MISSING"
    ]
    
    print("Filtrage des données dans la colonne 'Alarm text'...")
    # On s'assure que la colonne ne contient pas des espaces au début/fin au moment de la comparaison (optionnel mais recommandé)
    df_filtre = df[df['Alarm text'].astype(str).str.strip().isin(alarmes_a_garder)].copy()
    
    # Plage de dates : début = 00:00:00 du premier jour, fin = 23:59:00 du dernier jour
    debut_jour = pd.to_datetime(f"{date_rapport} 00:00:00")
    if date_fin is None:
        fin_jour = pd.to_datetime(f"{date_rapport} 23:59:00")
    else:
        fin_jour = pd.to_datetime(f"{date_fin} 23:59:00")

    # Utiliser dayfirst=True pour parser le format jj-mm-aaaa HH:MM:SS
    # On garde les colonnes originales intactes pour l'affichage final, on crée des colonnes temporaires pour le calcul
    df_filtre['Alarm Time Tmp'] = pd.to_datetime(df_filtre['Alarm Time'], dayfirst=True, errors='coerce')
    df_filtre['Cancel Time Tmp'] = pd.to_datetime(df_filtre['Cancel Time'], dayfirst=True, errors='coerce')

    # Remplir les Cancel Time NaT (toujours ouvert) par une date dans le futur éloignée pour simplifier la logique
    # ou on peut traiter les NaT en utilisant la fonction mask
    
    # 1. Incident commençant avant le 14 et se terminant le 14 (ou après) ou toujours ouvert
    cond_avant = (df_filtre['Alarm Time Tmp'] < debut_jour) & ((df_filtre['Cancel Time Tmp'] >= debut_jour) | df_filtre['Cancel Time Tmp'].isna())
    df_filtre.loc[cond_avant, 'Alarm Time Tmp'] = debut_jour

    # 2. Incident ouvert (commençant) le 14 (ou avant) et se terminant après le 14 ou toujours ouvert
    cond_apres_ou_ouvert = (df_filtre['Alarm Time Tmp'] <= fin_jour) & ((df_filtre['Cancel Time Tmp'] > fin_jour) | df_filtre['Cancel Time Tmp'].isna())
    df_filtre.loc[cond_apres_ou_ouvert, 'Cancel Time Tmp'] = fin_jour
    
    # 3. Supprimer tout ce qui ne s'est pas passé le JOUR J (les lignes doivent maintenant avoir Alarm Time et Cancel Time le JOUR J)
    # C'est à dire, exclure ceux qui se sont terminés avant le JOUR J, ou qui ont commencé après le JOUR J
    cond_garder = (df_filtre['Alarm Time Tmp'] >= debut_jour) & (df_filtre['Alarm Time Tmp'] <= fin_jour) & (df_filtre['Cancel Time Tmp'] >= debut_jour) & (df_filtre['Cancel Time Tmp'] <= fin_jour)
    df_filtre = df_filtre[cond_garder].copy()
    
    # Calcul de la durée dans la colonne "Duration" en utilisant les heures temporaires (bornées à la journée)
    duree_timedelta = df_filtre['Cancel Time Tmp'] - df_filtre['Alarm Time Tmp']
    df_filtre['Duration_Sec'] = duree_timedelta.dt.total_seconds()
    
    # Formatage de la durée en HH:MM:SS
    # .astype(str) forces string dtype even when the series is empty (avoids Timedelta from .sum())
    df_filtre['Duration'] = duree_timedelta.apply(
        lambda x: f"{int(x.total_seconds() // 3600):02d}:{int((x.total_seconds() % 3600) // 60):02d}:{int(x.total_seconds() % 60):02d}" if pd.notnull(x) else ""
    ).astype(str)

    # === NOUVEAUTÉ / DEMANDE : Mettre à jour l'affichage des colonnes Alarm & Cancel Time
    # Les pannes qui ont commencé avant le jour J ou finissent après sont affichées 
    # avec la borne 00:00:00 et 23:59:00 respectivement.
    # On écrase les colonnes originales avec nos calculs temporaires "bornés"
    df_filtre['Alarm Time'] = df_filtre['Alarm Time Tmp']
    df_filtre['Cancel Time'] = df_filtre['Cancel Time Tmp']

    # Supprimer les colonnes temporaires pour garder le fichier propre
    df_filtre = df_filtre.drop(columns=['Alarm Time Tmp', 'Cancel Time Tmp'])
    if 'DURATION' in df_filtre.columns:
        df_filtre = df_filtre.drop(columns=['DURATION'])

    # Sauvegarder TOUTES les alarmes pour l'export complet (Avant dédoublonnage)
    df_complet = df_filtre.sort_values('Alarm Time').copy()
    
    # === NOUVEAUTÉ : SUPPRESSION DES DOUBLONS (UNIQUEMENT POUR LA SYNTHÈSE) ===
    # On dédoublonne uniquement pour calculer le tableau de synthèse.
    # Un doublon est un événement qui a la MÊME heure (Alarm Time) ET le MÊME site racine.
    # Pour déterminer le site racine, on regarde en priorité le "Site Parent".
    # - Si le "Site Parent" est renseigné, c'est lui la racine.
    # - Si le "Site Parent" est N/A ou vide, on prend le "Site Name".
    # Deux pannes à la même heure sur le même "site racine" sont considérées comme un doublon.

    df_pour_synthese = df_filtre.copy()
    
    # Créer une colonne temporaire 'Site Racine' pour la déduplication
    if 'Site Parent' in df_pour_synthese.columns and 'Site Name' in df_pour_synthese.columns:
        # On remplit 'Site Racine' avec 'Site Parent'
        # Si 'Site Parent' est nul/vide/na/N/A, on comble par le 'Site Name'
        df_pour_synthese['Site Racine'] = df_pour_synthese['Site Parent'].replace(['', 'N/A', 'nan', 'NaN'], pd.NA)
        df_pour_synthese['Site Racine'] = df_pour_synthese['Site Racine'].fillna(df_pour_synthese['Site Name'])
        
        # Supprimer les doublons basés sur la racine ET l'heure
        df_pour_synthese = df_pour_synthese.drop_duplicates(subset=['Site Racine', 'Alarm Time'], keep='first')
        
        # Nettoyer la colonne temporaire
        df_pour_synthese = df_pour_synthese.drop(columns=['Site Racine'])
    else:
        # Fallback si les colonnes spécifiques n'existent pas
        df_pour_synthese = df_pour_synthese.drop_duplicates(subset=['Site Parent', 'Site Name', 'Alarm Time'], keep='first')
    
    # --- CREATION DU TABLEAU DE SYNTHESE (ESCALADE) ---
    rapport_lignes = []
    escalades_ordre = [
        "ENERGIE", "RAN", "TRANS FH", "ENERGIE / TRANS / RAN", 
        "TRANS / RAN", "INFRA", "PROJET", "TRANS FO", 
        "TRANS FTTM", "TRANS IP", "ENVIRONNEMENT", "BSS","TRANS FH-FIELD O","RAN-FIELD O",
    ]
    
    for esc in escalades_ordre:
        # Données SANS doublons pour cette escalade
        df_esc_synth = df_pour_synthese[df_pour_synthese['Escalade'] == esc] if 'Escalade' in df_pour_synthese.columns else pd.DataFrame()
        # Données AVEC doublons (Fichier complet) pour cette escalade
        df_esc_comp = df_complet[df_complet['Escalade'] == esc] if 'Escalade' in df_complet.columns else pd.DataFrame()
        
        count = len(df_esc_synth)
        
        if count > 0:
            # DUREE dedupliquée
            duree_totale_sec = df_esc_synth['Duration_Sec'].sum()
            # MTTR de la durée dédupliquée
            mttr_sec = duree_totale_sec / count
            # OUTAGE est le temps global (AVEC les doublons de la vue complète)
            outage_sec = df_esc_comp['Duration_Sec'].sum()
            
            # Compter ceux qui sont "OUVERT"
            if 'Status' in df_esc_synth.columns:
                non_resolu = len(df_esc_synth[df_esc_synth['Status'].astype(str).str.upper() == 'OUVERT'])
            else:
                non_resolu = len(df_esc_synth[df_esc_synth['Cancel Time'].isna()])
                
            if non_resolu > 0:
                statut_text = f"{non_resolu} Non resolu"
            else:
                statut_text = "Résolu"
        else:
            duree_totale_sec = 0
            mttr_sec = 0
            outage_sec = 0
            statut_text = "N/A"
            
        # Formatage Secondes -> HH:MM:SS
        def format_sec(secs):
            if pd.isna(secs): return "0:00:00"
            return f"{(secs // 3600)}:{int((secs % 3600) // 60):02d}:{int(secs % 60):02d}"
            
        rapport_lignes.append({
            "Escalade": esc,
            "Inc count": count,
            "DUREE": format_sec(duree_totale_sec),
            "MTTR": format_sec(mttr_sec),
            "OUTAGE": format_sec(outage_sec),
            "Status": statut_text
        })
        
    df_synthese = pd.DataFrame(rapport_lignes)
    
    # Ligne de TOTAL
    total_count = df_synthese['Inc count'].sum()
    
    if 'Duration_Sec' in df_complet.columns:
        # Somme pure des calculs des lignes (Données sans doublons)
        total_duree = df_pour_synthese[df_pour_synthese['Escalade'].isin(escalades_ordre)]['Duration_Sec'].sum() if 'Escalade' in df_pour_synthese.columns else 0
        # OUTAGE = Le grand total du fichier complet
        total_outage = df_complet['Duration_Sec'].sum()
    else:
        total_duree = 0
        total_outage = 0
        
    total_mttr = total_duree / total_count if total_count > 0 else 0
    
    total_row = pd.DataFrame([{
        "Escalade": "TOTAL",
        "Inc count": total_count,
        "DUREE": format_sec(total_duree),
        "MTTR": format_sec(total_mttr),
        "OUTAGE": format_sec(total_outage),
        "Status": ""
    }])
    df_synthese = pd.concat([df_synthese, total_row], ignore_index=True)
    
    # Exporter le résultat détaillé COMPLET (sans dédoublonnage)
    # L'utilisateur a demandé d'afficher ce fichier là dans l'aperçu
    df_export = df_complet.drop(columns=['Duration_Sec']) # Nettoyer avant export
    
    # Nettoyer également df_pour_synthese des colonnes temporaires si nécessaire pour l'affichage
    if 'Duration_Sec' in df_pour_synthese.columns:
        df_pour_synthese_clean = df_pour_synthese.drop(columns=['Duration_Sec'])
    else:
        df_pour_synthese_clean = df_pour_synthese.copy()
    
    return df_export, df_pour_synthese_clean, df_synthese

def obtenir_donnees_propres():
    # En environnement Power BI, le chemin de base peut être différent. 
    # Mieux vaut utiliser des chemins absolus vers votre fichier Excel
    chemin_projet = r"C:\Users\user\OneDrive\Desktop\Projet_rapport_reseau_mobile"
    fichier_entree = f"{chemin_projet}\\TAGBA_Tchédré Saturnin.xlsx"
    date_a_traiter = "2026-03-14"
    fichier_sortie = f"{chemin_projet}\\Nouveau_Fichier_Traiter.xlsx" 
    
    return process_file(fichier_entree, date_a_traiter)

if __name__ == "__main__":
    fichier_entree = "TAGBA_Tchédré Saturnin.xlsx"
    fichier_sortie = "Nouveau_Fichier_Traiter.xlsx"
    date_a_traiter = "2026-03-14" # Remplacez cette date par celle du jour que vous voulez traiter
    df_export_out, _, df_synthese_out = process_file(fichier_entree, date_a_traiter)
    df_export_out.to_excel(fichier_sortie, index=False)
    fichier_synthese = fichier_sortie.replace(".xlsx", "_Synthese.xlsx")
    df_synthese_out.to_excel(fichier_synthese, index=False)
    print("Fichiers créés avec succès.")