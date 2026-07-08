import pandas as pd
import sys

def process_file(input_file, output_file, date_rapport):
    # Charger le fichier Excel
    print(f"Chargement du fichier: {input_file}")
    df = pd.read_excel(input_file)
    
    # Filtrage des alarmes à garder dans alarmes texte
    alarmes_a_garder = [
        "BTS O&M LINK FAILURE / WCDMA BASE STATION OUT OF USE",
        "WCDMA BASE STATION OUT OF USE"
    ]
    
    print("Filtrage des données dans la colonne 'Alarm text'...")
    # On s'assure que la colonne ne contient pas des espaces au début/fin au moment de la comparaison (optionnel mais recommandé)
    df_filtre = df[df['Alarm text'].astype(str).str.strip().isin(alarmes_a_garder)].copy()
    
    # Traitement des dates pour le rapport du jour spécifié
    date_jour = date_rapport
    debut_jour = pd.to_datetime(f"{date_jour} 00:00:00")
    fin_jour = pd.to_datetime(f"{date_jour} 23:59:00")

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
    
    # 3. Supprimer tout ce qui ne s'est pas passé le 14 (les lignes doivent maintenant avoir Alarm Time et Cancel Time le 14)
    # C'est à dire, exclure ceux qui se sont terminés avant le 14, ou qui ont commencé après le 14
    cond_garder = (df_filtre['Alarm Time Tmp'] >= debut_jour) & (df_filtre['Alarm Time Tmp'] <= fin_jour) & (df_filtre['Cancel Time Tmp'] >= debut_jour) & (df_filtre['Cancel Time Tmp'] <= fin_jour)
    df_filtre = df_filtre[cond_garder].copy()
    
    # Calcul de la durée dans la colonne "Duration" en utilisant les heures temporaires (bornées à la journée)
    df_filtre['Duration'] = df_filtre['Cancel Time Tmp'] - df_filtre['Alarm Time Tmp']
    
    # Remplacer les valeurs originales par les valeurs bornées pour correspondre au fichier attendu
    df_filtre['Alarm Time'] = df_filtre['Alarm Time Tmp'].dt.strftime('%d-%m-%Y %H:%M:%S')
    
    # Pour Cancel Time, garder vide si c'était NaT à l'origine, sinon mettre la valeur bornée
    cancel_isna_orig = df['Cancel Time'].isna()
    # Mettre à jour avec le format string, et remplacer par NaN (ou vide) si besoin
    # Attention: df_filtre['Cancel Time'] = df_filtre['Cancel Time Tmp'].dt.strftime('%d-%m-%Y %H:%M:%S')
    # Les NaT n'ont pas été remplacés par des dates futures dans les colonnes Tmp dans ce script,
    # c'était fait en condition. La ligne 32 le fait de façon conditionnelle.
    # Appliquons le formatage en chaîne de caractères pour forcer Excel à afficher l'heure
    df_filtre['Alarm Time'] = df_filtre['Alarm Time Tmp'].dt.strftime('%d/%m/%Y %H:%M:%S')
    df_filtre['Cancel Time'] = df_filtre['Cancel Time Tmp'].dt.strftime('%d/%m/%Y %H:%M:%S')
    
    # Remplacer les valeurs NaT formatées (qui deviennent NaN ou 'NaT') par des chaînes vides
    df_filtre['Cancel Time'] = df_filtre['Cancel Time'].fillna('').replace('NaT', '')
    
    # Formatage de la durée en HH:MM:SS
    df_filtre['Duration'] = df_filtre['Duration'].apply(
        lambda x: f"{int(x.total_seconds() // 3600):02d}:{int((x.total_seconds() % 3600) // 60):02d}:{int(x.total_seconds() % 60):02d}" if pd.notnull(x) else ""
    )
    
    # Supprimer les colonnes temporaires pour garder le fichier propre
    df_filtre = df_filtre.drop(columns=['Alarm Time Tmp', 'Cancel Time Tmp'])
    
    # Supprimer l'ancienne colonne 'DURATION' si elle existe pour éviter un doublon
    if 'DURATION' in df_filtre.columns:
        df_filtre = df_filtre.drop(columns=['DURATION'])
    
    # Formater les dates pour l'exportation selon le format demandé (optionnel)
    # df_filtre['DURATION'] peut être formaté en HH:MM:SS ou laissé en timedelta
    
    # Exporter le résultat
    print(f"Sauvegarde du fichier traité: {output_file}")
    df_filtre.to_excel(output_file, index=False)
    print(f"Fichier traité avec succès. Lignes retenues: {len(df_filtre)} sur {len(df)}")

if __name__ == "__main__":
    fichier_entree = "TAGBA_Tchédré Saturnin.xlsx"
    fichier_sortie = "Nouveau_Fichier_Traiter.xlsx"
    date_a_traiter = "2026-03-14" # Remplacez cette date par celle du jour que vous voulez traiter
    process_file(fichier_entree, fichier_sortie, date_a_traiter)