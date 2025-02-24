import os
import pandas as pd

# 📂 Vérifier le dossier de travail
print("📂 Dossier de travail actuel :", os.getcwd())

# 🔹 Chemin du fichier Excel (corrigé)
fichier_excel = "/workspaces/ExcelToCSV/ExcelToCSV/occurence-Bilan-2025.xlsx"
fichier_csv = "/workspaces/ExcelToCSV/ExcelToCSV/export.csv"

# 🔹 Vérifier si le fichier existe
if not os.path.exists(fichier_excel):
    print(f"❌ ERREUR : Le fichier '{fichier_excel}' est introuvable.")
    print("📂 Vérifiez que le fichier est bien placé dans le bon dossier.")
    exit(1)

try:
    # 🔹 Charger le fichier Excel
    xls = pd.ExcelFile(fichier_excel, engine="openpyxl")
except Exception as e:
    print(f"❌ ERREUR : Impossible de lire le fichier Excel.\nDétail de l'erreur : {e}")
    exit(1)

# 🔹 Liste pour stocker les données extraites
donnees = []

# 🔹 Boucle sur chaque feuille du fichier Excel
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine="openpyxl")

    # Vérifier que le fichier contient assez de lignes
    if df.shape[0] > 6:
        url = df.iloc[2, 0]  # 🔹 URL en A3
        mots_cles = df.iloc[6:, 0].dropna().tolist()  # 🔹 Mots-clés à partir de A7

        # 🔹 Supprimer les lignes contenant "Mots-clés"
        mots_cles = [mot for mot in mots_cles if mot.lower().strip() != "mots-clés"]

        # 🔹 Convertir en une chaîne séparée par "|"
        mots_cles_str = "|".join(mots_cles)

        # 🔹 Ajouter à la liste
        donnees.append([url, mots_cles_str])

# 🔹 Vérifier s'il y a des données avant d'exporter
if donnees:
    final_df = pd.DataFrame(donnees, columns=["URL", "Keywords"])
    final_df.to_csv(fichier_csv, index=False, sep=",", encoding="utf-8-sig")
    print(f"✅ Export terminé : {os.path.abspath(fichier_csv)}")
else:
    print("⚠️ Aucune donnée valide trouvée. Le fichier CSV n'a pas été généré.")
