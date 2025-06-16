import streamlit as st
import pandas as pd
from io import BytesIO

# 📌 Titre de l'application
st.title("📂 Convertisseur Excel ➡️ CSV (consolidé)")

# 📝 Instructions
st.markdown("""
**Mode d'emploi :**  
1️⃣ Téléversez un fichier Excel (.xlsx)  
2️⃣ Cliquez sur **« Convertir & Télécharger »** pour obtenir un CSV unique consolidant toutes les feuilles  
""")

# 🔹 Widget pour uploader un fichier
fichier_excel = st.file_uploader("📂 Téléchargez votre fichier Excel", type=["xlsx"])

if fichier_excel:
    try:
        # 🔹 Lire toutes les feuilles sous forme de dictionnaire {nom_feuille: DataFrame}
        sheets_dict = pd.read_excel(fichier_excel, sheet_name=None, header=None, engine="openpyxl")
        st.success(f"✅ {len(sheets_dict)} feuille(s) détectée(s) : {', '.join(sheets_dict.keys())}")

        lignes = []  # contiendra les lignes finales issues de chaque feuille

        for sheet_name, df in sheets_dict.items():
            # --- Contrôle du volume de lignes ---
            if df.shape[0] <= 6:
                st.warning(f"⚠️ Feuille « {sheet_name} » ignorée (moins de 7 lignes).")
                continue

            # --- Extraction des données suivant votre structure ---
            url = df.iloc[2, 0]  # A3
            mot_cle_principal = f"{df.iloc[4, 0]} {df.iloc[4, 1]}".strip()  # A5 + B5
            mots_cles = df.iloc[6:, 0].dropna().tolist()  # A7 et après
            frequence = df.iloc[6:, 1].dropna().tolist()  # B7 et après

            # Supprimer une éventuelle ligne d'en-tête « Mots-clés »
            mots_cles = [m for m in mots_cles if m.lower().strip() != "mots-clés"]

            # --- Construction de la ligne finale ---
            lignes.append({
                "URL": url,
                "Mot Clé Principal": mot_cle_principal,
                "Keywords": " | ".join(mots_cles),  # espacement autour du pipe
                "Fréquence conseillée": " | ".join(map(str, frequence))  # idem
            })

        # --- Consolidation & téléchargement ---
        if lignes:
            consolidated_df = pd.DataFrame(lignes)
            st.write("📊 **Aperçu des données consolidées** :")
            st.dataframe(consolidated_df, use_container_width=True)

            buffer = BytesIO()
            consolidated_df.to_csv(buffer, index=False, sep=",", encoding="utf-8-sig")
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le CSV consolidé",
                data=buffer,
                file_name="export_consolidé.csv",
                mime="text/csv"
            )
        else:
            st.error("❌ Aucune feuille valide à convertir.")
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
