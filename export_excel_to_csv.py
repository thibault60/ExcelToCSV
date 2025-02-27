import streamlit as st
import pandas as pd
import os

# 📌 Titre de l'application
st.title("📂 Convertisseur Excel ➡️ CSV avec Fréquence Conseillée")

# 📝 Instructions
st.markdown("""
**Instructions :**  
1️⃣ **Téléversez un fichier Excel (.xlsx)**  
2️⃣ **Sélectionnez la feuille contenant les données**  
3️⃣ **Cliquez sur "Convertir en CSV"**  
4️⃣ **Téléchargez le fichier CSV généré** 🎯  
""")

# 🔹 Widget pour uploader un fichier
fichier_excel = st.file_uploader("📂 Téléchargez votre fichier Excel", type=["xlsx"])

if fichier_excel:
    try:
        # 🔹 Charger le fichier Excel
        xls = pd.ExcelFile(fichier_excel, engine="openpyxl")
        st.success("✅ Fichier chargé avec succès !")

        # 🔹 Sélection de la feuille (si plusieurs feuilles)
        sheet_name = st.selectbox("📑 Sélectionnez une feuille :", xls.sheet_names)

        if st.button("🚀 Convertir en CSV"):
            # 🔹 Lire la feuille sélectionnée
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine="openpyxl")

            # Vérifier que le fichier contient assez de lignes
            if df.shape[0] > 6:
                url = df.iloc[2, 0]  # 🔹 URL en A3
                mot_cle_principal = f"{df.iloc[4, 0]} {df.iloc[4, 1]}".strip()  # 🔹 Concaténation A5 et B5
                mots_cles = df.iloc[6:, 0].dropna().tolist()  # 🔹 Mots-clés à partir de A7
                frequence_conseillee = df.iloc[6:, 1].dropna().tolist()  # 🔹 Fréquence conseillée en B7 et après

                # 🔹 Supprimer les lignes contenant "Mots-clés"
                mots_cles = [mot for mot in mots_cles if mot.lower().strip() != "mots-clés"]

                # 🔹 Convertir en une chaîne séparée par "|"
                mots_cles_str = "|".join(mots_cles)
                frequence_conseillee_str = "|".join(map(str, frequence_conseillee))  # Convertir en chaîne
                
                # 🔹 Création du DataFrame avec 4 colonnes
                final_df = pd.DataFrame({
                    "URL": [url],
                    "Mot Clé Principal": [mot_cle_principal],
                    "Keywords": [mots_cles_str],
                    "Fréquence conseillée": [frequence_conseillee_str]
                })

                # 🔹 Affichage du tableau en 4 colonnes dans Streamlit
                st.write("📊 **Aperçu des données extraites** :")
                st.dataframe(final_df)

                # 🔹 Enregistrement en CSV
                csv_path = "export.csv"
                final_df.to_csv(csv_path, index=False, sep=",", encoding="utf-8-sig")

                # 🔹 Télécharger le fichier CSV
                with open(csv_path, "rb") as file:
                    st.download_button("📥 Télécharger le CSV", file, "export.csv", "text/csv")
                    st.success("✅ Conversion terminée, cliquez pour télécharger le CSV !")
            else:
                st.warning("⚠️ Le fichier ne contient pas assez de lignes pour être traité.")

    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture du fichier : {e}")
