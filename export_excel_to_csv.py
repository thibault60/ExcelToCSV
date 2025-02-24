import streamlit as st
import pandas as pd
import os

# 📌 Titre de l'application
st.title("📂 Convertisseur Excel ➡️ CSV")

# 📝 Description
st.markdown("""
**Instructions :**  
1️⃣ **Téléversez un fichier Excel (.xlsx)**  
2️⃣ **Cliquez sur "Convertir en CSV"**  
3️⃣ **Téléchargez le fichier CSV généré** 🎯
""")

# 🔹 Widget pour uploader un fichier
fichier_excel = st.file_uploader("📂 Téléchargez votre fichier Excel", type=["xlsx"])

if fichier_excel:
    try:
        # 🔹 Lire le fichier Excel
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
                mots_cles_principal = df.iloc[3, 0] if df.shape[0] > 3 else "N/A"  # 🔹 Mot clé principal en A4
                mots_cles = df.iloc[6:, 0].dropna().tolist()  # 🔹 Mots-clés à partir de A7

                # 🔹 Supprimer les lignes contenant "Mots-clés"
                mots_cles = [mot for mot in mots_cles if mot.lower().strip() != "mots-clés"]

                # 🔹 Convertir en une chaîne séparée par "|"
                mots_cles_str = "|".join(mots_cles)

                # 🔹 Créer le DataFrame final avec les 3 colonnes
                final_df = pd.DataFrame([[url, mots_cles_principal, mots_cles_str]],
                                        columns=["URL", "Mot Clé Principal", "Keywords"])

                # 🔹 Affichage dans Streamlit
                st.dataframe(final_df)

                # 🔹 Enregistrer en CSV
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
