import streamlit as st
import pandas as pd
import os
from io import BytesIO

# 📌 Titre de l'application
st.title("📂 Convertisseur Excel ➡️ CSV (toutes feuilles)")

# 📝 Instructions
st.markdown("""
**Mode d'emploi :**  
1️⃣ Téléversez un fichier Excel (.xlsx)  
2️⃣ Cliquez sur **« Convertir toutes les feuilles »**  
3️⃣ Téléchargez, feuille par feuille, les CSV générés ✅  
""")

# 🔹 Widget pour uploader un fichier
fichier_excel = st.file_uploader("📂 Téléchargez votre fichier Excel", type=["xlsx"])

if fichier_excel:
    try:
        # 🔹 Charger le fichier Excel
        xls = pd.ExcelFile(fichier_excel, engine="openpyxl")
        st.success(f"✅ Fichier chargé ! {len(xls.sheet_names)} feuille(s) détectée(s) : {', '.join(xls.sheet_names)}")

        if st.button("🚀 Convertir toutes les feuilles en CSV"):
            for sheet_name in xls.sheet_names:
                # --- Lecture de la feuille courante ---
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None, engine="openpyxl")

                # --- Contrôle du volume de lignes ---
                if df.shape[0] <= 6:
                    st.warning(f"⚠️ Feuille « {sheet_name} » ignorée : moins de 7 lignes.")
                    continue

                # --- Extraction des données ---
                url = df.iloc[2, 0]                                   # A3
                mot_cle_principal = f"{df.iloc[4, 0]} {df.iloc[4, 1]}".strip()  # A5 + B5
                mots_cles = df.iloc[6:, 0].dropna().tolist()          # À partir de A7
                frequence_conseillee = df.iloc[6:, 1].dropna().tolist()  # B7+

                # Supprimer la possible ligne d’en-tête « Mots-clés »
                mots_cles = [m for m in mots_cles if m.lower().strip() != "mots-clés"]

                # --- Transformation pour export ---
                final_df = pd.DataFrame({
                    "URL": [url],
                    "Mot Clé Principal": [mot_cle_principal],
                    "Keywords": ["|".join(mots_cles)],
                    "Fréquence conseillée": ["|".join(map(str, frequence_conseillee))]
                })

                # --- Affichage Streamlit ---
                st.subheader(f"📑 Feuille : {sheet_name}")
                st.dataframe(final_df, use_container_width=True)

                # --- Préparation du CSV en mémoire ---
                buffer = BytesIO()
                final_df.to_csv(buffer, index=False, sep=",", encoding="utf-8-sig")
                buffer.seek(0)

                # --- Bouton de téléchargement ---
                st.download_button(
                    label=f"📥 Télécharger le CSV de « {sheet_name} »",
                    data=buffer,
                    file_name=f"{sheet_name}.csv",
                    mime="text/csv"
                )

        st.info("ℹ️ Chaque bouton produit un fichier CSV distinct, nommé comme la feuille d'origine.")
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
