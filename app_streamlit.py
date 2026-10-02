"""
app_streamlit.py
=================

Petite interface Streamlit pour générer un Dossier Technique à partir d'un
devis, sans utiliser la ligne de commande : glisser-déposer le devis,
cliquer sur "Générer", télécharger le résultat.

Lancement :
    pip install -r requirements.txt
    streamlit run app_streamlit.py
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from dossier_technique import generate_dossier_technique, read_devis

st.set_page_config(page_title="Générateur de Dossier Technique", page_icon="📄", layout="centered")

st.title("📄 Générateur de Dossier Technique")
st.write(
    "Dépose ton fichier de **devis** (nouveau format, feuille « Bilan ») pour "
    "générer automatiquement le Dossier Technique à importer dans le logiciel tiers."
)

DEFAULT_TEMPLATE = Path(__file__).parent / "template_dossier_technique.xlsx"

with st.expander("⚙️ Options avancées"):
    st.caption(
        "Par défaut, le script utilise le modèle `template_dossier_technique.xlsx` "
        "fourni avec l'application. Tu peux le remplacer par un autre modèle si le "
        "format du Dossier Technique évolue."
    )
    custom_template = st.file_uploader(
        "Modèle de Dossier Technique personnalisé (optionnel, .xlsx)", type=["xlsx"], key="template"
    )

devis_file = st.file_uploader("Devis à traiter (.xlsx)", type=["xlsx", "xlsm"], key="devis")

if devis_file is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        devis_path = tmpdir / devis_file.name
        devis_path.write_bytes(devis_file.getvalue())

        # Aperçu rapide des données extraites, avant génération
        try:
            preview = read_devis(devis_path)
            st.subheader("Aperçu des données extraites")
            col1, col2, col3 = st.columns(3)
            col1.metric("Client", preview.client or "—")
            col2.metric("Référence", preview.reference or "—")
            col3.metric("Quantité", preview.quantite if preview.quantite is not None else "—")

            st.write("**Catégories d'achats détectées :**")
            if preview.achats:
                st.table(
                    [{"Catégorie": f"Achats {s}", "Lignes": len(v)} for s, v in preview.achats.items()]
                )
            else:
                st.info("Aucune ligne d'achat détectée (lignes 29-58 de la feuille Bilan).")

            st.write("**Catégories de temps de fabrication détectées :**")
            if preview.temps:
                st.table(
                    [{"Catégorie": f"Temps {s}", "Lignes": len(v)} for s, v in preview.temps.items()]
                )
            else:
                st.info("Aucune ligne de temps détectée (lignes 64-93 de la feuille Bilan).")
        except Exception as exc:  # noqa: BLE001
            st.error(f"Impossible de lire ce devis : {exc}")
            st.stop()

        if st.button("🚀 Générer le Dossier Technique", type="primary"):
            if custom_template is not None:
                template_path = tmpdir / custom_template.name
                template_path.write_bytes(custom_template.getvalue())
            else:
                if not DEFAULT_TEMPLATE.exists():
                    st.error(
                        "Le modèle par défaut `template_dossier_technique.xlsx` est introuvable "
                        "à côté de ce script. Fournis un modèle personnalisé ci-dessus."
                    )
                    st.stop()
                template_path = DEFAULT_TEMPLATE

            output_path = tmpdir / f"{devis_path.stem}_DT.xlsx"
            try:
                with st.spinner("Génération en cours…"):
                    generate_dossier_technique(devis_path, template_path, output_path)
            except OverflowError as exc:
                st.error(f"⚠️ {exc}")
                st.stop()
            except Exception as exc:  # noqa: BLE001
                st.error(f"Erreur lors de la génération : {exc}")
                st.stop()

            st.success("Dossier Technique généré avec succès !")
            st.download_button(
                "⬇️ Télécharger le Dossier Technique",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
else:
    st.info("En attente d'un fichier de devis…")
