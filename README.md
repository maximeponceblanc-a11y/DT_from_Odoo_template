# Générateur de Dossier Technique

Remplace la macro VBA `cmdCreerDossierFabrication_Click` : lit un devis au
**nouveau format** (feuille « Bilan ») et génère automatiquement le
**Dossier Technique** Excel conforme au format attendu par le logiciel tiers.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `dossier_technique.py` | Le script Python (bibliothèque + ligne de commande) |
| `app_streamlit.py` | Interface glisser-déposer (optionnelle) |
| `template_dossier_technique.xlsx` | Le modèle Excel cible (mise en forme du Dossier Technique) |
| `requirements.txt` | Dépendances Python |

## 1. Installation (une seule fois)

Il faut Python 3.10+ installé sur le poste.

```bash
cd dossier_technique_generator
pip install -r requirements.txt
```

## 2. Utilisation en ligne de commande

```bash
python dossier_technique.py "mon_devis.xlsx"
```

Cela crée automatiquement `mon_devis_DT.xlsx` à côté du devis.

Options :

```bash
# Choisir le nom/emplacement du fichier de sortie
python dossier_technique.py "mon_devis.xlsx" -o "Dossier_Technique.xlsx"

# Utiliser un autre modèle que celui fourni par défaut
python dossier_technique.py "mon_devis.xlsx" -t "autre_modele.xlsx"
```

En cas de devis trop volumineux pour tenir dans le modèle (voir limites
ci-dessous), le script s'arrête avec un message d'erreur clair plutôt que de
produire un fichier incorrect.

## 3. Utilisation avec l'interface Streamlit (glisser-déposer)

```bash
streamlit run app_streamlit.py
```

Une page s'ouvre dans le navigateur : déposer le devis, vérifier l'aperçu des
données détectées (client, achats, temps), cliquer sur **Générer**, puis
télécharger le Dossier Technique produit.

## 4. Ce que fait le script, précisément

**En-tête**
- N° devis / N° dossier : laissés vides
- Client ← `Bilan!B3`, Contact ← `Inconnu`, Téléphone ← `Inconnu`
- Référence ← `Bilan!B4`, Quantité commandée ← `Bilan!B5`

**Achats** (`Bilan` lignes 29-58, en-têtes ligne 28)
- Une catégorie *Achats \<Section\>* est créée automatiquement pour chaque
  valeur distincte trouvée dans la colonne A (SECTION), dans l'ordre
  d'apparition dans le devis.
- Pour chaque ligne : Article = SECTION, Support = COMPOSANT, Référence =
  CODE, Prix unitaire = P VENTE UNITAIRE, Quantité = QUANTITÉ, Coût devisé =
  PRIX TOTAL. Surface et Fournisseur sont laissés vides (non fournis par le
  nouveau devis).

**Temps de fabrication** (`Bilan` lignes 64-93, en-têtes ligne 63)
- Une catégorie *Temps \<Section\>* par valeur distincte de la colonne A.
- Article = SECTION, Code = Code opération, Nom opération = Opération,
  Tps = Temps total opérateur.

## 5. Hypothèses retenues (à valider / ajuster si besoin)

Le cahier des charges ne précisait pas où positionner les catégories
dynamiques dans le fichier de sortie (l'ancien système avait 6 catégories
d'achats et 5 catégories de temps à des lignes fixes). Le modèle fourni
(`exemple fichier export.xlsx`) est en réalité une mise en page imprimable
avec des sauts de page tous les 62 lignes ; le script a été construit pour
reproduire fidèlement cette logique de pagination (identique à celle de
l'ancienne macro) mais en l'appliquant à un nombre **variable** de
catégories :

- Les blocs achats/temps sont désormais tous au **même format** (celui du
  bloc "Achats des conditionnements" de l'exemple, qui contenait déjà les 8
  champs demandés : Article, Support, Référence, Prix unitaire, Surface,
  Quantité, Fournisseur, Coût devisé).
- Une colonne "Unité" est affichée à côté du prix unitaire (ex. `€/pièce`) à
  titre indicatif — désactivable via `INCLUDE_UNIT_LABEL = False` en haut de
  `dossier_technique.py`.
- Les champs qui ne figurent pas dans le mapping demandé (Descriptif,
  Observations, Livraison, Pré-presse/BAT, Départ/Franco/Catégorie, Synthèse
  des coûts et des temps...) sont **laissés vides** plutôt que de conserver
  les valeurs de l'ancien exemple — pour éviter qu'un dossier généré affiche
  par erreur les informations d'un autre devis. Ils restent modifiables à la
  main dans le fichier généré, ou pourront être ajoutés au script plus tard
  si une source de données est identifiée pour eux.

**Capacité maximale.** Le modèle prévoit 2 pages pour les achats et 2 pages
pour les temps (comme l'ancien système). Cela permet dans la plupart des cas
courants d'aller jusqu'au maximum annoncé de 30 lignes d'achats et 30 lignes
d'opérations. Si un devis est réparti sur de très nombreuses catégories
différentes (chaque catégorie ajoutant un titre + une mise en forme), la
capacité peut être atteinte plus tôt ; dans ce cas le script s'arrête avec un
message d'erreur explicite (plutôt que de produire un fichier corrompu) en
indiquant de réduire le nombre de catégories ou de lignes. Si cette limite
est atteinte régulièrement en usage réel, il faudra étendre le modèle Excel
avec des pages supplémentaires (extension possible, mais non incluse pour
l'instant — l'insertion fiable de nouvelles pages dans un classeur déjà mis
en forme est plus risquée qu'utile tant que le besoin n'est pas confirmé).

## 6. Fichiers corrompus / data validation invalide

Certains exports (dont `exemple nouveau devis.xlsx` fourni) contiennent une
data validation Excel légèrement invalide qui fait planter la bibliothèque
`openpyxl` par défaut. Le script détecte et corrige automatiquement ce cas
(`repair_xlsx`) avant de relire le fichier — aucune action nécessaire.

## 7. Personnalisation

Tous les emplacements de cellules (devis source et modèle cible) sont définis
comme des constantes en haut de `dossier_technique.py` (`DEVIS_CLIENT_CELL`,
`CELL_CLIENT`, `ACHATS_FIRST_ROW`, etc.). Si la structure du devis ou du
modèle change légèrement, il suffit d'ajuster ces constantes plutôt que de
réécrire la logique.
