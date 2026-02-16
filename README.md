# SNIS RDC - Dashboard Performance

Application Streamlit pour analyser la performance SNIS (DHIS2) avec:
- complétude
- promptitude
- comparaison entre zones/aires de santé
- analyse des violations de règles de validation
- export commenté en **Excel / Word / PowerPoint**

Le script principal est: `dashbord.py`.

## 1. Fonctionnalités

Le dashboard est organisé en 5 onglets:

1. **Base de données**
- Affichage des données brutes (sans `Organisation unit ID`)
- Filtrage des lignes parent pour afficher les aires de santé

2. **Complétude**
- Calcul `Reports_Actual`, `Reports_Attendu`, `Complétude_Globale (%)`
- Mise en forme couleur selon seuils
- Graphique de classement des zones

3. **Promptitude**
- Calcul `Promptitude_Globale (%)`
- Score du nombre de datasets avec promptitude `>= 95%`

4. **Analyse Comparative**
- Tableau comparatif complétude/promptitude
- Quadrant de performance
- Top 5 complétude / Flop 5 promptitude
- Tableau fusionné (zone filtrée): indicateurs dataset reporting + actual

5. **Éléments de catégorisation**
- Violations de règles par zone de santé
- Colonnes M-1 / M:
  - `Règles violées (M-1)`
  - `Règles corrigées (M-1 -> M)`
  - `Règles violées (M)`
- `Ratio / 100 rapports` calculé par:
  - `(Règles violées (M) / Reports_Actual) * 100`
- `Score de qualité`

## 2. Export de rapport

Depuis la **sidebar**:
- Choix du type de téléchargement: `Excel`, `Word`, `PowerPoint`
- Bouton de téléchargement dynamique selon le type choisi

Le rapport exporté inclut:
- tableaux principaux
- commentaires automatiques
- (PowerPoint) images des graphiques et tableaux

## 3. Prérequis

- Python 3.10+ recommandé
- Accès DHIS2 valide (URL, utilisateur, mot de passe)

## 4. Installation

Depuis le dossier `Extraction_Favori`:

```powershell
python -m pip install -r requirements.txt
```

## 5. Configuration DHIS2

Créer le fichier `Extraction_Favori/.streamlit/secrets.toml`:

```toml
DHIS2_URL = "https://votre-instance-dhis2"
DHIS2_USER = "votre_utilisateur"
DHIS2_PASS = "votre_mot_de_passe"
```

Le projet lit ces clés via `st.secrets`.

## 6. Lancement

```powershell
streamlit run dashbord.py
```

Ensuite:
1. Choisir le favori DHIS2 (ou un ID personnalisé)
2. Choisir la période (année + mois début/fin)
3. Filtrer par zone/aire de santé
4. Consulter les onglets et exporter le rapport

## 7. Structure du dossier

```text
Extraction_Favori/
|- dashbord.py
|- requirements.txt
|- .streamlit/
|  |- secrets.toml
|- .gitignore
```

## 8. Dépannage rapide

- **PowerPoint ne génère pas les images**
  - vérifier `kaleido` et `matplotlib` dans l’environnement

- **Erreur de connexion DHIS2**
  - vérifier `DHIS2_URL`, `DHIS2_USER`, `DHIS2_PASS` dans `secrets.toml`

- **Pas de données**
  - vérifier l’ID du favori DHIS2
  - vérifier la période sélectionnée
  - vérifier les droits utilisateur DHIS2

## 9. Dépendances utilisées

`requirements.txt`:
- `streamlit`
- `pandas`
- `numpy`
- `plotly`
- `openhexa.toolbox`
- `xlsxwriter`
- `python-docx`
- `python-pptx`
- `matplotlib`
- `kaleido`
