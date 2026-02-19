# SNIS RDC - Dashboard Performance

Application Streamlit pour analyser la performance SNIS (DHIS2) avec:
- complétude
- promptitude
- comparaison entre zones/aires de santé
- filtrage hiérarchique `Synthèse pays` / `Province` / `Zone de santé`
- analyse des violations de règles de validation
- export commenté en **Excel / PowerPoint**

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
- Choix du type de téléchargement: `Excel`, `PowerPoint`
- Bouton de téléchargement dynamique selon le type choisi
- Bouton `Visualiser le rapport` pour consulter l'aperçu complet sans télécharger

Le rapport exporté inclut:
- tableaux principaux
- commentaires automatiques (lecture des graphiques et tableaux)
- conservation des colorations conditionnelles du dashboard dans les tableaux exportés
- (PowerPoint) images des graphiques et tableaux
- feuilles/slides dédiées: base de données, rapports détaillés (réels/attendus), performance finale, promptitude, comparatif, top/flop, résultats des règles

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
# Optionnel (connexion lente):
# DHIS2_TIMEOUT_CONNECT = 10
# DHIS2_TIMEOUT_READ = 120
# DHIS2_HTTP_RETRIES = 2
```

Le projet lit `DHIS2_URL` via `st.secrets`.
Le **nom d'utilisateur** et le **mot de passe** sont saisis par chaque utilisateur dans la barre latérale (`🔐 Connexion DHIS2`).

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
  - vérifier `DHIS2_URL` dans `secrets.toml`
  - vérifier les identifiants saisis dans `🔐 Connexion DHIS2`

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
- `requests`
- `xlsxwriter`
- `python-docx`
- `python-pptx`
- `matplotlib`
- `kaleido`
